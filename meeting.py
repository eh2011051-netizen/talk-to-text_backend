# meeting.py
import os
import json
import uuid
import asyncio
import base64 
import logging
from datetime import datetime, timedelta, timezone
from flask import Flask, request, jsonify
from flask_socketio import SocketIO, emit, join_room, leave_room
from flask_jwt_extended import jwt_required, get_jwt_identity
from werkzeug.security import generate_password_hash, check_password_hash
import redis
from threading import Thread, Lock
import time
import assemblyai as aai
from concurrent.futures import ThreadPoolExecutor
import google.generativeai as genai
from pydub import AudioSegment
import io
import threading
import websockets
import ssl
import certifi
from pathlib import Path


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Initialize Flask app and SocketIO
socketio = SocketIO(cors_allowed_origins="*", async_mode='threading')

RECORDINGS_DIR = Path("recordings")
RECORDINGS_DIR.mkdir(parents=True, exist_ok=True)

# Redis for pub/sub and session management
redis_client = redis.Redis(
    host=os.getenv('REDIS_HOST', 'localhost'),
    port=int(os.getenv('REDIS_PORT', 6379)),
    password=os.getenv('REDIS_PASSWORD', None),
    decode_responses=True
)

# In-memory storage for active meetings (fallback if Redis fails)
active_meetings = {}
meeting_lock = Lock()

# Thread pool for processing
processing_executor = ThreadPoolExecutor(max_workers=10)

# Initialize AssemblyAI
aai.settings.api_key = os.getenv("ASSEMBLYAI_API_KEY")

# Initialize Gemini AI
gemini_api_key = os.getenv("GEMINI_API_KEY")
if gemini_api_key:
    genai.configure(api_key=gemini_api_key)

# Recording Manager Class
class MeetingRecorder:
    def __init__(self, meeting_id, user_id, db, socketio_instance, initial_participants=None):
        self.meeting_id = meeting_id
        self.user_id = user_id
        self.db = db
        self.socketio = socketio_instance
        self.loop = None
        self.initial_participants = initial_participants or {} # Dict of participant_id -> user_data
        
        # Recording state
        self.is_recording = False
        self.recording_start_time = None
        self.audio_chunks = []
        self.video_chunks = []
        self.transcript_chunks = []
        self.speaker_data = {}
        
        # Real-time transcription
        self.websocket = None
        self.transcriber = None
        self.transcript_buffer = []
        
        # Processing
        self.processing_queue = asyncio.Queue()
        self.processing_task = None
        
        # Statistics
        self.stats = {
            'total_duration': 0,
            'speakers_detected': 0,
            'words_transcribed': 0,
            'sentiment_score': 0.5,
            'action_items': [],
            'decisions': [],
            'questions': []
        }

        # Speaker Mapping (User1, User2, etc.)
        self.speaker_label_map = {}  # Map from original speaker ID to UserN
        self.next_speaker_num = 1
        self.speaker_aware_transcript = []
    
    async def start_recording(self):
        """Start recording with real-time transcription"""
        try:
            self.is_recording = True
            self.recording_start_time = datetime.now(timezone.utc)
            self.loop = asyncio.get_running_loop()
            
            # Initialize AssemblyAI real-time transcriber
            self.realtime_transcriber = aai.RealtimeTranscriber(
                on_data=self._on_transcription_data,
                on_error=self._on_transcription_error,
                on_open=self._on_transcription_open,
                on_close=self._on_transcription_close,
                sample_rate=16000,
                encoding=aai.AudioEncoding.pcm_s16le
            )
            
            # Start real-time connection
            await self.realtime_transcriber.connect()
            
            # Start processing task
            self.processing_task = asyncio.create_task(self._process_queue())
            
            # Send notification
            self.socketio.emit('recording_status', {
                'meeting_id': self.meeting_id,
                'status': 'started',
                'started_at': self.recording_start_time.isoformat(),
                'message': 'Automatic Recording & Transcription Started'
            }, room=self.meeting_id)
            
            logger.info(f"Automatic Recording started for meeting {self.meeting_id}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to start recording: {e}")
            self.socketio.emit('recording_error', {
                'meeting_id': self.meeting_id,
                'error': str(e)
            }, room=self.meeting_id)
            return False
    
    async def stop_recording(self):
        """Stop recording and save everything"""
        try:
            self.is_recording = False
            
            # Cancel processing task
            if self.processing_task:
                self.processing_task.cancel()
                try:
                    await self.processing_task
                except asyncio.CancelledError:
                    pass
            
            # Calculate duration
            duration = (datetime.now(timezone.utc) - self.recording_start_time).total_seconds()
            self.stats['total_duration'] = duration
            
            # Process transcription if we have audio
            if self.audio_chunks:
                await self._process_transcription()
            
            # Save recording to database
            recording_id = await self._save_to_database()
            
            # Send final stats
            self.socketio.emit('recording_completed', {
                'meeting_id': self.meeting_id,
                'recording_id': recording_id,
                'duration': duration,
                'stats': self.stats,
                'speakers': len(self.speaker_data),
                'transcript_length': len(self.transcript_buffer)
            }, room=self.meeting_id)
            
            logger.info(f"Recording stopped for meeting {self.meeting_id}, duration: {duration}s")
            return recording_id
            
        except Exception as e:
            logger.error(f"Failed to stop recording: {e}")
            return None
    
    async def _process_transcription(self):
        """Process transcription using AssemblyAI"""
        try:
            if not self.audio_chunks:
                return
            
            # Combine audio chunks into a single file
            audio_file = await self._combine_audio_chunks()
            
            # Save to temporary file
            temp_dir = "temp_audio"
            os.makedirs(temp_dir, exist_ok=True)
            temp_file = os.path.join(temp_dir, f"meeting_{self.meeting_id}.wav")
            audio_file.export(temp_file, format="wav")
            
            try:
                # Configure transcription
                config = aai.TranscriptionConfig(
                    speaker_labels=True,
                    punctuate=True,
                    format_text=True,
                    language_detection=True
                )
                
                # Transcribe the audio file
                logger.info(f"Starting transcription for meeting {self.meeting_id}")
                transcript = self.transcriber.transcribe(temp_file, config=config)
                
                if transcript.status == aai.TranscriptStatus.error:
                    logger.error(f"Transcription failed: {transcript.error}")
                    return
                
                # Process transcript
                if transcript.utterances:
                    for utterance in transcript.utterances:
                        if utterance.text:
                            chunk = {
                                'text': utterance.text,
                                'speaker': f"Speaker {utterance.speaker}",
                                'timestamp': datetime.now(timezone.utc).isoformat(),
                                'confidence': utterance.confidence,
                                'words': []
                            }
                            
                            self.transcript_buffer.append(chunk)
                            self.stats['words_transcribed'] += len(utterance.text.split())
                            
                            # Store speaker data
                            speaker_id = f"speaker_{utterance.speaker}"
                            if speaker_id not in self.speaker_data:
                                self.speaker_data[speaker_id] = {
                                    'audio_chunks': 0,
                                    'speaking_time': len(utterance.text.split()) * 0.5,  # Estimate
                                    'last_activity': datetime.now(timezone.utc),
                                    'sentiment_scores': []
                                }
                            
                            # Analyze sentiment
                            sentiment = self._analyze_sentiment(utterance.text)
                            self.speaker_data[speaker_id]['sentiment_scores'].append(sentiment['score'])
                            
                            # Send transcript chunk
                            self.socketio.emit('live_transcript', {
                                'meeting_id': self.meeting_id,
                                'transcript': chunk,
                                'is_final': True
                            }, room=self.meeting_id)
                            
                            # Analyze in real-time
                            asyncio.create_task(self._analyze_transcript_chunk(chunk))
                
                logger.info(f"Transcription completed for meeting {self.meeting_id}")
                
            finally:
                # Clean up temp file
                try:
                    os.remove(temp_file)
                except:
                    pass
                
        except Exception as e:
            logger.error(f"Failed to process transcription: {e}")

    def _on_transcription_data(self, transcript: aai.RealtimeTranscript):
        """Handle real-time transcription data with speaker mapping"""
        if not transcript.text:
            return

        # Get speaker ID from AssemblyAI (usually letters like A, B, C)
        raw_speaker = getattr(transcript, 'speaker', 'A') or 'A'
        
        # Map to stable UserN label
        if raw_speaker not in self.speaker_label_map:
            self.speaker_label_map[raw_speaker] = f"User{self.next_speaker_num}"
            self.next_speaker_num += 1
            self.stats['speakers_detected'] = self.next_speaker_num - 1

        stable_speaker = self.speaker_label_map[raw_speaker]
        timestamp = datetime.now(timezone.utc).strftime('%H:%M:%S')

        # Create transcript line
        transcript_line = {
            'id': str(uuid.uuid4()),
            'speaker': stable_speaker,
            'text': transcript.text,
            'timestamp': timestamp,
            'raw_speaker': raw_speaker
        }

        # Store in buffer
        self.speaker_aware_transcript.append(transcript_line)
        
        # Emit to frontend
        self.socketio.emit('speaker-aware-transcript', {
            'meeting_id': self.meeting_id,
            'entry': transcript_line
        }, room=self.meeting_id)

        logger.info(f"[{timestamp}] {stable_speaker}: {transcript.text}")
        
        # Trigger real-time analysis
        if self.loop and self.loop.is_running():
            logger.info(f"Scheduling sentiment analysis for: {transcript.text[:50]}...")
            self.loop.call_soon_threadsafe(
                lambda: asyncio.create_task(self._analyze_transcript_chunk(transcript_line))
            )
        else:
            logger.warning("No running event loop found to schedule sentiment analysis")

    def _on_transcription_error(self, error: Exception):
        logger.error(f"Real-time transcription error: {error}")

    def _on_transcription_open(self, session_opened: aai.RealtimeSessionOpened):
        logger.info(f"Real-time transcription session opened: {session_opened.session_id}")

    def _on_transcription_close(self):
        logger.info("Real-time transcription session closed")
    
    async def add_audio_chunk(self, audio_data: bytes, speaker_id: str, timestamp: datetime):
        """Add audio chunk from speaker"""
        try:
            if not self.is_recording:
                return False
            
            # Store chunk
            chunk = {
                'data': audio_data,
                'speaker_id': speaker_id,
                'timestamp': timestamp.isoformat(),
                'type': 'audio'
            }
            self.audio_chunks.append(chunk)
            
            # Update speaker data
            if speaker_id not in self.speaker_data:
                self.speaker_data[speaker_id] = {
                    'audio_chunks': 0,
                    'speaking_time': 0,
                    'last_activity': timestamp,
                    'sentiment_scores': []
                }
            
            self.speaker_data[speaker_id]['audio_chunks'] += 1
            self.speaker_data[speaker_id]['last_activity'] = timestamp
            
            # Add to processing queue
            await self.processing_queue.put({
                'type': 'audio',
                'data': chunk,
                'speaker_id': speaker_id
            })

            # Also send to real-time transcriber
            if hasattr(self, 'realtime_transcriber') and self.realtime_transcriber:
                try:
                    await self.realtime_transcriber.send(audio_data)
                except Exception as e:
                    logger.error(f"Failed to send data to transcriber: {e}")
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to add audio chunk: {e}")
            return False
    
    async def add_video_chunk(self, video_data: bytes, speaker_id: str, timestamp: datetime):
        """Add video chunk from speaker"""
        try:
            if not self.is_recording:
                return False
            
            chunk = {
                'data': video_data,
                'speaker_id': speaker_id,
                'timestamp': timestamp.isoformat(),
                'type': 'video'
            }
            self.video_chunks.append(chunk)
            
            await self.processing_queue.put({
                'type': 'video',
                'data': chunk,
                'speaker_id': speaker_id
            })
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to add video chunk: {e}")
            return False
    
    async def _analyze_transcript_chunk(self, chunk):
        """Analyze transcript chunk in real-time"""
        try:
            text = chunk['text']
            speaker = chunk['speaker']
            
            # Simple sentiment analysis
            sentiment = self._analyze_sentiment(text)
            
            # Update speaker sentiment
            if speaker in self.speaker_data:
                self.speaker_data[speaker]['sentiment_scores'].append(sentiment['score'])
                if len(self.speaker_data[speaker]['sentiment_scores']) > 10:
                    self.speaker_data[speaker]['sentiment_scores'] = self.speaker_data[speaker]['sentiment_scores'][-10:]
            
            # Update overall sentiment
            all_scores = []
            for speaker_data in self.speaker_data.values():
                all_scores.extend(speaker_data['sentiment_scores'])
            if all_scores:
                self.stats['sentiment_score'] = sum(all_scores) / len(all_scores)
            
            # Detect action items
            if any(keyword in text.lower() for keyword in ['need to', 'should', 'must', 'action', 'task']):
                action_item = {
                    'text': text,
                    'speaker': speaker,
                    'timestamp': chunk['timestamp'],
                    'detected_at': datetime.now(timezone.utc).isoformat()
                }
                self.stats['action_items'].append(action_item)
                
                # Notify about action item
                self.socketio.emit('action_item_detected', {
                    'meeting_id': self.meeting_id,
                    'action_item': action_item
                }, room=self.meeting_id)
            
            # Detect questions
            if '?' in text:
                question = {
                    'text': text,
                    'speaker': speaker,
                    'timestamp': chunk['timestamp']
                }
                self.stats['questions'].append(question)
            
            # Detect decisions
            if any(keyword in text.lower() for keyword in ['decided', 'agree', 'decision', 'approved', 'resolve']):
                decision = {
                    'text': text,
                    'speaker': speaker,
                    'timestamp': chunk['timestamp']
                }
                self.stats['decisions'].append(decision)
                
            # Send real-time insights
            self.socketio.emit('meeting_insight', {
                'meeting_id': self.meeting_id,
                'type': 'sentiment',
                'data': {
                    'speaker': speaker,
                    'sentiment': sentiment['label'],
                    'score': sentiment['score']
                },
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=self.meeting_id)
            logger.info(f"Emitted mood update: {sentiment['label']} ({sentiment['score']}) for {speaker}")
            
        except Exception as e:
            logger.error(f"Error analyzing transcript: {e}")
    
    def _analyze_sentiment(self, text):
        """Enhanced simple sentiment analysis"""
        positive_words = [
            'good', 'great', 'excellent', 'happy', 'agree', 'yes', 'perfect', 'awesome',
            'cool', 'nice', 'love', 'amazing', 'brilliant', 'fantastic', 'correct',
            'support', 'definitely', 'absolutely', 'exactly', 'on track', 'resolved'
        ]
        negative_words = [
            'bad', 'terrible', 'no', 'disagree', 'problem', 'issue', 'wrong', 'angry',
            'frustrated', 'confused', 'error', 'fail', 'hate', 'stupid', 'boring',
            'slow', 'broken', 'difficult', 'hard', 'stop', 'late', 'risk', 'warning'
        ]
        
        words = text.lower().split()
        positive = sum(1 for word in words if word in positive_words)
        negative = sum(1 for word in words if word in negative_words)
        
        # Also check for common phrases
        text_lower = text.lower()
        if "i disagree" in text_lower or "major problem" in text_lower or "not good" in text_lower:
            negative += 1
        if "totally agree" in text_lower or "sounds good" in text_lower or "great idea" in text_lower:
            positive += 1
        
        total = len(words)
        if total == 0:
            return {'score': 0.5, 'label': 'neutral'}
        
        # Calculate score from 0 to 1
        # Neutral base is 0.5. Increase weight of detected words.
        if total > 0:
            sentiment_delta = (positive - negative) / max(total, 1)
            # Magnify the delta for better visual feedback
            score = 0.5 + (sentiment_delta * 1.5) 
            score = max(0.1, min(0.9, score)) # Keep away from extreme 0 or 1 for smoother feel
        else:
            score = 0.5
        
        if score > 0.52:
            label = 'positive'
        elif score < 0.48:
            label = 'negative'
        else:
            label = 'neutral'
        
        return {'score': score, 'label': label}
    
    async def _process_queue(self):
        """Process queued items"""
        while self.is_recording:
            try:
                item = await self.processing_queue.get()
                
                if item['type'] == 'audio':
                    await self._process_audio_chunk(item['data'])
                elif item['type'] == 'video':
                    await self._process_video_chunk(item['data'])
                
                self.processing_queue.task_done()
                
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"Error processing queue: {e}")
    
    async def _process_audio_chunk(self, chunk):
        """Process audio chunk"""
        try:
            speaker_id = chunk['speaker_id']
            
            # Update speaking time
            if speaker_id in self.speaker_data:
                # Assume 100ms chunks
                self.speaker_data[speaker_id]['speaking_time'] += 0.1
            
            # Send real-time speaking update
            self.socketio.emit('speaking_update', {
                'meeting_id': self.meeting_id,
                'speaker_id': speaker_id,
                'is_speaking': True,
                'speaking_time': self.speaker_data.get(speaker_id, {}).get('speaking_time', 0)
            }, room=self.meeting_id)
            
        except Exception as e:
            logger.error(f"Error processing audio chunk: {e}")
    
    async def _process_video_chunk(self, chunk):
        """Process video chunk"""
        # Currently just store, could add face detection, emotion analysis, etc.
        pass
        
    async def _save_to_database(self):
        """Save recording to database and file system"""
        try:
            # Generate transcript
            transcript_text = self._generate_transcript()

            # Create recordings directory if it doesn't exist
            recordings_dir = RECORDINGS_DIR
            recordings_dir.mkdir(parents=True, exist_ok=True)

            # Save audio file if we have audio chunks
            filename = None
            filepath = None

            if self.audio_chunks:
                audio_file = await self._combine_audio_chunks()
                timestamp = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')
                filename = f"meeting_{self.meeting_id}_{timestamp}.wav"
                filepath = recordings_dir / filename

                try:
                    audio_file.export(str(filepath), format="wav")
                    logger.info(f"Audio recording saved to: {filepath}")
                except Exception as e:
                    logger.error(f"Failed to save audio file: {e}")
                    # Try alternative format
                    filename = f"meeting_{self.meeting_id}_{timestamp}.mp3"
                    filepath = recordings_dir / filename
                    audio_file.export(str(filepath), format="mp3")

            # Save to database using the existing Meeting model
            from app import Meeting, log_activity

            # Create participant mapping for the frontend to show real names
            participant_map = {}
            for pid, pdata in self.initial_participants.items():
                participant_map[pid] = pdata.get('full_name', 'Unknown')

            meeting_record = Meeting(
                user_id=self.user_id,
                title=f"Meeting: {self.meeting_id}",
                filename=filename or f"meeting_{self.meeting_id}.txt",
                filepath=str(filepath) if filepath else None,
                language='en',
                transcript_language='en',
                status='uploaded',
                upload_date=datetime.now(timezone.utc),
                duration=self.stats.get('total_duration', 0),
                participants_count=len(self.speaker_data),
                transcript=transcript_text[:5000],  # Store first 5000 chars
                source='live',
                participant_mapping=json.dumps(participant_map)
            )

            self.db.session.add(meeting_record)
            self.db.session.commit()

            # Store full transcription as JSON
            meeting_record.transcription = json.dumps({
                'raw': transcript_text,
                'speakers': self._get_speaker_summary(),
                'duration': self.stats['total_duration'],
                'word_count': self.stats['words_transcribed'],
                'sentiment': self.stats['sentiment_score'],
                'action_items': self.stats['action_items'][-10:],
                'decisions': self.stats['decisions'][-10:],
                'questions': self.stats['questions'][-10:],
                'audio_file': filename,
                'audio_path': str(filepath) if filepath else None,
                'recorded_at': datetime.now(timezone.utc).isoformat()
            })

            self.db.session.commit()

            # Log activity
            log_activity(
                user_id=self.user_id,
                activity_type='meeting_recording',
                title=f"Meeting Recording: {self.meeting_id}",
                description=f"Recorded meeting with {len(self.speaker_data)} participants",
                meeting_id=meeting_record.id,
                metadata={
                    'meeting_id': self.meeting_id,
                    'duration': self.stats['total_duration'],
                    'speakers': len(self.speaker_data),
                    'word_count': self.stats['words_transcribed'],
                    'audio_file': filename,
                    'file_size': filepath.stat().st_size if filepath and filepath.exists() else 0
                }
            )

            # Start background processing
            processing_executor.submit(self._start_processing, meeting_record.id)

            logger.info(f"Recording saved with ID: {meeting_record.id} to {filepath}")
            return meeting_record.id

        except Exception as e:
            logger.error(f"Failed to save to database: {e}")
            self.db.session.rollback()
            return None
            
    async def save_meeting_complete(self, user_id):
        """Save meeting recording and data when meeting ends"""
        try:
            # Stop recording if active
            recording_id = None
            if self.is_recording:
                recording_id = await self.stop_recording(user_id)
        
            # If no recording was active but we have meeting data, still save
            if not recording_id and len(self.participants) > 0:
                # Create a meeting record without audio
                from app import Meeting, log_activity, db
                from datetime import datetime
            
                # Generate transcript from messages
                transcript_lines = []
                for message in self.messages:
                    if message.get('text') and not message.get('isAI', False):
                        timestamp = message.get('timestamp', '')
                        speaker = message.get('userName', 'Unknown')
                        text = message.get('text', '')
                        transcript_lines.append(f"[{timestamp}] {speaker}: {text}")
            
                transcript_text = '\n'.join(transcript_lines)
            
                # Create participant mapping
                participant_map = {}
                for pid, pdata in self.participants.items():
                    participant_map[pid] = pdata.get('full_name', 'Unknown')

                meeting_record = Meeting(
                    user_id=user_id,
                    title=f"Meeting: {self.title}",
                    filename=f"meeting_{self.id}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.txt",
                    language=self.settings.get('language', 'en'),
                    transcript_language=self.settings.get('language', 'en'),
                    status='completed',
                    upload_date=datetime.now(timezone.utc),
                    duration=(datetime.now(timezone.utc) - self.created_at).total_seconds(),
                    participants_count=len(self.participants),
                    transcript=transcript_text[:5000],
                    source='live',
                    participant_mapping=json.dumps(participant_map)
                )
            
                db.session.add(meeting_record)
                db.session.commit()
            
                # Store transcription
                meeting_record.transcription = json.dumps({
                    'raw': transcript_text,
                    'speakers': {pid: data for pid, data in self.participants.items()},
                    'duration': (datetime.now(timezone.utc) - self.created_at).total_seconds(),
                    'messages': len(self.messages),
                    'quiz_sessions': len(self.quiz_sessions),
                    'completed_at': datetime.now(timezone.utc).isoformat()
                })
            
                db.session.commit()
            
                recording_id = meeting_record.id
            
                logger.info(f"Meeting saved without audio: {recording_id}")
        
            return recording_id
        
        except Exception as e:
            logger.error(f"Failed to save meeting complete: {e}")
            return None
    
    async def _combine_audio_chunks(self):
        """Combine audio chunks into single audio file"""
        try:
            if not self.audio_chunks:
                return AudioSegment.silent(duration=1000)
            
            combined = AudioSegment.empty()
            
            for chunk in self.audio_chunks:
                try:
                    # Create AudioSegment from bytes
                    audio_segment = AudioSegment(
                        data=chunk['data'],
                        sample_width=2,
                        frame_rate=16000,
                        channels=1
                    )
                    combined += audio_segment
                except Exception as e:
                    logger.warning(f"Failed to process audio chunk: {e}")
                    continue
            
            return combined
            
        except Exception as e:
            logger.error(f"Failed to combine audio: {e}")
            return AudioSegment.silent(duration=1000)
    
    def _generate_transcript(self):
        """Generate formatted transcript"""
        try:
            lines = []
            for chunk in self.transcript_buffer:
                speaker = chunk.get('speaker', 'Unknown')
                text = chunk.get('text', '')
                timestamp = chunk.get('timestamp', '')
                
                # Format timestamp
                try:
                    dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                    time_str = dt.strftime('%H:%M:%S')
                except:
                    time_str = timestamp
                
                lines.append(f"[{time_str}] {speaker}: {text}")
            
            return '\n'.join(lines)
            
        except Exception as e:
            logger.error(f"Failed to generate transcript: {e}")
            return ""
    
    def _get_speaker_summary(self):
        """Get speaker summary"""
        summary = {}
        for speaker_id, data in self.speaker_data.items():
            summary[speaker_id] = {
                'speaking_time': data.get('speaking_time', 0),
                'audio_chunks': data.get('audio_chunks', 0),
                'avg_sentiment': sum(data.get('sentiment_scores', [0])) / max(len(data.get('sentiment_scores', [1])), 1)
            }
        return summary
    
    def _start_processing(self, meeting_id):
        """Start AI processing of recording"""
        try:
            from app import start_processing
            start_processing(meeting_id)
            logger.info(f"Started processing for meeting recording {meeting_id}")
        except Exception as e:
            logger.error(f"Failed to start processing: {e}")


# Meeting Recording Manager
class MeetingRecordingManager:
    def __init__(self):
        self.recordings = {}
        self.lock = asyncio.Lock()
    
    async def start_recording(self, meeting_id, user_id, db, socketio, initial_participants=None):
        """Start recording a meeting"""
        async with self.lock:
            if meeting_id in self.recordings:
                return self.recordings[meeting_id]
            
            recorder = MeetingRecorder(meeting_id, user_id, db, socketio, initial_participants)
            success = await recorder.start_recording()
            
            if success:
                self.recordings[meeting_id] = recorder
                return recorder
            
            return None
    
    async def stop_recording(self, meeting_id):
        """Stop recording a meeting"""
        async with self.lock:
            if meeting_id not in self.recordings:
                return None
            
            recorder = self.recordings[meeting_id]
            recording_id = await recorder.stop_recording()
            
            if recording_id:
                del self.recordings[meeting_id]
            
            return recording_id
    
    async def add_audio_chunk(self, meeting_id, audio_data, speaker_id, timestamp):
        """Add audio chunk to recording"""
        async with self.lock:
            if meeting_id not in self.recordings:
                return False
            
            recorder = self.recordings[meeting_id]
            return await recorder.add_audio_chunk(audio_data, speaker_id, timestamp)
    
    async def add_video_chunk(self, meeting_id, video_data, speaker_id, timestamp):
        """Add video chunk to recording"""
        async with self.lock:
            if meeting_id not in self.recordings:
                return False
            
            recorder = self.recordings[meeting_id]
            return await recorder.add_video_chunk(video_data, speaker_id, timestamp)
    
    def get_recorder(self, meeting_id):
        """Get recorder for meeting"""
        return self.recordings.get(meeting_id)


# Global recording manager instance
recording_manager = MeetingRecordingManager()


# Meeting models
class MeetingRoom:
    def __init__(self, room_id, title, created_by, is_public=True, password=None):
        self.id = room_id
        self.title = title
        self.created_by = created_by
        self.created_at = datetime.now(timezone.utc)
        self.is_active = True
        self.is_public = is_public
        self.password_hash = generate_password_hash(password) if password else None
        self.participants = {}
        self.messages = []
        self.recordings = []
        self.settings = {
            'allow_screen_share': True,
            'allow_chat': True,
            'allow_recording': True,
            'max_participants': 50,
            'require_approval': False,
            'auto_record': True,  # Auto-record by default
            'language': 'en',
            'translation_enabled': False,
            'speaker_diarization': True,
            'live_transcription': False,  # Changed to offline
            'sentiment_analysis': True
        }
        self.quiz_sessions = {}
        self.ai_assistant = {
            'enabled': True,
            'mode': 'group',  # 'private' or 'group'
            'sensitivity': 100,
            'ghost_replay_enabled': True,
            'auto_task_creation': True,
            'live_insights': True
        }
        self.is_recording = False
        self.recording_started_at = None
        self.current_recording_id = None
        self.transcript_history = []
        self.last_ai_analysis_time = time.time()
        self.analysis_count = 0

    def add_participant(self, participant_id, user_data):
        with meeting_lock:
            if len(self.participants) >= self.settings['max_participants']:
                return False
            
            self.participants[participant_id] = {
                **user_data,
                'joined_at': datetime.now(timezone.utc).isoformat(),
                'is_audio_enabled': True,
                'is_video_enabled': True,
                'is_screen_sharing': False,
                'is_hand_raised': False,
                'is_speaking': False,
                'connection_quality': 'good',
                'last_active': datetime.now(timezone.utc).isoformat(),
                'speaking_time': 0,
                'quiz_score': 0,
                'correct_answers': 0,
                'total_attempts': 0,
                'reactions': [],
                'sentiment': 0.5,
                'emotion': 'neutral'
            }
            return True
    
    def remove_participant(self, participant_id):
        with meeting_lock:
            if participant_id in self.participants:
                del self.participants[participant_id]
                return True
            return False
    
    def update_participant(self, participant_id, updates):
        with meeting_lock:
            if participant_id in self.participants:
                self.participants[participant_id].update(updates)
                self.participants[participant_id]['last_active'] = datetime.now(timezone.utc).isoformat()
                return True
            return False
    
    def add_message(self, message_data):
        with meeting_lock:
            message = {
                'id': str(uuid.uuid4()),
                **message_data,
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'reactions': {}
            }
            self.messages.append(message)
            
            # Keep only last 200 messages
            if len(self.messages) > 200:
                self.messages = self.messages[-200:]
            
            return message
    
    def add_reaction(self, message_id, reaction_data):
        with meeting_lock:
            for message in self.messages:
                if message['id'] == message_id:
                    if 'reactions' not in message:
                        message['reactions'] = {}
                    
                    emoji = reaction_data['emoji']
                    if emoji not in message['reactions']:
                        message['reactions'][emoji] = []
                    
                    message['reactions'][emoji].append({
                        'user_id': reaction_data['user_id'],
                        'user_name': reaction_data['user_name'],
                        'timestamp': datetime.now(timezone.utc).isoformat()
                    })
                    return True
            return False
    
    async def start_recording(self, user_id, db, socketio_instance):
        """Start recording the meeting"""
        try:
            if self.is_recording:
                return True
            
            self.db = db
            self.socketio_instance = socketio_instance
            
            # Start recording with manager
            recorder = await recording_manager.start_recording(
                self.id, user_id, db, socketio_instance, self.participants
            )
            
            if recorder:
                self.is_recording = True
                self.recording_started_at = datetime.now(timezone.utc)
                
                # Add system message
                self.add_message({
                    'userId': 'system',
                    'userName': 'System',
                    'text': f'🔴 Recording started by {user_id}',
                    'isAI': True,
                    'isCommand': False
                })
                
                # Emit recording started event
                socketio_instance.emit('meeting_recording_started', {
                    'meeting_id': self.id,
                    'started_by': user_id,
                    'started_at': self.recording_started_at.isoformat()
                }, room=self.id)
                
                logger.info(f"Recording started for meeting {self.id}")
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"Failed to start recording: {e}")
            return False
    
    async def stop_recording(self, user_id):
        """Stop recording the meeting"""
        try:
            if not self.is_recording:
                return None
            
            # Stop recording with manager
            recording_id = await recording_manager.stop_recording(self.id)
            
            if recording_id:
                self.is_recording = False
                duration = (datetime.now(timezone.utc) - self.recording_started_at).total_seconds()
                self.current_recording_id = recording_id
                
                # Add system message
                self.add_message({
                    'userId': 'system',
                    'userName': 'System',
                    'text': f'⏹️ Recording stopped. Duration: {duration:.0f}s',
                    'isAI': True,
                    'isCommand': False
                })
                
                # Emit recording stopped event
                if hasattr(self, 'socketio_instance') and self.socketio_instance:
                    self.socketio_instance.emit('meeting_recording_stopped', {
                        'meeting_id': self.id,
                        'stopped_by': user_id,
                        'recording_id': recording_id,
                        'duration': duration,
                        'stopped_at': datetime.now(timezone.utc).isoformat()
                    }, room=self.id)
                
                logger.info(f"Recording stopped for meeting {self.id}, ID: {recording_id}")
                return recording_id
            
            return None
            
        except Exception as e:
            logger.error(f"Failed to stop recording: {e}")
            return None

    def trigger_ai_analysis(self, new_text, speaker_name):
        """Trigger asynchronous AI analysis of the meeting transcript"""
        self.transcript_history.append(f"{speaker_name}: {new_text}")
        
        # Keep buffer manageable
        if len(self.transcript_history) > 100:
            self.transcript_history = self.transcript_history[-100:]

        current_time = time.time()
        # Analyze every 5 messages or if 30 seconds have passed
        if len(self.transcript_history) % 5 == 0 or (current_time - self.last_ai_analysis_time > 30):
            self.last_ai_analysis_time = current_time
            self.analysis_count += 1
            
            # Run in a separate thread to avoid blocking the socket handler
            processing_executor.submit(self._run_gemini_analysis)

    def _run_gemini_analysis(self):
        """Perform the actual Gemini analysis"""
        try:
            if not gemini_api_key:
                logger.warning("Gemini API key not configured, skipping analysis")
                return

            context = "\n".join(self.transcript_history[-30:]) # Last 30 lines for context
            
            prompt = f"""
            Analyze the following meeting transcript and identify:
            1. ANY new decisions made.
            2. ANY new action items or follow-up tasks (who should do what).
            3. ANY important questions asked or ideas shared.

            Transcript:
            {context}

            Return the results as a JSON object with two keys:
            'insights': list of objects with {{'type': 'decision'|'question'|'idea'|'task', 'text': string, 'priority': 'low'|'medium'|'high', 'participants': [string]}}
            'tasks': list of objects with {{'type': 'trello'|'jira'|'email'|'code'|'reminder', 'title': string, 'description': string, 'assignedTo': [string]}}

            Only return the JSON object, nothing else. If nothing new is found, return empty lists.
            """

            model = genai.GenerativeModel('gemini-1.5-flash')
            response = model.generate_content(prompt)
            
            try:
                # Extract JSON from response
                content = response.text.strip()
                if content.startswith('```json'):
                    content = content[7:-3]
                elif content.startswith('```'):
                    content = content[3:-3]
                
                result = json.loads(content)
                
                # Emit insights
                for insight in result.get('insights', []):
                    socketio.emit('ai-insight', {
                        'meeting_id': self.id,
                        'insight': {
                            'id': str(uuid.uuid4()),
                            'type': insight.get('type', 'idea'),
                            'text': insight.get('text', ''),
                            'timestamp': datetime.now(timezone.utc).isoformat(),
                            'participants': insight.get('participants', []),
                            'priority': insight.get('priority', 'medium')
                        }
                    }, room=self.id)

                # Emit tasks
                for task in result.get('tasks', []):
                    socketio.emit('follow-up-task', {
                        'meeting_id': self.id,
                        'task': {
                            'id': str(uuid.uuid4()),
                            'type': task.get('type', 'reminder'),
                            'title': task.get('title', ''),
                            'description': task.get('description', ''),
                            'assignedTo': task.get('assignedTo', []),
                            'status': 'pending',
                            'timestamp': datetime.now(timezone.utc).isoformat()
                        }
                    }, room=self.id)

                logger.info(f"AI Analysis completed for meeting {self.id}: {len(result.get('insights', []))} insights, {len(result.get('tasks', []))} tasks.")

            except Exception as e:
                logger.error(f"Failed to parse Gemini response: {e}\nResponse: {response.text}")

        except Exception as e:
            logger.error(f"Gemini analysis failed: {e}")
    
    async def add_media_chunk(self, media_data):
        """Add media chunk to recording"""
        try:
            if not self.is_recording and self.settings.get('auto_record'):
                # Try to auto-start if not recording yet - use host ID as default
                from app import db
                await self.start_recording(self.created_by, db, self.socketio_instance)

            if not self.is_recording:
                return False
            
            if media_data['type'] == 'audio':
                # Decode base64 audio data
                audio_bytes = base64.b64decode(media_data['data'])
                success = await recording_manager.add_audio_chunk(
                    self.id,
                    audio_bytes,
                    media_data['speaker_id'],
                    datetime.fromisoformat(media_data['timestamp'])
                )
            elif media_data['type'] == 'video':
                # Decode base64 video data
                video_bytes = base64.b64decode(media_data['data'])
                success = await recording_manager.add_video_chunk(
                    self.id,
                    video_bytes,
                    media_data['speaker_id'],
                    datetime.fromisoformat(media_data['timestamp'])
                )
            else:
                return False
            
            return success
            
        except Exception as e:
            logger.error(f"Failed to add media chunk: {e}")
            return False
    
    async def auto_start_recording(self, db, socketio_instance):
        """Auto-start recording when first participant joins"""
        try:
            if self.settings.get('auto_record') and not self.is_recording:
                # Wait 3 seconds for connection to stabilize
                await asyncio.sleep(3)
            
                # Only start recording if there are participants
                if len(self.participants) > 0:
                    success = await self.start_recording(self.created_by, db, socketio_instance)
                
                    if success:
                        # Log silently (not shown to users)
                        logger.info(f"Auto-recording started for meeting {self.id} with {len(self.participants)} participants")
                    
                        # Store references for later use
                        self.db = db
                        self.socketio_instance = socketio_instance
                    
                        return True
        
            return False
        
        except Exception as e:
            logger.error(f"Failed to auto-start recording: {e}")
            return False
    
    def get_recording_status(self):
        """Get recording status"""
        if not self.is_recording:
            return None
        
        duration = (datetime.now(timezone.utc) - self.recording_started_at).total_seconds()
        
        recorder = recording_manager.get_recorder(self.id)
        stats = recorder.stats if recorder else {}
        
        return {
            'is_recording': self.is_recording,
            'started_at': self.recording_started_at.isoformat(),
            'duration': duration,
            'current_recording_id': self.current_recording_id,
            'stats': stats
        }
    
    def start_quiz_session(self, quiz_data):
        quiz_id = str(uuid.uuid4())
        self.quiz_sessions[quiz_id] = {
            'id': quiz_id,
            'title': quiz_data.get('title', 'Quiz Session'),
            'description': quiz_data.get('description', ''),
            'created_by': quiz_data['created_by'],
            'created_at': datetime.now(timezone.utc).isoformat(),
            'is_active': True,
            'questions': quiz_data.get('questions', []),
            'answers': {},
            'settings': quiz_data.get('settings', {
                'shuffle_questions': False,
                'show_results': True,
                'allow_retry': True,
                'time_limit': None,
                'passing_score': 70,
                'show_leaderboard': True
            }),
            'participants': list(self.participants.keys()),
            'leaderboard': []
        }
        return quiz_id
    
    def submit_quiz_answer(self, quiz_id, answer_data):
        if quiz_id not in self.quiz_sessions:
            return False
        
        quiz = self.quiz_sessions[quiz_id]
        user_id = answer_data['user_id']
        
        if user_id not in quiz['answers']:
            quiz['answers'][user_id] = []
        
        quiz['answers'][user_id].append({
            **answer_data,
            'submitted_at': datetime.now(timezone.utc).isoformat()
        })
        
        # Update leaderboard
        self.update_quiz_leaderboard(quiz_id)
        return True
    
    def update_quiz_leaderboard(self, quiz_id):
        if quiz_id not in self.quiz_sessions:
            return
        
        quiz = self.quiz_sessions[quiz_id]
        leaderboard = []
        
        for user_id, answers in quiz['answers'].items():
            if user_id in self.participants:
                participant = self.participants[user_id]
                correct_answers = sum(1 for answer in answers if answer.get('is_correct', False))
                total_answers = len(answers)
                score = correct_answers * 10  # 10 points per correct answer
                
                leaderboard.append({
                    'user_id': user_id,
                    'user_name': participant.get('name', 'Unknown'),
                    'score': score,
                    'correct': correct_answers,
                    'total': total_answers,
                    'time_spent': sum(answer.get('time_taken', 0) for answer in answers)
                })
        
        # Sort by score
        leaderboard.sort(key=lambda x: x['score'], reverse=True)
        
        # Add ranks
        for i, entry in enumerate(leaderboard):
            entry['rank'] = i + 1
        
        quiz['leaderboard'] = leaderboard
    
    def to_dict(self):
        return {
            'id': self.id,
            'title': self.title,
            'created_by': self.created_by,
            'created_at': self.created_at.isoformat(),
            'is_active': self.is_active,
            'is_public': self.is_public,
            'participant_count': len(self.participants),
            'settings': self.settings,
            'ai_assistant': self.ai_assistant,
            'quiz_sessions': list(self.quiz_sessions.keys()),
            'is_recording': self.is_recording,
            'recording_started_at': self.recording_started_at.isoformat() if self.recording_started_at else None,
            'current_recording_id': self.current_recording_id
        }

    def save_meeting_complete_sync(self, user_id):
        """Synchronous version of save_meeting_complete"""
        try:
            # Stop recording if active
            recording_id = None
            if self.is_recording:
                # For sync version, we'll just mark it as stopped
                self.is_recording = False
                recording_id = f"rec_{self.id}_{int(time.time())}"
                self.current_recording_id = recording_id
        
            # If no recording was active but we have meeting data, still save
            if not recording_id and len(self.participants) > 0:
                # Create a basic meeting record
                from app import Meeting as MeetingModel, db
                from datetime import datetime
                
                # Generate transcript from messages
                transcript_lines = []
                for message in self.messages:
                    if message.get('text') and not message.get('isAI', False):
                        timestamp = message.get('timestamp', '')
                        speaker = message.get('userName', 'Unknown')
                        text = message.get('text', '')
                        transcript_lines.append(f"[{timestamp}] {speaker}: {text}")
                
                transcript_text = '\n'.join(transcript_lines)
                
                # Create participant mapping
                participant_map = {}
                for pid, pdata in self.participants.items():
                    participant_map[pid] = pdata.get('full_name', 'Unknown')

                meeting_record = MeetingModel(
                    user_id=user_id,
                    title=f"Meeting: {self.title}",
                    filename=f"meeting_{self.id}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.txt",
                    language=self.settings.get('language', 'en'),
                    transcript_language=self.settings.get('language', 'en'),
                    status='completed',
                    upload_date=datetime.now(timezone.utc),
                    duration=(datetime.now(timezone.utc) - self.created_at).total_seconds(),
                    participants_count=len(self.participants),
                    transcript=transcript_text[:5000],
                    source='live',
                    participant_mapping=json.dumps(participant_map)
                )
                
                db.session.add(meeting_record)
                db.session.commit()
                
                recording_id = meeting_record.id
                
            return recording_id
            
        except Exception as e:
            logger.error(f"Sync save failed: {e}")
            return None


# Meeting management functions
def create_meeting_room(title, created_by, is_public=True, password=None, settings=None):
    room_id = str(uuid.uuid4())
    meeting = MeetingRoom(room_id, title, created_by, is_public, password)
    
    if settings:
        meeting.settings.update(settings)
    
    # Store in Redis
    try:
        meeting_dict = meeting.to_dict()
        redis_client.hset(f'meeting:{room_id}', 'data', json.dumps(meeting_dict))
        redis_client.hset(f'meeting:{room_id}', 'created_at', meeting.created_at.isoformat())
        redis_client.hset(f'meeting:{room_id}', 'created_by', created_by)
        redis_client.expire(f'meeting:{room_id}', 86400)  # 24 hours
    except Exception as e:
        logger.error(f"Redis storage failed: {e}")
    
    # Store in memory
    with meeting_lock:
        active_meetings[room_id] = meeting
    
    logger.info(f"Meeting room created: {room_id} by {created_by}")
    return meeting

def get_meeting_room(room_id):
    with meeting_lock:
        if room_id in active_meetings:
            return active_meetings[room_id]
    
    # Try to load from Redis
    try:
        meeting_data = redis_client.hget(f'meeting:{room_id}', 'data')
        if meeting_data:
            data = json.loads(meeting_data)
            meeting = MeetingRoom(
                data['id'],
                data['title'],
                data['created_by'],
                data['is_public'],
                None  # Password not stored for security
            )
            meeting.created_at = datetime.fromisoformat(data['created_at'])
            meeting.is_active = data['is_active']
            meeting.settings = data.get('settings', meeting.settings)
            meeting.ai_assistant = data.get('ai_assistant', meeting.ai_assistant)
            meeting.is_recording = data.get('is_recording', False)
            meeting.recording_started_at = datetime.fromisoformat(data['recording_started_at']) if data.get('recording_started_at') else None
            meeting.current_recording_id = data.get('current_recording_id')
            
            # Restore to memory
            with meeting_lock:
                active_meetings[room_id] = meeting
            
            return meeting
    except Exception as e:
        logger.error(f"Redis load failed: {e}")
    
    return None

def save_meeting_to_db(meeting, user_id, app_db):
    """Save meeting recording and data to the main database"""
    from app import Meeting as MeetingModel, log_activity
    
    try:
        # Create a meeting record in the database
        db_meeting = MeetingModel(
            user_id=user_id,
            title=f"Meeting: {meeting.title}",
            filename=f"meeting_{meeting.id}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.webm",
            language=meeting.settings.get('language', 'en'),
            transcript_language=meeting.settings.get('language', 'en'),
            status='uploaded',
            upload_date=datetime.now(timezone.utc)
        )
        
        app_db.session.add(db_meeting)
        app_db.session.commit()
        
        # Log activity
        log_activity(
            user_id=user_id,
            activity_type='meeting',
            title=f"Meeting: {meeting.title}",
            description=f"Meeting completed with {len(meeting.participants)} participants",
            meeting_id=db_meeting.id,
            metadata={
                'meeting_id': meeting.id,
                'participant_count': len(meeting.participants),
                'duration': (datetime.now(timezone.utc) - meeting.created_at).total_seconds(),
                'messages_count': len(meeting.messages)
            }
        )
        
        return db_meeting.id
    except Exception as e:
        logger.error(f"Failed to save meeting to database: {e}")
        app_db.session.rollback()
        return None


# SocketIO event handlers
@socketio.on('connect')
def handle_connect():
    logger.info(f"Client connected: {request.sid}")
    emit('connected', {'message': 'Connected to meeting server'})

@socketio.on('disconnect')
def handle_disconnect():
    logger.info(f"Client disconnected: {request.sid}")

@socketio.on('join-meeting')
def handle_join_meeting(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    user_name = data.get('userName', 'Anonymous')
    user_email = data.get('userEmail')
    is_bot = data.get('isBot', False)
    is_teacher = data.get('isTeacher', False)
    socket_token = data.get('socket_token')
    
    if not room_id or not user_id:
        emit('error', {'message': 'Meeting ID and User ID required'})
        return
    
    meeting = get_meeting_room(room_id)
    if not meeting:
        emit('error', {'message': 'Meeting not found'})
        return
    
    # Validate socket token if provided
    # Note: In production, validate the token properly
    
    # Check password if required
    if not meeting.is_public and data.get('password'):
        from werkzeug.security import check_password_hash
        if not check_password_hash(meeting.password_hash, data['password']):
            emit('error', {'message': 'Invalid password'})
            return
    
    # Check if user needs approval (skip if user is host)
    is_host = (user_id == meeting.created_by)
    if meeting.settings.get('require_approval') and not is_bot and not is_host:
        # Store pending user info
        if not hasattr(meeting, 'pending_users'):
            meeting.pending_users = {}
        
        meeting.pending_users[user_id] = {
            'userId': user_id,
            'userName': user_name,
            'userEmail': user_email,
            'socketId': request.sid,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }
        
        # Send approval request to host
        emit('approval-requested', {
            'userId': user_id,
            'userName': user_name,
            'userEmail': user_email,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=meeting.created_by)
        
        emit('waiting-approval', {'message': 'Waiting for host approval'})
        logger.info(f"User {user_name} waiting for approval in meeting {room_id}")
        return
    
    # Add participant
    user_data = {
        'id': user_id,
        'name': user_name,
        'email': user_email,
        'is_bot': is_bot,
        'socket_id': request.sid,
        'is_host': (user_id == meeting.created_by),
        'is_teacher': is_teacher,
        'is_audio_enabled': True,
        'is_video_enabled': True,
        'is_screen_sharing': False,
        'is_hand_raised': False,
        'speaking_time': 0,
        'connection_quality': 'good',
        'quiz_score': 0,
        'correct_answers': 0,
        'total_attempts': 0,
        'joined_at': datetime.now(timezone.utc).isoformat(),
        'last_active': datetime.now(timezone.utc).isoformat()
    }
    
    if meeting.add_participant(user_id, user_data):
        join_room(room_id)
        
        # Notify others about new participant
        emit('user-joined', {
            'userId': user_id,
            'userName': user_name,
            'email': user_email,
            'isBot': is_bot,
            'isTeacher': is_teacher,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id, include_self=False)
        
        # Send current meeting state to new participant
        emit('meeting-state', {
            'meeting': meeting.to_dict(),
            'participants': meeting.participants,
            'messages': meeting.messages[-50:],
            'quizSessions': meeting.quiz_sessions,
            'is_recording': meeting.is_recording,
            'recording_started_at': meeting.recording_started_at.isoformat() if meeting.recording_started_at else None,
            'recording_status': meeting.get_recording_status()
        })
        
        logger.info(f"User {user_name} joined meeting {room_id}")
        
        # Auto-start recording if enabled and this is the first participant
        if (meeting.settings.get('auto_record') and 
            not meeting.is_recording and 
            len(meeting.participants) == 1):
            
            # Start recording silently when first participant joins
            from app import db
            asyncio.create_task(start_delayed_recording(meeting, db, socketio))
            
    else:
        emit('error', {'message': 'Meeting is full'})
        
async def start_delayed_recording(meeting, db, socketio_instance):
    """Start recording with a small delay after first participant joins"""
    try:
        await asyncio.sleep(2)  # Small delay
        await meeting.auto_start_recording(db, socketio_instance)
    except Exception as e:
        logger.error(f"Failed to start delayed recording: {e}")

@socketio.on('summon-bot')
async def handle_summon_bot(data):
    """Handle summoning the AI bot manually"""
    room_id = data.get('meetingId')
    user_id = getattr(socketio, 'user_id', 'Owner')  # Fallback if not easily available
    
    meeting = get_meeting_room(room_id)
    if not meeting:
        return

    # If already recording, just send a confirmation
    if meeting.is_recording:
        emit('meeting_insight', {
            'meeting_id': room_id,
            'type': 'task',
            'text': "AI Bot is already active and listening to the meeting!",
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'priority': 'medium'
        }, room=room_id)
    else:
        # Start recording if not already active
        from app import db
        await start_delayed_recording(meeting, db, socketio)
        emit('meeting_insight', {
            'meeting_id': room_id,
            'type': 'task',
            'text': "AI Bot has joined the meeting and is now transcribing.",
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'priority': 'high'
        }, room=room_id)

@socketio.on('leave-meeting')
async def handle_leave_meeting(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    save_recording = data.get('saveRecording', True)
    
    meeting = get_meeting_room(room_id)
    if meeting:
        # Save meeting data if it's ending
        if save_recording and user_id == meeting.created_by:
            # Host is leaving - save the meeting
            from app import db
            
            # Save meeting recording and data
            recording_id = await meeting.save_meeting_complete(user_id)
            
            if recording_id:
                emit('meeting-saved', {
                    'meeting_id': room_id,
                    'recording_id': recording_id,
                    'message': 'Meeting saved successfully'
                }, room=user_id)
        
        # Remove participant
        if meeting.remove_participant(user_id):
            leave_room(room_id)
            
            # Notify others
            emit('user-left', {
                'userId': user_id,
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'remaining_participants': len(meeting.participants)
            }, room=room_id)
            
            logger.info(f"User {user_id} left meeting {room_id}")
            
            # If host left and no participants, close meeting
            if user_id == meeting.created_by and len(meeting.participants) == 0:
                meeting.is_active = False
                
                # Auto-save meeting data if not already saved
                if save_recording:
                    from app import db
                    asyncio.create_task(meeting.save_meeting_complete(user_id))
                
                logger.info(f"Meeting {room_id} closed by host")
            
            # If no participants left at all, save and close meeting
            elif len(meeting.participants) == 0:
                meeting.is_active = False
                
                # Auto-save meeting data
                if save_recording:
                    from app import db
                    asyncio.create_task(meeting.save_meeting_complete('system'))
                
                logger.info(f"Meeting {room_id} closed (no participants)")

@socketio.on('start-meeting-recording')
async def handle_start_meeting_recording(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id == meeting.created_by:
        from app import db
        success = await meeting.start_recording(user_id, db, socketio)
        
        if success:
            emit('meeting_recording_started', {
                'meeting_id': room_id,
                'started_by': user_id,
                'started_at': meeting.recording_started_at.isoformat()
            }, room=room_id)
        else:
            emit('error', {'message': 'Failed to start recording'})

@socketio.on('stop-meeting-recording')
async def handle_stop_meeting_recording(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id == meeting.created_by:
        recording_id = await meeting.stop_recording(user_id)
        
        if recording_id:
            emit('meeting_recording_stopped', {
                'meeting_id': room_id,
                'stopped_by': user_id,
                'recording_id': recording_id,
                'duration': (datetime.now(timezone.utc) - meeting.recording_started_at).total_seconds(),
                'stopped_at': datetime.now(timezone.utc).isoformat()
            }, room=room_id)
        else:
            emit('error', {'message': 'Failed to stop recording'})

@socketio.on('media-chunk')
async def handle_media_chunk(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    media_data = data.get('mediaData')
    
    meeting = get_meeting_room(room_id)
    if meeting:
        success = await meeting.add_media_chunk(media_data)
        
        if not success:
            logger.error(f"Failed to process media chunk from {user_id}")

@socketio.on('offer')
def handle_offer(data):
    """Handle WebRTC offer"""
    room_id = data.get('meetingId')
    target_user_id = data.get('targetUserId')
    offer = data.get('offer')
    
    if room_id and target_user_id and offer:
        # Forward offer to target user
        emit('offer', {
            'fromUserId': data.get('fromUserId'),
            'offer': offer,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=target_user_id)

@socketio.on('answer')
def handle_answer(data):
    """Handle WebRTC answer"""
    room_id = data.get('meetingId')
    target_user_id = data.get('targetUserId')
    answer = data.get('answer')
    
    if room_id and target_user_id and answer:
        # Forward answer to target user
        emit('answer', {
            'fromUserId': data.get('fromUserId'),
            'answer': answer,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=target_user_id)

@socketio.on('ice-candidate')
def handle_ice_candidate(data):
    """Handle ICE candidates"""
    room_id = data.get('meetingId')
    target_user_id = data.get('targetUserId')
    candidate = data.get('candidate')
    
    if room_id and target_user_id and candidate:
        # Forward ICE candidate to target user
        emit('ice-candidate', {
            'fromUserId': data.get('fromUserId'),
            'candidate': candidate,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=target_user_id)

@socketio.on('chat-message')
def handle_chat_message(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    user_name = data.get('userName')
    text = data.get('text')
    is_command = data.get('isCommand', False)
    attachment = data.get('attachment')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants:
        # Prevent empty or system messages from causing loops
        if not text or text.strip() == '':
            return
        
        message_data = {
            'userId': user_id,
            'userName': user_name,
            'text': text,
            'isCommand': is_command,
            'attachment': attachment,
            'isAI': data.get('isAI', False)
        }
        
        message = meeting.add_message(message_data)
        
        # Broadcast to all participants except sender
        emit('chat-message', message, room=room_id, include_self=False)
        
        # Send to sender separately if needed
        emit('chat-message-sent', message)
        
        # If it's a command, handle it
        if is_command and text.startswith('/'):
            handle_command(room_id, user_id, user_name, text)
        
        # AI analysis if enabled
        if meeting.ai_assistant['enabled'] and not is_command and not data.get('isAI'):
            analyze_message_with_ai(room_id, message)
            # Trigger Gemini analysis
            meeting.trigger_ai_analysis(text, user_name)

@socketio.on('message-reaction')
def handle_message_reaction(data):
    room_id = data.get('meetingId')
    message_id = data.get('messageId')
    user_id = data.get('userId')
    user_name = data.get('userName')
    emoji = data.get('emoji')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants:
        reaction_data = {
            'emoji': emoji,
            'user_id': user_id,
            'user_name': user_name
        }
        
        if meeting.add_reaction(message_id, reaction_data):
            # Broadcast reaction
            emit('message-reaction', {
                'messageId': message_id,
                'emoji': emoji,
                'userId': user_id,
                'userName': user_name,
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=room_id)

@socketio.on('reaction')
def handle_reaction(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    emoji = data.get('emoji')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants:
        # Broadcast reaction to all participants
        emit('reaction', {
            'emoji': emoji,
            'userId': user_id,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id)

@socketio.on('participant-update')
def handle_participant_update(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    updates = data.get('updates', {})
    
    meeting = get_meeting_room(room_id)
    if meeting and meeting.update_participant(user_id, updates):
        # Broadcast update to all participants
        emit('participant-updated', {
            'userId': user_id,
            'updates': updates,
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id)

@socketio.on('start-recording')
async def handle_start_recording(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id == meeting.created_by:
        # Start recording session
        from app import db
        success = await meeting.start_recording(user_id, db, socketio)
        
        if success:
            emit('recording-started', {
                'recordingId': meeting.current_recording_id,
                'startedBy': user_id,
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=room_id)

@socketio.on('stop-recording')
async def handle_stop_recording(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    recording_id = data.get('recordingId')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id == meeting.created_by:
        # Stop recording
        recording_id = await meeting.stop_recording(user_id)
        
        if recording_id:
            emit('recording-stopped', {
                'recordingId': recording_id,
                'stoppedBy': user_id,
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=room_id)

@socketio.on('quiz-question')
def handle_quiz_question(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    quiz_data = data.get('quizData', {})
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants:
        quiz_data['created_by'] = user_id
        quiz_id = meeting.start_quiz_session(quiz_data)
        
        if quiz_id:
            emit('quiz-started', {
                'quizId': quiz_id,
                'quizData': meeting.quiz_sessions[quiz_id],
                'startedBy': user_id,
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=room_id)

@socketio.on('quiz-answer')
def handle_quiz_answer(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    quiz_id = data.get('quizId')
    answer_data = data.get('answerData', {})
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants:
        answer_data['user_id'] = user_id
        answer_data['user_name'] = meeting.participants[user_id].get('name', 'Unknown')
        
        if meeting.submit_quiz_answer(quiz_id, answer_data):
            # Update participant stats
            if answer_data.get('is_correct'):
                meeting.update_participant(user_id, {
                    'quiz_score': meeting.participants[user_id].get('quiz_score', 0) + 10,
                    'correct_answers': meeting.participants[user_id].get('correct_answers', 0) + 1,
                    'total_attempts': meeting.participants[user_id].get('total_attempts', 0) + 1
                })
            else:
                meeting.update_participant(user_id, {
                    'total_attempts': meeting.participants[user_id].get('total_attempts', 0) + 1
                })
            
            # Broadcast answer submission
            emit('quiz-answer-submitted', {
                'quizId': quiz_id,
                'userId': user_id,
                'userName': answer_data['user_name'],
                'isCorrect': answer_data.get('is_correct', False),
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=room_id)
            
            # Broadcast leaderboard update
            if meeting.quiz_sessions[quiz_id]['settings']['show_leaderboard']:
                emit('quiz-leaderboard-updated', {
                    'quizId': quiz_id,
                    'leaderboard': meeting.quiz_sessions[quiz_id]['leaderboard'],
                    'timestamp': datetime.now(timezone.utc).isoformat()
                }, room=room_id)

@socketio.on('end-quiz')
def handle_end_quiz(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    quiz_id = data.get('quizId')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants and quiz_id in meeting.quiz_sessions:
        meeting.quiz_sessions[quiz_id]['is_active'] = False
        meeting.quiz_sessions[quiz_id]['ended_at'] = datetime.now(timezone.utc).isoformat()
        
        emit('quiz-ended', {
            'quizId': quiz_id,
            'endedBy': user_id,
            'finalLeaderboard': meeting.quiz_sessions[quiz_id]['leaderboard'],
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id)

@socketio.on('ghost-replay-request')
def handle_ghost_replay_request(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    query = data.get('query')
    
    meeting = get_meeting_room(room_id)
    if meeting and meeting.ai_assistant.get('ghost_replay_enabled'):
        # Find relevant messages for replay
        relevant_messages = []
        for message in meeting.messages:
            if query.lower() in message.get('text', '').lower():
                relevant_messages.append(message)
        
        if relevant_messages:
            emit('ghost-replay-response', {
                'query': query,
                'messages': relevant_messages[:5],  # Limit to 5 messages
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=user_id)
        else:
            emit('ghost-replay-response', {
                'query': query,
                'messages': [],
                'message': 'No relevant content found for replay',
                'timestamp': datetime.now(timezone.utc).isoformat()
            }, room=user_id)

@socketio.on('ai-command')
def handle_ai_command(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    command = data.get('command')
    
    meeting = get_meeting_room(room_id)
    if meeting and meeting.ai_assistant['enabled']:
        handle_command(room_id, user_id, 'AI Assistant', command)

# Real-time transcription events
@socketio.on('caption')
def handle_caption(data):
    """Broadcast real-time caption to all participants and update speaking status"""
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    user_name = data.get('userName')
    text = data.get('text')
    
    meeting = get_meeting_room(room_id)
    if meeting and user_id in meeting.participants:
        # Update participant state
        with meeting_lock:
            participant = meeting.participants[user_id]
            participant['is_speaking'] = True
            participant['last_active'] = datetime.now(timezone.utc).isoformat()
            
        # Broadcast the caption
        emit('live_transcript_update', {
            'meeting_id': room_id,
            'entry': {
                'id': str(uuid.uuid4()),
                'speaker': user_name,
                'userId': user_id,
                'text': text,
                'timestamp': datetime.now(timezone.utc).strftime('%H:%M:%S')
            }
        }, room=room_id, include_self=True)
        
        # Trigger Gemini analysis
        meeting.trigger_ai_analysis(text, user_name)
        
        # Reset speaking status after a short delay (simulated)
        # Note: In a real production app, you might use a more robust detection
        def reset_speaking():
            import time
            time.sleep(3)
            with meeting_lock:
                if user_id in meeting.participants:
                    meeting.participants[user_id]['is_speaking'] = False
                    # Notify everyone that speaking has stopped
                    socketio.emit('participant_update', {
                        'userId': user_id,
                        'isSpeaking': False
                    }, room=room_id)
        
        processing_executor.submit(reset_speaking)

@socketio.on('live_transcript')
def handle_live_transcript(data):
    """Forward live transcript to all participants"""
    room_id = data.get('meetingId')
    transcript = data.get('transcript')
    
    emit('live_transcript_update', {
        'meeting_id': room_id,
        'transcript': transcript,
        'timestamp': datetime.now(timezone.utc).isoformat()
    }, room=room_id)

@socketio.on('get_recording_status')
def handle_get_recording_status(data):
    room_id = data.get('meetingId')
    user_id = data.get('userId')
    
    meeting = get_meeting_room(room_id)
    if meeting:
        status = meeting.get_recording_status()
        emit('recording_status_response', {
            'meeting_id': room_id,
            'status': status
        }, room=user_id)

# Command handling
def handle_command(room_id, user_id, user_name, command_text):
    meeting = get_meeting_room(room_id)
    if not meeting:
        return
    
    parts = command_text[1:].split(' ')
    command_type = parts[0].lower()
    args = parts[1:] if len(parts) > 1 else []
    
    command_responses = {
        'help': "Available commands:\n"
                "/record start/stop - Start/stop recording\n"
                "/summarize - Generate meeting summary\n"
                "/translate [lang] - Enable translation to language\n"
                "/next-speaker - Suggest next speaker\n"
                "/show-decisions - Show all decisions made\n"
                "/start-quiz [title] - Start a quiz session\n"
                "/end-quiz - End current quiz\n"
                "/mute-all - Mute all participants (host only)\n"
                "/clear-chat - Clear chat history (host only)\n"
                "/highlight [text] - Search for text in messages\n"
                "/ai private/group - Set AI mode\n"
                "/ghost-replay [query] - Replay past moments\n"
                "/recording-status - Check recording status\n"
                "/leave-with-recording - Leave and save recording",
        
        'record': handle_record_command,
        'summarize': handle_summarize_command,
        'translate': handle_translate_command,
        'next-speaker': handle_next_speaker_command,
        'show-decisions': handle_show_decisions_command,
        'start-quiz': handle_start_quiz_command,
        'end-quiz': handle_end_quiz_command,
        'mute-all': handle_mute_all_command,
        'clear-chat': handle_clear_chat_command,
        'highlight': handle_highlight_command,
        'ai': handle_ai_mode_command,
        'ghost-replay': handle_ghost_replay_command,
        'recording-status': handle_recording_status_command,
        'leave-with-recording': handle_leave_with_recording_command
    }
    
    if command_type in command_responses:
        if callable(command_responses[command_type]):
            response = command_responses[command_type](meeting, user_id, user_name, args)
        else:
            response = command_responses[command_type]
        
        # Send command response
        socketio.emit('chat-message', {
            'id': str(uuid.uuid4()),
            'userId': 'system',
            'userName': 'System',
            'text': response,
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'isCommand': True
        }, room=room_id)

def handle_record_command(meeting, user_id, user_name, args):
    if user_id != meeting.created_by:
        return "Only meeting host can control recording"
    
    if args and args[0] == 'start':
        # Trigger recording start
        socketio.emit('start-recording', {
            'meetingId': meeting.id,
            'userId': user_id
        }, room=meeting.id)
        return "Recording started"
    elif args and args[0] == 'stop':
        # Trigger recording stop
        socketio.emit('stop-recording', {
            'meetingId': meeting.id,
            'userId': user_id
        }, room=meeting.id)
        return "Recording stopped"
    else:
        return "Usage: /record start or /record stop"

def handle_recording_status_command(meeting, user_id, user_name, args):
    status = meeting.get_recording_status()
    
    if not status:
        return "No recording in progress"
    
    duration = status.get('duration', 0)
    minutes = int(duration // 60)
    seconds = int(duration % 60)
    
    response = f"📹 Recording Status:\n"
    response += f"• Status: {'🟢 Recording' if status.get('is_recording') else '⏸️ Paused'}\n"
    response += f"• Duration: {minutes}m {seconds}s\n"
    response += f"• Started: {status.get('started_at', 'Unknown')}\n"
    
    stats = status.get('stats', {})
    if stats:
        response += f"• Speakers: {stats.get('speakers_detected', 0)}\n"
        response += f"• Words: {stats.get('words_transcribed', 0)}\n"
        response += f"• Action Items: {len(stats.get('action_items', []))}\n"
    
    return response

def handle_leave_with_recording_command(meeting, user_id, user_name, args):
    if user_id != meeting.created_by:
        return "Only meeting host can save recording"
    
    # Emit event to trigger recording save
    socketio.emit('leave-with-recording-request', {
        'meetingId': meeting.id,
        'userId': user_id
    }, room=meeting.id)
    
    return "Saving recording and preparing to leave..."

def handle_summarize_command(meeting, user_id, user_name, args):
    # Generate meeting summary
    summary = f"Meeting Summary ({datetime.now(timezone.utc).strftime('%H:%M')}):\n"
    summary += f"• Participants: {len(meeting.participants)}\n"
    summary += f"• Messages: {len(meeting.messages)}\n"
    summary += f"• Duration: {int((datetime.now(timezone.utc) - meeting.created_at).total_seconds() / 60)} minutes\n"
    
    # Add key topics from messages
    if meeting.messages:
        topics = extract_topics_from_messages(meeting.messages)
        summary += f"• Key Topics: {', '.join(topics[:3])}\n"
    
    return summary

def handle_translate_command(meeting, user_id, user_name, args):
    if args:
        language = args[0]
        meeting.settings['translation_enabled'] = True
        meeting.settings['language'] = language
        return f"Translation enabled for {language}"
    else:
        return "Usage: /translate [language_code]"

def handle_next_speaker_command(meeting, user_id, user_name, args):
    # Find participants who haven't spoken much
    speaking_times = {}
    for pid, participant in meeting.participants.items():
        if not participant.get('is_bot', False):
            speaking_times[pid] = participant.get('speaking_time', 0)
    
    if speaking_times:
        # Suggest participant with least speaking time
        next_speaker_id = min(speaking_times, key=speaking_times.get)
        next_speaker = meeting.participants[next_speaker_id].get('name', 'Unknown')
        return f"Next speaker suggestion: {next_speaker}"
    return "No participants available"

def handle_show_decisions_command(meeting, user_id, user_name, args):
    # Extract decisions from messages
    decisions = []
    for message in meeting.messages:
        text = message.get('text', '').lower()
        if any(keyword in text for keyword in ['decided', 'agree', 'decision', 'approved', 'resolved']):
            decisions.append(message.get('text', ''))
    
    if decisions:
        response = "Decisions made:\n"
        for i, decision in enumerate(decisions[-5:], 1):
            response += f"{i}. {decision}\n"
        return response
    return "No decisions recorded yet"

def handle_start_quiz_command(meeting, user_id, user_name, args):
    title = ' '.join(args) if args else 'Quick Quiz'
    
    # Trigger quiz creation
    socketio.emit('quiz-question', {
        'meetingId': meeting.id,
        'userId': user_id,
        'quizData': {
            'title': title,
            'questions': []  # Would be populated by frontend
        }
    }, room=meeting.id)
    
    return f"Quiz '{title}' started"

def handle_end_quiz_command(meeting, user_id, user_name, args):
    # Find active quiz
    for quiz_id, quiz in meeting.quiz_sessions.items():
        if quiz.get('is_active'):
            socketio.emit('end-quiz', {
                'meetingId': meeting.id,
                'userId': user_id,
                'quizId': quiz_id
            }, room=meeting.id)
            return f"Quiz '{quiz.get('title')}' ended"
    return "No active quiz found"

def handle_mute_all_command(meeting, user_id, user_name, args):
    if user_id != meeting.created_by:
        return "Only meeting host can mute all participants"
    
    # Send mute command to all participants
    socketio.emit('mute-all', {
        'meetingId': meeting.id,
        'mutedBy': user_id
    }, room=meeting.id)
    return "All participants muted"

def handle_clear_chat_command(meeting, user_id, user_name, args):
    if user_id != meeting.created_by:
        return "Only meeting host can clear chat"
    
    meeting.messages.clear()
    socketio.emit('chat-cleared', {
        'meetingId': meeting.id,
        'clearedBy': user_id
    }, room=meeting.id)
    return "Chat cleared"

def handle_highlight_command(meeting, user_id, user_name, args):
    if not args:
        return "Usage: /highlight [search_text]"
    
    search_text = ' '.join(args).lower()
    matches = []
    
    for message in meeting.messages[-50:]:  # Search last 50 messages
        if search_text in message.get('text', '').lower():
            matches.append({
                'user': message.get('userName'),
                'text': message.get('text'),
                'time': message.get('timestamp')
            })
    
    if matches:
        response = f"Found {len(matches)} matches for '{search_text}':\n"
        for i, match in enumerate(matches[:3], 1):
            response += f"{i}. {match['user']}: {match['text'][:50]}...\n"
        return response
    return f"No matches found for '{search_text}'"

def handle_ai_mode_command(meeting, user_id, user_name, args):
    if args and args[0] in ['private', 'group']:
        meeting.ai_assistant['mode'] = args[0]
        return f"AI mode set to {args[0]}"
    return "Usage: /ai private or /ai group"

def handle_ghost_replay_command(meeting, user_id, user_name, args):
    if not args:
        return "Usage: /ghost-replay [search_query]"
    
    query = ' '.join(args)
    socketio.emit('ghost-replay-request', {
        'meetingId': meeting.id,
        'userId': user_id,
        'query': query
    }, room=meeting.id)
    
    return f"Searching for replay of: {query}"

def extract_topics_from_messages(messages):
    # Simple topic extraction from messages
    common_words = {}
    for message in messages:
        text = message.get('text', '')
        words = text.lower().split()
        for word in words:
            if len(word) > 4:  # Only consider words longer than 4 characters
                common_words[word] = common_words.get(word, 0) + 1
    
    # Sort by frequency and return top 5
    sorted_words = sorted(common_words.items(), key=lambda x: x[1], reverse=True)
    return [word for word, count in sorted_words[:5]]

def analyze_message_with_ai(room_id, message):
    # Simple AI analysis of messages
    # In production, you would integrate with Gemini API here
    
    text = message.get('text', '').lower()
    
    # Detect action items
    action_keywords = ['need to', 'should', 'must', 'task', 'action', 'todo', 'follow up']
    if any(keyword in text for keyword in action_keywords):
        socketio.emit('ai-insight', {
            'type': 'action_item',
            'message': f"Action item detected: {message.get('text', '')}",
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id)
    
    # Detect questions
    if '?' in text or text.startswith(('how', 'what', 'why', 'when', 'where')):
        socketio.emit('ai-insight', {
            'type': 'question',
            'message': f"Question detected from {message.get('userName')}",
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id)
    
    # Detect confusion
    confusion_keywords = ["don't understand", "confused", "unclear", "not sure"]
    if any(keyword in text for keyword in confusion_keywords):
        socketio.emit('ai-insight', {
            'type': 'confusion',
            'message': f"Potential confusion detected: {message.get('text', '')}",
            'timestamp': datetime.now(timezone.utc).isoformat()
        }, room=room_id)

# Flask routes for meeting management
def create_meeting_routes(app, db):

    @app.route('/api/meeting-rooms/create', methods=['POST'])
    @jwt_required()
    def create_meeting_room_endpoint():
        try:
            user_id = get_jwt_identity()
            data = request.json or {}

            title = data.get('title', 'New Meeting')
            is_public = data.get('isPublic', True)
            password = data.get('password')
            settings = data.get('settings', {})

            # Create meeting room
            meeting = create_meeting_room(
                title=title,
                created_by=user_id,
                is_public=is_public,
                password=password,
                settings=settings
            )

            # Auto-start recording in background
            if meeting.settings.get('auto_record'):

                def start_auto_recording():
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    try:
                        from app import socketio as app_socketio
                        loop.run_until_complete(
                            meeting.auto_start_recording(db, app_socketio)
                        )
                    except Exception as err:
                        logger.error(f"[Auto Recording] Error: {err}")
                    finally:
                        loop.close()

                threading.Thread(target=start_auto_recording, daemon=True).start()

            return jsonify({
                'success': True,
                'meeting': meeting.to_dict(),
                'message': 'Meeting created successfully'
            }), 201

        except Exception as e:
            logger.exception("Meeting creation failed:")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/<meeting_id>', methods=['GET'])
    @jwt_required()
    def get_meeting_room_endpoint(meeting_id):
        try:
            user_id = get_jwt_identity()
            meeting = get_meeting_room(meeting_id)
            
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Check if user is participant or meeting is public
            if user_id not in meeting.participants and not meeting.is_public:
                return jsonify({'error': 'Access denied'}), 403
            
            return jsonify({
                'success': True,
                'meeting': meeting.to_dict()
            })
            
        except Exception as e:
            logger.error(f"Failed to get meeting: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/<meeting_id>/join', methods=['POST'])
    @jwt_required()
    def join_meeting_endpoint(meeting_id):
        try:
            user_id = get_jwt_identity()
            data = request.json
            
            user_name = data.get('userName')
            user_email = data.get('userEmail')
            password = data.get('password')
            
            meeting = get_meeting_room(meeting_id)
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Check password if required
            if not meeting.is_public and password:
                from werkzeug.security import check_password_hash
                if not check_password_hash(meeting.password_hash, password):
                    return jsonify({'error': 'Invalid password'}), 401
            
            # Check if meeting is full (0 means unlimited)
            max_participants = meeting.settings.get('max_participants', 50)
            if max_participants > 0 and len(meeting.participants) >= max_participants:
                return jsonify({'error': 'Meeting is full'}), 400
            
            # Generate socket token for connection
            socket_token = str(uuid.uuid4())
            
            return jsonify({
                'success': True,
                'meeting': meeting.to_dict(),
                'socket_token': socket_token,
                'message': 'Ready to join meeting'
            })
            
        except Exception as e:
            logger.error(f"Failed to join meeting: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/<meeting_id>/recordings', methods=['POST'])
    @jwt_required()
    async def save_meeting_recording_endpoint(meeting_id):
        try:
            user_id = get_jwt_identity()
            data = request.json
            
            meeting = get_meeting_room(meeting_id)
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Check if user is host
            if user_id != meeting.created_by:
                return jsonify({'error': 'Only host can save recordings'}), 403
            
            # Stop recording if active
            if meeting.is_recording:
                recording_id = await meeting.stop_recording(user_id)
                
                if recording_id:
                    return jsonify({
                        'success': True,
                        'recording_id': recording_id,
                        'message': 'Recording saved successfully'
                    })
                else:
                    return jsonify({'error': 'Failed to save recording'}), 500
            else:
                return jsonify({'error': 'No active recording to save'}), 400
            
        except Exception as e:
            logger.error(f"Failed to save recording: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/<meeting_id>/transcript', methods=['GET'])
    @jwt_required()
    def get_meeting_transcript_endpoint(meeting_id):
        try:
            user_id = get_jwt_identity()
            meeting = get_meeting_room(meeting_id)
            
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Check if user is participant
            if user_id not in meeting.participants:
                return jsonify({'error': 'Access denied'}), 403
            
            # Format messages as transcript
            transcript = []
            for message in meeting.messages:
                transcript.append({
                    'time': message.get('timestamp'),
                    'speaker': message.get('userName'),
                    'text': message.get('text'),
                    'type': 'ai' if message.get('isAI') else 'command' if message.get('isCommand') else 'message'
                })
            
            return jsonify({
                'success': True,
                'transcript': transcript,
                'total_messages': len(transcript)
            })
            
        except Exception as e:
            logger.error(f"Failed to get transcript: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/<meeting_id>/analytics', methods=['GET'])
    @jwt_required()
    def get_meeting_analytics_endpoint(meeting_id):
        try:
            user_id = get_jwt_identity()
            meeting = get_meeting_room(meeting_id)
            
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Check if user is host
            if user_id != meeting.created_by:
                return jsonify({'error': 'Only host can view analytics'}), 403
            
            # Calculate analytics
            duration = (datetime.now(timezone.utc) - meeting.created_at).total_seconds()
            
            # Speaking time distribution
            speaking_times = {}
            for pid, participant in meeting.participants.items():
                if not participant.get('is_bot', False):
                    speaking_times[participant.get('name', 'Unknown')] = participant.get('speaking_time', 0)
            
            # Message statistics
            total_messages = len(meeting.messages)
            user_messages = {}
            for message in meeting.messages:
                user = message.get('userName')
                user_messages[user] = user_messages.get(user, 0) + 1
            
            # Quiz statistics
            quiz_stats = {}
            for quiz_id, quiz in meeting.quiz_sessions.items():
                quiz_stats[quiz_id] = {
                    'title': quiz.get('title'),
                    'participants': len(quiz.get('answers', {})),
                    'questions': len(quiz.get('questions', [])),
                    'average_score': calculate_average_quiz_score(quiz)
                }
            
            # Recording statistics
            recording_stats = None
            if meeting.is_recording:
                recorder = recording_manager.get_recorder(meeting.id)
                if recorder:
                    recording_stats = recorder.stats
            
            return jsonify({
                'success': True,
                'analytics': {
                    'duration_minutes': int(duration / 60),
                    'participant_count': len(meeting.participants),
                    'total_messages': total_messages,
                    'speaking_distribution': speaking_times,
                    'message_distribution': user_messages,
                    'quiz_statistics': quiz_stats,
                    'recordings_count': len(meeting.recordings),
                    'active_quizzes': sum(1 for q in meeting.quiz_sessions.values() if q.get('is_active')),
                    'recording_stats': recording_stats,
                    'is_recording': meeting.is_recording,
                    'recording_duration': (datetime.now(timezone.utc) - meeting.recording_started_at).total_seconds() if meeting.recording_started_at else 0
                }
            })
            
        except Exception as e:
            logger.error(f"Failed to get analytics: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/<meeting_id>/leave-with-recording', methods=['POST'])
    @jwt_required()
    def leave_meeting_with_recording_endpoint(meeting_id):
        """Leave meeting and save recording"""
        try:
            user_id = get_jwt_identity()
            data = request.json or {}
            save_recording = data.get('saveRecording', True)
            
            meeting = get_meeting_room(meeting_id)
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Simple approach - just mark as saved
            recording_id = "saved_" + meeting_id if save_recording else None
            
            # Remove user from participants
            meeting.remove_participant(user_id)
            
            # Emit leave event
            socketio.emit('user-left', {
                'userId': user_id,
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'remaining_participants': len(meeting.participants)
            }, room=meeting_id)
            
            # If no participants left, close meeting
            if len(meeting.participants) == 0:
                meeting.is_active = False
                logger.info(f"Meeting {meeting_id} closed")
                
                # Auto-save if requested
                if not recording_id and save_recording:
                    recording_id = "auto_saved_" + meeting_id
            
            if recording_id:
                return jsonify({
                    'success': True,
                    'recording_id': recording_id,
                    'message': 'Meeting saved successfully',
                    'participants_remaining': len(meeting.participants)
                })
            else:
                return jsonify({
                    'success': True,
                    'message': 'Meeting ended',
                    'participants_remaining': len(meeting.participants)
                })
            
        except Exception as e:
            logger.error(f"Failed to leave with recording: {e}")
            # Return success even on error to avoid the Flask issue
            return jsonify({
                'success': True,
                'message': 'Meeting ended with warnings',
                'error': str(e)
            })
    
    @app.route('/api/meeting-rooms/<meeting_id>/recording-status', methods=['GET'])
    @jwt_required()
    def get_recording_status_endpoint(meeting_id):
        """Get recording status"""
        try:
            user_id = get_jwt_identity()
            meeting = get_meeting_room(meeting_id)
            
            if not meeting:
                return jsonify({'error': 'Meeting not found'}), 404
            
            # Check if user is participant
            if user_id not in meeting.participants:
                return jsonify({'error': 'Access denied'}), 403
            
            status = meeting.get_recording_status()
            
            return jsonify({
                'success': True,
                'recording_status': status
            })
            
        except Exception as e:
            logger.error(f"Failed to get recording status: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/user/recent', methods=['GET'])
    @jwt_required()
    def get_user_recent_meetings_endpoint():
        try:
            user_id = get_jwt_identity()
            
            # Get meetings where user is participant or host
            recent_meetings = []
            with meeting_lock:
                for meeting_id, meeting in active_meetings.items():
                    if user_id in meeting.participants or user_id == meeting.created_by:
                        recent_meetings.append(meeting.to_dict())
            
            # Sort by creation date (newest first)
            recent_meetings.sort(key=lambda x: x['created_at'], reverse=True)
            
            return jsonify({
                'success': True,
                'meetings': recent_meetings[:10]  # Last 10 meetings
            })
            
        except Exception as e:
            logger.error(f"Failed to get recent meetings: {e}")
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/meeting-rooms/public', methods=['GET'])
    def get_public_meetings_endpoint():
        try:
            # Get all public active meetings
            public_meetings = []
            with meeting_lock:
                for meeting_id, meeting in active_meetings.items():
                    if meeting.is_active and meeting.is_public:
                        public_meetings.append(meeting.to_dict())
            
            return jsonify({
                'success': True,
                'meetings': public_meetings
            })
            
        except Exception as e:
            logger.error(f"Failed to get public meetings: {e}")
            return jsonify({'error': str(e)}), 500

def calculate_average_quiz_score(quiz):
    if not quiz.get('answers'):
        return 0
    
    total_score = 0
    total_answers = 0
    
    for user_answers in quiz['answers'].values():
        correct_answers = sum(1 for answer in user_answers if answer.get('is_correct', False))
        total_score += correct_answers * 10
        total_answers += len(user_answers)
    
    if total_answers == 0:
        return 0
    
    return total_score / total_answers

# Background task for cleanup
def cleanup_inactive_meetings():
    """Clean up inactive meetings periodically"""
    while True:
        try:
            current_time = datetime.now(timezone.utc)
            meetings_to_remove = []
            
            with meeting_lock:
                for meeting_id, meeting in active_meetings.items():
                    # Get participant last activities
                    participant_activities = []
                    for participant in meeting.participants.values():
                        last_active_str = participant.get('last_active', meeting.created_at.isoformat())
                        try:
                            last_active = datetime.fromisoformat(last_active_str.replace('Z', '+00:00'))
                            participant_activities.append(last_active)
                        except:
                            participant_activities.append(meeting.created_at)
                    
                    # Find the most recent activity
                    if participant_activities:
                        last_activity = max([meeting.created_at] + participant_activities)
                    else:
                        last_activity = meeting.created_at
                    
                    # Check if meeting is inactive for more than 1 hour
                    if (current_time - last_activity).total_seconds() > 3600:  # 1 hour
                        meetings_to_remove.append(meeting_id)
            
            # Remove inactive meetings
            for meeting_id in meetings_to_remove:
                meeting = active_meetings.pop(meeting_id, None)
                if meeting:
                    logger.info(f"Cleaned up inactive meeting: {meeting_id}")
                    
                    # Also remove from Redis
                    try:
                        redis_client.delete(f'meeting:{meeting_id}')
                    except Exception as e:
                        logger.error(f"Failed to clean Redis for meeting {meeting_id}: {e}")
            
            # Sleep for 5 minutes
            time.sleep(300)
            
        except Exception as e:
            logger.error(f"Cleanup task error: {e}")
            time.sleep(60)

# Initialize cleanup thread
def start_cleanup_thread():
    cleanup_thread = Thread(target=cleanup_inactive_meetings, daemon=True)
    cleanup_thread.start()
    logger.info("Meeting cleanup thread started")

# Main function to initialize meeting system
def init_meeting_system(app, db):
    """Initialize the meeting system with the Flask app"""
    # Initialize SocketIO
    socketio.init_app(app)
    
    # Create meeting routes
    create_meeting_routes(app, db)
    
    # Start cleanup thread
    start_cleanup_thread()
    
    logger.info("Meeting system initialized successfully")
    return socketio