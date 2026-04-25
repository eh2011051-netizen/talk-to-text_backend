# meeting_recording.py
import asyncio
import json
import os
import uuid
from datetime import datetime, timezone
from typing import Dict, List, Optional
try:
    import numpy as np
except ImportError:
    np = None
from pydub import AudioSegment
import assemblyai as aai
import google.generativeai as genai
import logging
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger(__name__)

class MeetingRecorder:
    def __init__(self, meeting_id: str, user_id: str, app_db, socketio):
        self.meeting_id = meeting_id
        self.user_id = user_id
        self.db = app_db
        self.socketio = socketio
        
        # Recording state
        self.is_recording = False
        self.recording_start_time = None
        self.audio_chunks: List[bytes] = []
        self.speaker_audio: Dict[str, List[bytes]] = {}
        self.speaker_sessions: Dict[str, Dict] = {}
        
        # Real-time processing
        self.transcriber = None
        self.realtime_config = None
        self.transcript_buffer = []
        
        # Processing queue
        self.processing_queue = asyncio.Queue()
        self.processing_task = None
        
    async def start_recording(self):
        """Start recording meeting with speaker diarization"""
        try:
            self.is_recording = True
            self.recording_start_time = datetime.now(timezone.utc)
            self.audio_chunks = []
            self.speaker_audio = {}
            self.speaker_sessions = {}
            
            # Initialize AssemblyAI real-time transcription with speaker diarization
            self.realtime_config = aai.RealtimeConfig(
                sample_rate=16000,
                encoding=aai.AudioEncoding.pcm_s16le,
                language="en",
                speaker_diarization=True,
                diarization=True,
                enable_entities=True,
                punctuate=True
            )
            
            # Create transcriber
            self.transcriber = aai.RealtimeTranscriber(
                config=self.realtime_config,
                on_data=self._on_transcription_data,
                on_error=self._on_transcription_error,
                on_open=self._on_transcription_open,
                on_close=self._on_transcription_close
            )
            
            # Start real-time transcription
            await self.transcriber.connect()
            
            # Start processing task
            self.processing_task = asyncio.create_task(self._process_audio_queue())
            
            logger.info(f"Started recording for meeting {self.meeting_id}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to start recording: {e}")
            return False
    
    async def stop_recording(self):
        """Stop recording and save meeting"""
        try:
            self.is_recording = False
            
            # Close transcription
            if self.transcriber:
                await self.transcriber.close()
            
            # Cancel processing task
            if self.processing_task:
                self.processing_task.cancel()
                try:
                    await self.processing_task
                except asyncio.CancelledError:
                    pass
            
            # Save recording
            await self._save_recording()
            
            logger.info(f"Stopped recording for meeting {self.meeting_id}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to stop recording: {e}")
            return False
    
    async def add_audio_chunk(self, audio_data: bytes, speaker_id: str, timestamp: datetime):
        """Add audio chunk from a specific speaker"""
        try:
            if not self.is_recording:
                return False
            
            # Store in speaker-specific buffer
            if speaker_id not in self.speaker_audio:
                self.speaker_audio[speaker_id] = []
            
            self.speaker_audio[speaker_id].append({
                'data': audio_data,
                'timestamp': timestamp,
                'speaker_id': speaker_id
            })
            
            # Also add to combined audio chunks
            self.audio_chunks.append({
                'data': audio_data,
                'timestamp': timestamp,
                'speaker_id': speaker_id
            })
            
            # Send to real-time transcription
            if self.transcriber:
                await self.transcriber.send(audio_data)
            
            # Add to processing queue
            await self.processing_queue.put({
                'audio_data': audio_data,
                'speaker_id': speaker_id,
                'timestamp': timestamp
            })
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to add audio chunk: {e}")
            return False
    
    async def add_video_chunk(self, video_data: bytes, speaker_id: str, timestamp: datetime):
        """Add video chunk from a specific speaker"""
        try:
            if not self.is_recording:
                return False
            
            # Initialize speaker session if not exists
            if speaker_id not in self.speaker_sessions:
                self.speaker_sessions[speaker_id] = {
                    'video_chunks': [],
                    'audio_chunks': [],
                    'speaking_time': 0,
                    'last_speaking_time': None
                }
            
            # Store video chunk
            self.speaker_sessions[speaker_id]['video_chunks'].append({
                'data': video_data,
                'timestamp': timestamp
            })
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to add video chunk: {e}")
            return False
    
    async def _process_audio_queue(self):
        """Process audio queue for real-time analysis"""
        while self.is_recording:
            try:
                # Get audio chunk from queue
                chunk = await self.processing_queue.get()
                
                # Process in background
                asyncio.create_task(self._analyze_audio_chunk(chunk))
                
                self.processing_queue.task_done()
                
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"Error processing audio queue: {e}")
    
    async def _analyze_audio_chunk(self, chunk: dict):
        """Analyze audio chunk for speaker characteristics"""
        try:
            speaker_id = chunk['speaker_id']
            
            # Update speaking time for this speaker
            if speaker_id in self.speaker_sessions:
                self.speaker_sessions[speaker_id]['speaking_time'] += 0.1  # Assuming 100ms chunks
                
                # Send real-time speaking updates via socket
                await self.socketio.emit('speaking_update', {
                    'meeting_id': self.meeting_id,
                    'speaker_id': speaker_id,
                    'speaking_time': self.speaker_sessions[speaker_id]['speaking_time'],
                    'is_speaking': True
                })
            
        except Exception as e:
            logger.error(f"Error analyzing audio chunk: {e}")
    
    def _on_transcription_data(self, transcript: aai.RealtimeTranscript):
        """Handle real-time transcription data"""
        try:
            if not transcript.text:
                return
            
            # Extract speaker information
            speaker = transcript.speaker if hasattr(transcript, 'speaker') else "Unknown"
            
            # Create caption object
            caption = {
                'id': str(uuid.uuid4()),
                'text': transcript.text,
                'speaker': speaker,
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'confidence': transcript.confidence if hasattr(transcript, 'confidence') else 0.9,
                'words': transcript.words if hasattr(transcript, 'words') else []
            }
            
            # Add to buffer
            self.transcript_buffer.append(caption)
            
            # Emit to meeting participants
            asyncio.create_task(self.socketio.emit('live_caption', {
                'meeting_id': self.meeting_id,
                'caption': caption
            }))
            
            # Perform real-time analysis
            asyncio.create_task(self._analyze_transcript(caption))
            
        except Exception as e:
            logger.error(f"Error in transcription callback: {e}")
    
    async def _analyze_transcript(self, caption: dict):
        """Analyze transcript for insights"""
        try:
            # Basic sentiment analysis
            sentiment = await self._analyze_sentiment(caption['text'])
            
            # Detect action items
            if any(keyword in caption['text'].lower() for keyword in 
                  ['need to', 'should', 'must', 'action', 'task']):
                asyncio.create_task(self.socketio.emit('action_item_detected', {
                    'meeting_id': self.meeting_id,
                    'text': caption['text'],
                    'speaker': caption['speaker'],
                    'timestamp': caption['timestamp']
                }))
            
            # Detect questions
            if '?' in caption['text']:
                asyncio.create_task(self.socketio.emit('question_detected', {
                    'meeting_id': self.meeting_id,
                    'text': caption['text'],
                    'speaker': caption['speaker'],
                    'timestamp': caption['timestamp']
                }))
            
            # Update caption with sentiment
            caption['sentiment'] = sentiment
            
        except Exception as e:
            logger.error(f"Error analyzing transcript: {e}")
    
    async def _analyze_sentiment(self, text: str) -> Dict:
        """Analyze sentiment of text"""
        try:
            # Simple rule-based sentiment analysis
            positive_words = ['good', 'great', 'excellent', 'happy', 'agree', 'yes']
            negative_words = ['bad', 'terrible', 'no', 'disagree', 'problem', 'issue']
            
            words = text.lower().split()
            positive_count = sum(1 for word in words if word in positive_words)
            negative_count = sum(1 for word in words if word in negative_words)
            
            total = len(words)
            if total == 0:
                return {'score': 0.5, 'label': 'neutral'}
            
            score = 0.5 + (positive_count - negative_count) / total * 0.5
            score = max(0, min(1, score))
            
            if score > 0.6:
                label = 'positive'
            elif score < 0.4:
                label = 'negative'
            else:
                label = 'neutral'
            
            return {'score': score, 'label': label}
            
        except Exception as e:
            logger.error(f"Error analyzing sentiment: {e}")
            return {'score': 0.5, 'label': 'neutral'}
    
    def _on_transcription_error(self, error: Exception):
        logger.error(f"Transcription error: {error}")
    
    def _on_transcription_open(self):
        logger.info(f"Transcription opened for meeting {self.meeting_id}")
    
    def _on_transcription_close(self):
        logger.info(f"Transcription closed for meeting {self.meeting_id}")
    
    async def _save_recording(self):
        """Save recording to database and filesystem"""
        try:
            # Combine audio chunks
            combined_audio = await self._combine_audio_chunks()
            
            # Save to filesystem
            upload_dir = "uploads"
            os.makedirs(upload_dir, exist_ok=True)
            
            filename = f"meeting_{self.meeting_id}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.wav"
            filepath = os.path.join(upload_dir, filename)
            
            # Save audio file
            combined_audio.export(filepath, format="wav")
            
            # Generate transcript from buffer
            transcript_text = self._generate_transcript_text()
            
            # Save to database
            from app import Meeting, log_activity
            
            meeting_record = Meeting(
                user_id=self.user_id,
                title=f"Meeting Recording - {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M')}",
                filename=filename,
                language='en',
                transcript_language='en',
                status='uploaded',
                upload_date=datetime.now(timezone.utc)
            )
            
            self.db.session.add(meeting_record)
            self.db.session.commit()
            
            # Store initial transcription
            meeting_record.transcription = json.dumps({
                'raw': transcript_text,
                'translated': transcript_text,
                'optimized': transcript_text,
                'speakers': self._extract_speaker_info(),
                'timestamps': self._extract_timestamps()
            })
            
            self.db.session.commit()
            
            # Log activity
            log_activity(
                user_id=self.user_id,
                activity_type='meeting',
                title=f"Meeting Recording: {meeting_record.title}",
                description=f"Meeting recorded with {len(self.speaker_sessions)} participants",
                meeting_id=meeting_record.id,
                metadata={
                    'meeting_id': self.meeting_id,
                    'participant_count': len(self.speaker_sessions),
                    'recording_duration': (datetime.now(timezone.utc) - self.recording_start_time).total_seconds(),
                    'transcript_length': len(transcript_text)
                }
            )
            
            # Start processing in background
            await self._start_processing(meeting_record.id)
            
            logger.info(f"Recording saved: {filepath}")
            return meeting_record.id
            
        except Exception as e:
            logger.error(f"Failed to save recording: {e}")
            return None
    
    async def _combine_audio_chunks(self) -> AudioSegment:
        """Combine all audio chunks into single audio file"""
        try:
            combined = AudioSegment.empty()
            
            for chunk in self.audio_chunks:
                # Convert bytes to AudioSegment
                audio_segment = AudioSegment(
                    data=chunk['data'],
                    sample_width=2,
                    frame_rate=16000,
                    channels=1
                )
                combined += audio_segment
            
            return combined
            
        except Exception as e:
            logger.error(f"Failed to combine audio chunks: {e}")
            return AudioSegment.silent(duration=1000)
    
    def _generate_transcript_text(self) -> str:
        """Generate transcript text from buffer"""
        try:
            # Sort by timestamp
            sorted_transcripts = sorted(self.transcript_buffer, 
                                      key=lambda x: x.get('timestamp', ''))
            
            # Format as readable transcript
            transcript_lines = []
            for caption in sorted_transcripts:
                speaker = caption.get('speaker', 'Unknown')
                text = caption.get('text', '')
                timestamp = caption.get('timestamp', '')
                
                # Format timestamp
                try:
                    dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                    time_str = dt.strftime('%H:%M:%S')
                except:
                    time_str = timestamp
                
                transcript_lines.append(f"[{time_str}] {speaker}: {text}")
            
            return '\n'.join(transcript_lines)
            
        except Exception as e:
            logger.error(f"Failed to generate transcript: {e}")
            return ""
    
    def _extract_speaker_info(self) -> Dict:
        """Extract speaker information"""
        speaker_info = {}
        
        for speaker_id, session in self.speaker_sessions.items():
            speaker_info[speaker_id] = {
                'speaking_time': session.get('speaking_time', 0),
                'audio_chunks': len(session.get('audio_chunks', [])),
                'video_chunks': len(session.get('video_chunks', []))
            }
        
        return speaker_info
    
    def _extract_timestamps(self) -> List[Dict]:
        """Extract timestamps for transcript"""
        timestamps = []
        
        for caption in self.transcript_buffer:
            timestamps.append({
                'text': caption.get('text', ''),
                'speaker': caption.get('speaker', 'Unknown'),
                'timestamp': caption.get('timestamp', ''),
                'confidence': caption.get('confidence', 0)
            })
        
        return timestamps
    
    async def _start_processing(self, meeting_id: int):
        """Start processing the recorded meeting"""
        try:
            from app import executor
            
            # Submit processing task to thread pool
            executor.submit(self._process_meeting, meeting_id)
            
            logger.info(f"Started processing for meeting {meeting_id}")
            
        except Exception as e:
            logger.error(f"Failed to start processing: {e}")
    
    def _process_meeting(self, meeting_id: int):
        """Process meeting recording (runs in background thread)"""
        try:
            # Import here to avoid circular imports
            from app import start_processing
            
            # Call existing processing function
            start_processing(meeting_id)
            
        except Exception as e:
            logger.error(f"Failed to process meeting: {e}")


class MeetingRecordingManager:
    """Manager for all meeting recordings"""
    
    def __init__(self):
        self.recordings: Dict[str, MeetingRecorder] = {}
        self.lock = asyncio.Lock()
    
    async def start_recording(self, meeting_id: str, user_id: str, db, socketio) -> MeetingRecorder:
        """Start recording a meeting"""
        async with self.lock:
            if meeting_id in self.recordings:
                return self.recordings[meeting_id]
            
            recorder = MeetingRecorder(meeting_id, user_id, db, socketio)
            await recorder.start_recording()
            
            self.recordings[meeting_id] = recorder
            return recorder
    
    async def stop_recording(self, meeting_id: str) -> bool:
        """Stop recording a meeting"""
        async with self.lock:
            if meeting_id not in self.recordings:
                return False
            
            recorder = self.recordings[meeting_id]
            success = await recorder.stop_recording()
            
            if success:
                del self.recordings[meeting_id]
            
            return success
    
    async def add_audio_chunk(self, meeting_id: str, audio_data: bytes, speaker_id: str, timestamp: datetime) -> bool:
        """Add audio chunk to meeting recording"""
        async with self.lock:
            if meeting_id not in self.recordings:
                return False
            
            recorder = self.recordings[meeting_id]
            return await recorder.add_audio_chunk(audio_data, speaker_id, timestamp)
    
    async def add_video_chunk(self, meeting_id: str, video_data: bytes, speaker_id: str, timestamp: datetime) -> bool:
        """Add video chunk to meeting recording"""
        async with self.lock:
            if meeting_id not in self.recordings:
                return False
            
            recorder = self.recordings[meeting_id]
            return await recorder.add_video_chunk(video_data, speaker_id, timestamp)
    
    def get_recorder(self, meeting_id: str) -> Optional[MeetingRecorder]:
        """Get recorder for meeting"""
        return self.recordings.get(meeting_id)


# Global instance
recording_manager = MeetingRecordingManager()