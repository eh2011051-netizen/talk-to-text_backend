from flask import Blueprint, request, jsonify
from flask_jwt_extended import jwt_required, get_jwt_identity
from datetime import datetime

friends_bp = Blueprint('friends', __name__)

@friends_bp.route('/api/friends', methods=['GET'])
@jwt_required()
def get_friends():
    from app import db, User, Friendship, Message
    import os
    user_id = int(get_jwt_identity())
    # Build a base URL for serving stored images (falls back to localhost for dev)
    api_base = os.environ.get('NEXT_PUBLIC_API_URL', os.environ.get('API_BASE_URL', 'http://localhost:5000'))

    def _build_image_url(img):
        """Return a full URL for stored user/group images."""
        if not img:
            return None
        # Already a data-URI, blob or external URL â€” return as-is
        if img.startswith('data:') or img.startswith('blob:') or img.startswith('http'):
            return img
        # Server-relative path like /uploads/avatars/xyz.jpg
        return f"{api_base}{img}"

    friendships = Friendship.query.filter(
        ((Friendship.user_id == user_id) | (Friendship.friend_id == user_id)),
        Friendship.is_deleted == False
    ).all()
    
    friends = []
    for f in friendships:
        other_id = f.friend_id if f.user_id == user_id else f.user_id
        other_user = User.query.get(other_id)
        if other_user:
            # Check for unread count
            unread_count = Message.query.filter_by(
                sender_id=other_id, 
                receiver_id=user_id, 
                is_read=False
            ).count()
            
            friends.append({
                "id": other_user.id,
                "name": other_user.full_name,
                "email": other_user.email,
                "image": _build_image_url(other_user.image),
                "isFriend": True,
                "status": other_user.status,
                "lastSeen": other_user.last_seen.isoformat() if other_user.last_seen else None,
                "bio": other_user.bio,
                "isPinned": f.is_pinned,
                "isBlocked": f.is_blocked and f.blocked_by_id == user_id,
                "hasBlockedMe": f.is_blocked and f.blocked_by_id != user_id,
                "isMuted": f.is_muted,
                "isArchived": f.is_archived,
                "isFavourite": f.is_favourite,
                "unreadCount": unread_count
            })
    return jsonify(friends)

@friends_bp.route('/api/groups', methods=['GET'])
@jwt_required()
def get_groups():
    from app import Group, GroupMember
    user_id = int(get_jwt_identity())
    # Get all groups where the user is a member
    memberships = GroupMember.query.filter_by(user_id=user_id).all()
    
    groups = []
    for m in memberships:
        group = Group.query.get(m.group_id)
        if group:
            # Get members for the group
            group_members = GroupMember.query.filter_by(group_id=group.id).all()
            member_ids = [gm.user_id for gm in group_members if not gm.is_exited]
            admin_ids = [gm.user_id for gm in group_members if gm.role == 'admin' and not gm.is_exited]
            
            # The current user's membership
            my_membership = next((gm for gm in group_members if gm.user_id == user_id), None)
            is_exited = my_membership.is_exited if my_membership else False
            
            groups.append({
                "id": group.id,
                "name": group.name,
                "image": group.image,
                "isGroup": True,
                "bio": f"{len(group_members)} members",
                "description": group.description,
                "memberIds": member_ids,
                "adminIds": admin_ids,
                "creatorId": group.created_by_id,
                "createdAt": group.created_at.isoformat(),
                "isExited": is_exited,
                "groupSettings": {
                    "onlyAdminsCanEditInfo": False,
                    "onlyAdminsCanAddMembers": False,
                    "onlyAdminsCanSendMessages": False
                }
            })
    return jsonify(groups)

@friends_bp.route('/api/groups', methods=['POST'])
@jwt_required()
def create_group():
    from app import db, Group, GroupMember, GroupInvite
    user_id = int(get_jwt_identity())
    data = request.json
    name = data.get('name')
    image = data.get('image')
    description = data.get('description', '')
    member_ids = data.get('member_ids', [])
    
    if not name:
        return jsonify({"error": "Group name is required"}), 400
        
    new_group = Group(
        name=name,
        image=image,
        description=description,
        creator_id=user_id
    )
    db.session.add(new_group)
    db.session.flush() # Get ID before commit
    
    # Add creator as admin
    db.session.add(GroupMember(group_id=new_group.id, user_id=user_id, role='admin'))
    
    # Send invites to other members instead of adding them directly
    for m_id in member_ids:
        if m_id != user_id:
            existing_invite = GroupInvite.query.filter_by(group_id=new_group.id, invitee_id=m_id, status='pending').first()
            if not existing_invite:
                db.session.add(GroupInvite(
                    group_id=new_group.id,
                    inviter_id=user_id,
                    invitee_id=m_id,
                    status='pending'
                ))
            
    db.session.commit()
    
    return jsonify({
        "id": new_group.id,
        "name": new_group.name,
        "success": True
    })

@friends_bp.route('/api/groups/invites', methods=['GET'])
@jwt_required()
def get_group_invites():
    from app import Group, GroupInvite, GroupMember, User
    user_id = int(get_jwt_identity())
    invites = GroupInvite.query.filter_by(invitee_id=user_id, status='pending').all()

    result = []
    for i in invites:
        group = Group.query.get(i.group_id)
        inviter = User.query.get(i.inviter_id)
        if not group or not inviter:
            continue
        member_count = GroupMember.query.filter_by(group_id=group.id).count()
        result.append({
            "id": i.id,
            "groupId": group.id,
            "groupName": group.name,
            "groupImage": group.image,
            "groupDescription": group.description,
            "inviterName": inviter.full_name,
            "inviterImage": inviter.image,
            "memberCount": member_count,
            "timestamp": i.created_at.isoformat()
        })
    return jsonify(result)


@friends_bp.route('/api/groups/<int:group_id>/invite', methods=['POST'])
@jwt_required()
def send_group_invite(group_id):
    from app import db, Group, GroupMember, GroupInvite, User
    inviter_id = int(get_jwt_identity())
    data = request.json
    user_ids = data.get('user_ids', [])

    if not user_ids:
        return jsonify({"error": "No users specified"}), 400

    group = Group.query.get(group_id)
    if not group:
        return jsonify({"error": "Group not found"}), 404

    # Check inviter is a member of the group
    inviter_membership = GroupMember.query.filter_by(group_id=group_id, user_id=inviter_id).first()
    if not inviter_membership:
        return jsonify({"error": "You are not a member of this group"}), 403

    sent = []
    skipped = []
    for uid in user_ids:
        uid = int(uid)
        # Skip if already a member
        if GroupMember.query.filter_by(group_id=group_id, user_id=uid).first():
            skipped.append({"userId": uid, "reason": "already_member"})
            continue
        # Skip if already has a pending invite
        if GroupInvite.query.filter_by(group_id=group_id, invitee_id=uid, status='pending').first():
            skipped.append({"userId": uid, "reason": "already_invited"})
            continue
        invite = GroupInvite(
            group_id=group_id,
            inviter_id=inviter_id,
            invitee_id=uid,
            status='pending',
            created_at=datetime.utcnow()
        )
        db.session.add(invite)
        sent.append(uid)

    db.session.commit()
    return jsonify({"success": True, "sent": sent, "skipped": skipped})


@friends_bp.route('/api/groups/invites/<int:invite_id>/accept', methods=['POST'])
@jwt_required()
def accept_group_invite(invite_id):
    from app import db, GroupInvite, GroupMember
    user_id = int(get_jwt_identity())
    invite = GroupInvite.query.get(invite_id)

    if not invite or invite.invitee_id != user_id:
        return jsonify({"error": "Invite not found"}), 404
    if invite.status != 'pending':
        return jsonify({"error": "Invite already handled"}), 400

    invite.status = 'accepted'
    # Add as group member if not already
    existing = GroupMember.query.filter_by(group_id=invite.group_id, user_id=user_id).first()
    if not existing:
        db.session.add(GroupMember(group_id=invite.group_id, user_id=user_id, role='member'))
    db.session.commit()
    return jsonify({"success": True, "groupId": invite.group_id})


@friends_bp.route('/api/groups/invites/<int:invite_id>/reject', methods=['POST'])
@jwt_required()
def reject_group_invite(invite_id):
    from app import db, GroupInvite
    user_id = int(get_jwt_identity())
    invite = GroupInvite.query.get(invite_id)

    if not invite or invite.invitee_id != user_id:
        return jsonify({"error": "Invite not found"}), 404
    if invite.status != 'pending':
        return jsonify({"error": "Invite already handled"}), 400

    invite.status = 'rejected'
    db.session.commit()
    return jsonify({"success": True})


@friends_bp.route('/api/groups/<int:group_id>/pending-invites', methods=['GET'])
@jwt_required()
def get_pending_invites_for_group(group_id):
    from app import GroupInvite, GroupMember
    user_id = int(get_jwt_identity())
    # Only members can query this
    if not GroupMember.query.filter_by(group_id=group_id, user_id=user_id).first():
        return jsonify({"error": "Not a member"}), 403
    pending = GroupInvite.query.filter_by(group_id=group_id, status='pending').all()
    return jsonify({"pendingUserIds": [i.invitee_id for i in pending]})

@friends_bp.route('/api/groups/<int:group_id>/image', methods=['POST'])
@jwt_required()
def update_group_image(group_id):
    from app import db, Group, GroupMember
    user_id = int(get_jwt_identity())
    data = request.json
    image_url = data.get('image')
    
    membership = GroupMember.query.filter_by(group_id=group_id, user_id=user_id, role='admin').first()
    if not membership:
        return jsonify({"error": "Only admins can change group image"}), 403
        
    group = Group.query.get(group_id)
    group.image = image_url
    db.session.commit()
    
    return jsonify({"success": True})

@friends_bp.route('/api/calls/log', methods=['GET'])
@jwt_required()
def get_call_logs():
    from app import User, CallLog
    user_id = int(get_jwt_identity())
    logs = CallLog.query.filter(
        (CallLog.user_id == user_id)
    ).order_by(CallLog.timestamp.desc()).all()
    
    return jsonify([{
        "id": l.id,
        "name": User.query.get(l.other_user_id).full_name if l.other_user_id else "Unknown",
        "image": User.query.get(l.other_user_id).image if l.other_user_id else None,
        "type": l.type,
        "isVideo": l.is_video,
        "timestamp": l.timestamp.isoformat(),
        "duration": l.duration
    } for l in logs])

@friends_bp.route('/api/broadcast-lists', methods=['GET'])
@jwt_required()
def get_broadcast_lists():
    from app import BroadcastList, BroadcastRecipient
    user_id = int(get_jwt_identity())
    lists = BroadcastList.query.filter_by(user_id=user_id).all()
    
    result = []
    for l in lists:
        recipients = BroadcastRecipient.query.filter_by(list_id=l.id).all()
        result.append({
            "id": l.id,
            "name": l.name,
            "recipients": [r.recipient_id for r in recipients],
            "lastUsed": l.last_used.isoformat() if l.last_used else None
        })
    return jsonify(result)

@friends_bp.route('/api/friends/requests', methods=['GET'])
@jwt_required()
def get_requests():
    from app import User, FriendRequest
    user_id = int(get_jwt_identity())
    incoming = FriendRequest.query.filter_by(receiver_id=user_id, status='pending').all()
    outgoing = FriendRequest.query.filter_by(sender_id=user_id, status='pending').all()
    
    return jsonify({
        "incoming": [{
            "id": r.id,
            "sender": {
                "id": r.sender_id,
                "name": User.query.get(r.sender_id).full_name,
                "email": User.query.get(r.sender_id).email,
                "image": User.query.get(r.sender_id).image
            },
            "created_at": r.created_at.isoformat()
        } for r in incoming],
        "outgoing": [{
            "id": r.id,
            "receiver": {
                "id": r.receiver_id,
                "name": User.query.get(r.receiver_id).full_name,
                "email": User.query.get(r.receiver_id).email,
                "image": User.query.get(r.receiver_id).image
            },
            "created_at": r.created_at.isoformat()
        } for r in outgoing]
    })

@friends_bp.route('/api/friends/request', methods=['POST'])
@jwt_required()
def send_request():
    from app import db, Friendship, FriendRequest
    user_id = int(get_jwt_identity())
    data = request.json
    receiver_id = data.get('receiver_id')
    
    if not receiver_id:
        return jsonify({"error": "Missing receiver_id"}), 400
        
    if receiver_id == user_id:
        return jsonify({"error": "Cannot send request to yourself"}), 400
        
    # Check if already friends
    existing_friendship = Friendship.query.filter(
        ((Friendship.user_id == user_id) & (Friendship.friend_id == receiver_id)) |
        ((Friendship.user_id == receiver_id) & (Friendship.friend_id == user_id))
    ).first()
    
    if existing_friendship:
        return jsonify({"error": "Already friends"}), 400
        
    # Check if already pending
    existing_request = FriendRequest.query.filter_by(
        sender_id=user_id, receiver_id=receiver_id, status='pending'
    ).first()
    
    if existing_request:
        return jsonify({"error": "Request already pending"}), 400
        
    new_request = FriendRequest(sender_id=user_id, receiver_id=receiver_id)
    db.session.add(new_request)
    db.session.commit()
    return jsonify({"success": True})

@friends_bp.route('/api/friends/request/accept', methods=['POST'])
@jwt_required()
def accept_request():
    from app import db, Friendship, FriendRequest
    user_id = int(get_jwt_identity())
    data = request.json
    request_id = data.get('request_id')
    
    req = FriendRequest.query.get(request_id)
    if not req or req.receiver_id != user_id or req.status != 'pending':
        return jsonify({"error": "Request not found or invalid"}), 404
        
    req.status = 'accepted'
    
    # Create friendship
    friendship = Friendship(user_id=req.sender_id, friend_id=req.receiver_id)
    db.session.add(friendship)
    db.session.commit()
    return jsonify({"success": True})

@friends_bp.route('/api/friends/request/decline', methods=['POST'])
@jwt_required()
def decline_request():
    from app import db, FriendRequest
    user_id = int(get_jwt_identity())
    data = request.json
    request_id = data.get('request_id')
    
    req = FriendRequest.query.get(request_id)
    if not req or req.receiver_id != user_id or req.status != 'pending':
        return jsonify({"error": "Request not found or invalid"}), 404
        
    req.status = 'rejected'
    db.session.commit()
    return jsonify({"success": True})

@friends_bp.route('/api/friends/request/cancel', methods=['POST'])
@jwt_required()
def cancel_request():
    from app import db, FriendRequest
    user_id = int(get_jwt_identity())
    data = request.json
    request_id = data.get('request_id')
    
    req = FriendRequest.query.get(request_id)
    if not req or req.sender_id != user_id or req.status != 'pending':
        return jsonify({"error": "Request not found or invalid"}), 404
        
    db.session.delete(req)
    db.session.commit()
    return jsonify({"success": True})

@friends_bp.route('/api/users/search', methods=['GET'])
@jwt_required()
def search_users():
    from app import User, Friendship, FriendRequest
    q = request.args.get('q', '')
    if not q or len(q) < 3:
        return jsonify([])
    
    users = User.query.filter(
        (User.email.ilike(f"%{q}%")) | (User.full_name.ilike(f"%{q}%"))
    ).limit(10).all()
    
    user_id = int(get_jwt_identity())
    result = []
    for u in users:
        if u.id == user_id:
            result.append({
                "id": u.id,
                "name": u.full_name,
                "email": u.email,
                "image": u.image,
                "isFriend": True, # You are your own friend
                "requestPending": False,
                "bio": u.bio,
                "isSelf": True
            })
            continue
            
        # Check if already friends or requested
        friendship = Friendship.query.filter(
            ((Friendship.user_id == user_id) & (Friendship.friend_id == u.id)) |
            ((Friendship.user_id == u.id) & (Friendship.friend_id == user_id))
        ).first()
        
        request_pending = FriendRequest.query.filter_by(
            sender_id=user_id, receiver_id=u.id, status='pending'
        ).first() is not None
        
        result.append({
            "id": u.id,
            "name": u.full_name,
            "email": u.email,
            "image": u.image,
            "isFriend": friendship is not None,
            "requestPending": request_pending,
            "bio": u.bio
        })
    return jsonify(result)

@friends_bp.route('/api/messages/<int:other_user_id>', methods=['GET'])
@jwt_required()
def get_messages(other_user_id):
    from app import db, Message
    user_id = int(get_jwt_identity())
    messages = Message.query.filter(
        ((Message.sender_id == user_id) & (Message.receiver_id == other_user_id)) |
        ((Message.sender_id == other_user_id) & (Message.receiver_id == user_id))
    ).order_by(Message.timestamp.asc()).all()
    
    # Mark as read
    Message.query.filter_by(sender_id=other_user_id, receiver_id=user_id, is_read=False).update({"is_read": True})
    db.session.commit()
    
    def _serialize_msg(m):
        reply_data = None
        if m.reply_to_id:
            rm = db.session.get(Message, m.reply_to_id)
            if rm:
                reply_data = {
                    "id": rm.id,
                    "text": rm.text,
                    "sender": "me" if rm.sender_id == user_id else "other",
                    "type": rm.type,
                    "mediaUrl": rm.media_url
                }
        return {
            "id": m.id,
            "text": m.text,
            "sender": "me" if m.sender_id == user_id else "other",
            "time": m.timestamp.isoformat() + 'Z',
            "type": m.type,
            "mediaUrl": m.media_url,
            "isRead": m.is_read,
            "replyTo": reply_data,
            "reply_to_id": m.reply_to_id
        }

    return jsonify([_serialize_msg(m) for m in messages])

@friends_bp.route('/api/messages/group/<int:group_id>', methods=['GET'])
@jwt_required()
def get_group_messages(group_id):
    from app import Message, GroupMember
    user_id = int(get_jwt_identity())
    # Check if user is a member
    from app import GroupMember
    membership = GroupMember.query.filter_by(group_id=group_id, user_id=user_id).first()
    if not membership:
        return jsonify({"error": "Not a member of this group"}), 403
        
    messages = Message.query.filter_by(group_id=group_id).order_by(Message.timestamp.asc()).all()
    
    def _serialize_group_msg(m):
        reply_data = None
        if m.reply_to_id:
            rm = Message.query.get(m.reply_to_id)
            if rm:
                reply_data = {
                    "id": rm.id,
                    "text": rm.text,
                    "sender": "me" if rm.sender_id == user_id else "other",
                    "type": rm.type,
                    "mediaUrl": rm.media_url
                }
        return {
            "id": m.id,
            "sender_id": m.sender_id,
            "sender": "me" if m.sender_id == user_id else "other",
            "text": m.text,
            "type": m.type,
            "mediaUrl": m.media_url,
            "time": m.timestamp.isoformat() + 'Z',
            "isRead": m.is_read,
            "replyTo": reply_data,
            "reply_to_id": m.reply_to_id
        }

    return jsonify([_serialize_group_msg(m) for m in messages])

@friends_bp.route('/api/messages/send', methods=['POST'])
@jwt_required()
def send_message():
    from app import db, Message
    user_id = int(get_jwt_identity())
    data = request.json
    text = data.get('text')
    receiver_id = data.get('receiver_id')
    group_id = data.get('group_id')
    msg_type = data.get('type', 'text')
    media_url = data.get('mediaUrl')
    reply_to_id = data.get('reply_to_id')
    
    if not receiver_id and not group_id:
        return jsonify({"error": "Missing receiver_id or group_id"}), 400
        
    msg = Message(
        sender_id=user_id,
        receiver_id=receiver_id,
        group_id=group_id,
        text=text,
        type=msg_type,
        media_url=media_url,
        reply_to_id=reply_to_id
    )
    db.session.add(msg)
    db.session.commit()
    
    # If there's a soft-deleted friendship, restore it
    if receiver_id:
        from app import Friendship
        friendship = Friendship.query.filter(
            ((Friendship.user_id == user_id) & (Friendship.friend_id == receiver_id)) |
            ((Friendship.user_id == receiver_id) & (Friendship.friend_id == user_id))
        ).first()
        if friendship and friendship.is_deleted:
            friendship.is_deleted = False
            db.session.commit()

    reply_data = None
    if msg.reply_to_id:
        rm = db.session.get(Message, msg.reply_to_id)
        if rm:
            reply_data = {
                "id": rm.id,
                "text": rm.text,
                "sender": "me" if rm.sender_id == user_id else "other",
                "type": rm.type,
                "mediaUrl": rm.media_url
            }

    return jsonify({
        "id": msg.id,
        "text": msg.text,
        "sender": "me",
        "time": msg.timestamp.isoformat() + 'Z',
        "type": msg.type,
        "mediaUrl": msg.media_url,
        "replyTo": reply_data,
        "reply_to_id": msg.reply_to_id
    })

# â”€â”€ DELETE /api/friends/<id> â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@friends_bp.route('/api/friends/<int:friend_id>', methods=['DELETE'])
@jwt_required()
def delete_friend(friend_id):
    """Permanently remove a friendship (delete the DB row)."""
    from app import db, Friendship
    user_id = int(get_jwt_identity())

    friendship = Friendship.query.filter(
        ((Friendship.user_id == user_id) & (Friendship.friend_id == friend_id)) |
        ((Friendship.user_id == friend_id) & (Friendship.friend_id == user_id))
    ).first()

    if friendship:
        friendship.is_deleted = True
        db.session.commit()
    return jsonify({"success": True})

# â”€â”€ POST /api/friends/block â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@friends_bp.route('/api/friends/block', methods=['POST'])
@jwt_required()
def block_user():
    """Block a user. Creates a friendship record if one does not exist."""
    from app import db, Friendship
    user_id = int(get_jwt_identity())
    data = request.json
    target_id = data.get('user_id')

    if not target_id:
        return jsonify({"error": "user_id is required"}), 400

    target_id = int(target_id)

    friendship = Friendship.query.filter(
        ((Friendship.user_id == user_id) & (Friendship.friend_id == target_id)) |
        ((Friendship.user_id == target_id) & (Friendship.friend_id == user_id))
    ).first()

    if not friendship:
        # No existing relationship - create a blocked entry
        friendship = Friendship(
            user_id=user_id,
            friend_id=target_id,
            is_blocked=True,
            blocked_by_id=user_id
        )
        db.session.add(friendship)
    else:
        # Restore soft-deleted friendship then block
        friendship.is_deleted = False
        friendship.is_blocked = True
        friendship.blocked_by_id = user_id

    db.session.commit()
    return jsonify({"success": True})

# â”€â”€ POST /api/friends/unblock â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@friends_bp.route('/api/friends/unblock', methods=['POST'])
@jwt_required()
def unblock_user():
    """Unblock a previously blocked friend."""
    from app import db, Friendship
    user_id = int(get_jwt_identity())
    data = request.json
    target_id = data.get('user_id')

    if not target_id:
        return jsonify({"error": "user_id is required"}), 400

    friendship = Friendship.query.filter(
        ((Friendship.user_id == user_id) & (Friendship.friend_id == target_id)) |
        ((Friendship.user_id == target_id) & (Friendship.friend_id == user_id))
    ).first()

    if not friendship or friendship.blocked_by_id != user_id:
        return jsonify({"error": "Not blocked by you or friendship not found"}), 404

    friendship.is_blocked = False
    friendship.blocked_by_id = None
    db.session.commit()
    return jsonify({"success": True})

# â”€â”€ DELETE /api/messages/<other_user_id>/clear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@friends_bp.route('/api/messages/<int:other_user_id>/clear', methods=['DELETE'])
@jwt_required()
def clear_chat(other_user_id):
    """Delete all messages between current user and another user."""
    from app import db, Message
    user_id = int(get_jwt_identity())

    Message.query.filter(
        ((Message.sender_id == user_id) & (Message.receiver_id == other_user_id)) |
        ((Message.sender_id == other_user_id) & (Message.receiver_id == user_id))
    ).delete(synchronize_session=False)

    db.session.commit()
    return jsonify({"success": True})

# â”€â”€ DELETE /api/messages/group/<group_id>/clear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@friends_bp.route('/api/messages/group/<int:group_id>/clear', methods=['DELETE'])
@jwt_required()
def clear_group_chat(group_id):
    """Delete all messages in a group chat (only group admins)."""
    from app import db, Message, GroupMember
    user_id = int(get_jwt_identity())

    membership = GroupMember.query.filter_by(group_id=group_id, user_id=user_id).first()
    if not membership:
        return jsonify({"error": "Not a member of this group"}), 403

    Message.query.filter_by(group_id=group_id).delete(synchronize_session=False)
    db.session.commit()
    return jsonify({"success": True})


# â”€â”€ WebRTC Call Signaling (in-memory for real-time, DB for logs) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
import threading as _threading
_active_calls = {}   # call_id -> {offer, answer, ice_caller, ice_callee, state, ...}
_pending_signals = {}  # user_id -> list of signal dicts
_calls_lock = _threading.Lock()

def _push_signal(user_id: int, signal: dict):
    with _calls_lock:
        _pending_signals.setdefault(user_id, []).append(signal)

@friends_bp.route('/api/calls/initiate', methods=['POST'])
@jwt_required()
def call_initiate():
    """Caller initiates a call â†’ push ringing signal to callee."""
    from app import db, CallLog, User
    caller_id = int(get_jwt_identity())
    data = request.json or {}
    callee_id = int(data.get('callee_id', 0))
    is_video = bool(data.get('is_video', False))
    call_id = data.get('call_id', f"{caller_id}_{callee_id}_{int(datetime.utcnow().timestamp())}")

    callee = db.session.get(User, callee_id)
    caller = db.session.get(User, caller_id)
    if not callee or not caller:
        return jsonify({"error": "User not found"}), 404

    with _calls_lock:
        _active_calls[call_id] = {
            'caller_id': caller_id,
            'callee_id': callee_id,
            'is_video': is_video,
            'state': 'ringing',
            'offer': None,
            'answer': None,
            'ice_caller': [],
            'ice_callee': [],
            'started_at': datetime.utcnow(),
            'connected_at': None
        }

    _push_signal(callee_id, {
        'type': 'incoming_call',
        'call_id': call_id,
        'caller_id': caller_id,
        'caller_name': caller.full_name,
        'caller_image': caller.image,
        'is_video': is_video
    })

    return jsonify({"call_id": call_id, "status": "ringing"})


@friends_bp.route('/api/calls/offer', methods=['POST'])
@jwt_required()
def call_offer():
    """Caller sends WebRTC SDP offer."""
    caller_id = int(get_jwt_identity())
    data = request.json or {}
    call_id = data.get('call_id')
    offer = data.get('offer')

    with _calls_lock:
        call = _active_calls.get(call_id)
        if not call or call['caller_id'] != caller_id:
            return jsonify({"error": "Call not found"}), 404
        call['offer'] = offer

    _push_signal(call['callee_id'], {'type': 'offer', 'call_id': call_id, 'offer': offer})
    return jsonify({"success": True})


@friends_bp.route('/api/calls/answer', methods=['POST'])
@jwt_required()
def call_answer():
    """Callee answers (accept or reject)."""
    from app import db, CallLog, User
    callee_id = int(get_jwt_identity())
    data = request.json or {}
    call_id = data.get('call_id')
    accepted = bool(data.get('accepted', False))
    answer = data.get('answer')  # SDP answer

    with _calls_lock:
        call = _active_calls.get(call_id)
        if not call or call['callee_id'] != callee_id:
            return jsonify({"error": "Call not found"}), 404

        if accepted:
            call['state'] = 'connected'
            call['connected_at'] = datetime.utcnow()
            call['answer'] = answer
        else:
            call['state'] = 'rejected'

    caller_id = call['caller_id']

    if accepted:
        _push_signal(caller_id, {'type': 'answer', 'call_id': call_id, 'answer': answer})
    else:
        _push_signal(caller_id, {'type': 'call_rejected', 'call_id': call_id})
        # Log as missed for caller
        _log_call(db, caller_id, callee_id, 'missed', call['is_video'], 0)
        with _calls_lock:
            _active_calls.pop(call_id, None)

    return jsonify({"success": True})


@friends_bp.route('/api/calls/ice', methods=['POST'])
@jwt_required()
def call_ice():
    """Send ICE candidate to remote peer."""
    user_id = int(get_jwt_identity())
    data = request.json or {}
    call_id = data.get('call_id')
    candidate = data.get('candidate')

    with _calls_lock:
        call = _active_calls.get(call_id)
        if not call:
            return jsonify({"error": "Call not found"}), 404
        is_caller = call['caller_id'] == user_id
        remote_id = call['callee_id'] if is_caller else call['caller_id']

    _push_signal(remote_id, {'type': 'ice_candidate', 'call_id': call_id, 'candidate': candidate})
    return jsonify({"success": True})


@friends_bp.route('/api/calls/end', methods=['POST'])
@jwt_required()
def call_end():
    """End a call, log it, push hangup signal."""
    from app import db
    user_id = int(get_jwt_identity())
    data = request.json or {}
    call_id = data.get('call_id')

    with _calls_lock:
        call = _active_calls.pop(call_id, None)

    if not call:
        return jsonify({"success": True})  # Already ended

    caller_id = call['caller_id']
    callee_id = call['callee_id']
    is_video = call['is_video']
    duration = 0
    if call['connected_at']:
        duration = int((datetime.utcnow() - call['connected_at']).total_seconds())

    # Notify remote peer
    remote_id = callee_id if user_id == caller_id else caller_id
    _push_signal(remote_id, {'type': 'call_ended', 'call_id': call_id, 'duration': duration})

    # Log call for both parties
    call_type_caller = 'outgoing'
    call_type_callee = 'incoming' if duration > 0 else 'missed'

    _log_call(db, caller_id, callee_id, call_type_caller, is_video, duration)
    _log_call(db, callee_id, caller_id, call_type_callee, is_video, duration)

    return jsonify({"success": True, "duration": duration})


@friends_bp.route('/api/calls/signals', methods=['GET'])
@jwt_required()
def get_signals():
    """Long-poll endpoint: returns pending signals for this user (drains queue)."""
    user_id = int(get_jwt_identity())
    with _calls_lock:
        signals = _pending_signals.pop(user_id, [])
    return jsonify(signals)


@friends_bp.route('/api/calls/log', methods=['POST'])
@jwt_required()
def log_call_manual():
    """Manually log a call (fallback)."""
    from app import db
    user_id = int(get_jwt_identity())
    data = request.json or {}
    other_id = data.get('other_user_id')
    call_type = data.get('type', 'outgoing')
    is_video = bool(data.get('is_video', False))
    duration = int(data.get('duration', 0))
    _log_call(db, user_id, other_id, call_type, is_video, duration)
    return jsonify({"success": True})


def _log_call(db, user_id: int, other_user_id: int, call_type: str, is_video: bool, duration: int):
    from app import CallLog
    try:
        log = CallLog(
            user_id=user_id,
            other_user_id=other_user_id,
            type=call_type,
            is_video=is_video,
            duration=duration
        )
        db.session.add(log)
        db.session.commit()
    except Exception as e:
        db.session.rollback()