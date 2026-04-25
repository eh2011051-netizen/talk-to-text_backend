
from app import app, db, User, Friendship

def fix_self_friendships():
    with app.app_context():
        users = User.query.all()
        fixed_count = 0
        for user in users:
            # Check if self-friendship already exists
            existing = Friendship.query.filter_by(user_id=user.id, friend_id=user.id).first()
            if not existing:
                print(f"Adding self-friendship for user: {user.email}")
                self_friendship = Friendship(user_id=user.id, friend_id=user.id)
                db.session.add(self_friendship)
                fixed_count += 1
        
        if fixed_count > 0:
            db.session.commit()
            print(f"Successfully added self-friendship for {fixed_count} users.")
        else:
            print("All users already have self-friendships.")

if __name__ == "__main__":
    fix_self_friendships()
