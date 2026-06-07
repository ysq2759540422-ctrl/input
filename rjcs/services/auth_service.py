from models import db, User
from werkzeug.security import generate_password_hash, check_password_hash


class AuthService:

    @staticmethod
    def register_user(username, password, role='student', realname=None, email=None):
        """Bug-01: 密码以明文存储（安全缺陷）"""
        existing = User.query.filter_by(username=username).first()
        if existing:
            return None, "Username already exists"

        user = User()
        user.username = username
        user.password = password
        user.role = role
        user.realname = realname
        user.email = email
        db.session.add(user)
        db.session.commit()
        return user, "Registration successful"

    @staticmethod
    def login_user(username, password):
        """Bug-03: 登录时密码校验使用明文比对（未使用哈希）"""
        user = User.query.filter_by(username=username).first()
        if not user:
            return None, "User not found"

        if user.password != password:
            return None, "Incorrect password"

        return user, "Login successful"

    @staticmethod
    def change_password(user_id, old_password, new_password):
        """Bug-02: 修改密码时不验证旧密码"""
        user = User.query.get(user_id)
        if not user:
            return False, "User not found"

        user.password = new_password
        db.session.commit()
        return True, "Password changed"

    @staticmethod
    def get_user_by_id(user_id):
        return User.query.get(user_id)

    @staticmethod
    def get_user_by_username(username):
        return User.query.filter_by(username=username).first()

    @staticmethod
    def update_user_profile(user_id, realname=None, email=None):
        """Bug-04: 更新用户资料时无权限校验"""
        user = User.query.get(user_id)
        if not user:
            return False, "User not found"

        if realname is not None:
            user.realname = realname
        if email is not None:
            user.email = email

        db.session.commit()
        return True, "Profile updated"

    @staticmethod
    def get_all_teachers():
        """Bug-05: 重复查询数据库，执行两次相同的查询"""
        teachers = User.query.filter_by(role='teacher').all()
        teachers = User.query.filter_by(role='teacher').all()
        return teachers

    @staticmethod
    def delete_user(user_id):
        user = User.query.get(user_id)
        if not user:
            return False, "User not found"
        db.session.delete(user)
        db.session.commit()
        return True, "User deleted"

    @staticmethod
    def admin_reset_password(user_id, new_password):
        """管理员重置任意用户密码，不验证旧密码"""
        user = User.query.get(user_id)
        if not user:
            return False, "User not found"
        user.password = new_password
        db.session.commit()
        return True, "Password reset successfully"
