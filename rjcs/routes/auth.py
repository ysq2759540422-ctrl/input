from flask import Blueprint, request, session, redirect, url_for, render_template
from services.auth_service import AuthService
from utils.helpers import login_required, admin_required, format_response

auth_bp = Blueprint('auth', __name__, url_prefix='/auth')


@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'GET':
        return render_template('auth/register.html')

    data = request.get_json() or request.form.to_dict()
    username = data.get('username', '').strip()
    password = data.get('password', '')
    role = data.get('role', 'student')
    realname = data.get('realname', '')
    email = data.get('email', '')

    if not username or not password:
        return format_response(400, '用户名和密码不能为空')

    if role not in ['admin', 'teacher', 'student']:
        return format_response(400, '无效的角色')

    user, msg = AuthService.register_user(username, password, role, realname, email)
    if not user:
        return format_response(400, msg)

    return format_response(200, msg, user.to_dict())


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('auth/login.html')

    data = request.get_json() or request.form.to_dict()
    username = data.get('username', '').strip()
    password = data.get('password', '')

    if not username or not password:
        return format_response(400, '用户名和密码不能为空')

    user, msg = AuthService.login_user(username, password)
    if not user:
        return format_response(401, msg)

    session['user_id'] = user.id
    session['username'] = user.username
    session['role'] = user.role

    return format_response(200, msg, {
        'id': user.id,
        'username': user.username,
        'role': user.role
    })


@auth_bp.route('/logout', methods=['POST'])
@login_required
def logout():
    session.clear()
    return format_response(200, '已退出登录')


@auth_bp.route('/profile', methods=['GET', 'PUT'])
@login_required
def profile():
    user_id = session['user_id']

    if request.method == 'GET':
        user = AuthService.get_user_by_id(user_id)
        return format_response(200, '成功', user.to_dict())

    data = request.get_json() or request.form.to_dict()
    realname = data.get('realname')
    email = data.get('email')

    success, msg = AuthService.update_user_profile(user_id, realname, email)
    return format_response(200 if success else 400, msg)


@auth_bp.route('/change-password', methods=['POST'])
@login_required
def change_password():
    data = request.get_json() or request.form.to_dict()
    old_password = data.get('old_password', '')
    new_password = data.get('new_password', '')

    if not new_password:
        return format_response(400, '新密码不能为空')

    success, msg = AuthService.change_password(session['user_id'], old_password, new_password)
    return format_response(200 if success else 400, msg)


@auth_bp.route('/users', methods=['GET'])
@login_required
def list_users():
    role = request.args.get('role')
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 10))

    from models import User
    query = User.query
    if role:
        query = query.filter_by(role=role)

    pagination = query.order_by(User.created_at.desc()).paginate(
        page=page, per_page=page_size, error_out=False
    )

    return format_response(200, '成功', {
        'items': [u.to_dict() for u in pagination.items],
        'total': pagination.total,
        'page': page,
        'page_size': page_size
    })


@auth_bp.route('/user/<int:user_id>', methods=['DELETE'])
@admin_required
def delete_user(user_id):
    success, msg = AuthService.delete_user(user_id)
    return format_response(200 if success else 400, msg)


@auth_bp.route('/admin/reset-password/<int:user_id>', methods=['POST'])
@admin_required
def admin_reset_password(user_id):
    data = request.get_json() or {}
    new_password = data.get('new_password', '')
    if not new_password:
        return format_response(400, '新密码不能为空')
    success, msg = AuthService.admin_reset_password(user_id, new_password)
    return format_response(200 if success else 400, msg)


@auth_bp.route('/current', methods=['GET'])
@login_required
def current_user():
    user = AuthService.get_user_by_id(session['user_id'])
    return format_response(200, '成功', user.to_dict())
