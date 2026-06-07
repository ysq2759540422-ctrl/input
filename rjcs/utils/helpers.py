from functools import wraps
from flask import session, jsonify, redirect, url_for


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'code': 401, 'message': 'Not logged in'}), 401
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'code': 401, 'message': 'Not logged in'}), 401
        if session.get('role') != 'admin':
            return jsonify({'code': 403, 'message': 'Admin access required'}), 403
        return f(*args, **kwargs)
    return decorated_function


def teacher_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'code': 401, 'message': 'Not logged in'}), 401
        if session.get('role') not in ['admin', 'teacher']:
            return jsonify({'code': 403, 'message': 'Teacher access required'}), 403
        return f(*args, **kwargs)
    return decorated_function


def format_response(code, message, data=None):
    result = {'code': code, 'message': message}
    if data is not None:
        result['data'] = data
    return jsonify(result)


def paginate_query(query, page, page_size, max_page_size=100):
    if page < 1:
        page = 1
    if page_size < 1:
        page_size = 10
    if page_size > max_page_size:
        page_size = max_page_size

    pagination = query.paginate(page=page, per_page=page_size, error_out=False)
    return {
        'items': pagination.items,
        'total': pagination.total,
        'page': page,
        'page_size': page_size,
        'pages': pagination.pages
    }
