import os
from flask import Flask, render_template, session, redirect, url_for
from config.config import Config
from models import db, User, Question, Exam, Score
from routes.auth import auth_bp
from routes.question import question_bp
from routes.exam import exam_bp
from routes.score import score_bp


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)

    app.register_blueprint(auth_bp)
    app.register_blueprint(question_bp)
    app.register_blueprint(exam_bp)
    app.register_blueprint(score_bp)

    @app.route('/')
    def index():
        if 'user_id' in session:
            return redirect(url_for('dashboard'))
        return redirect(url_for('auth.login'))

    @app.route('/dashboard')
    def dashboard():
        if 'user_id' not in session:
            return redirect(url_for('auth.login'))
        role = session.get('role')
        role_cn = {'admin': '管理员', 'teacher': '教师', 'student': '学生'}.get(role, role)
        return render_template('dashboard.html', username=session.get('username'), role=role, role_cn=role_cn)

    @app.route('/questions')
    def questions_page():
        if 'user_id' not in session:
            return redirect(url_for('auth.login'))
        if session.get('role') not in ['admin', 'teacher']:
            return redirect(url_for('dashboard'))
        return render_template('questions/list.html')

    @app.route('/exams')
    def exams_page():
        if 'user_id' not in session:
            return redirect(url_for('auth.login'))
        return render_template('exams/list.html', role=session.get('role'))

    @app.route('/exams/<int:exam_id>')
    def exam_detail_page(exam_id):
        if 'user_id' not in session:
            return redirect(url_for('auth.login'))
        return render_template('exams/detail.html', exam_id=exam_id, role=session.get('role'))

    @app.route('/scores')
    def scores_page():
        if 'user_id' not in session:
            return redirect(url_for('auth.login'))
        return render_template('scores/my.html')

    @app.route('/scores/all')
    def all_scores_page():
        if 'user_id' not in session or session.get('role') not in ['admin', 'teacher']:
            return redirect(url_for('auth.login'))
        return render_template('scores/all.html', role=session.get('role'))

    @app.route('/admin/users')
    def admin_users_page():
        if 'user_id' not in session or session.get('role') not in ['admin', 'teacher']:
            return redirect(url_for('auth.login'))
        return render_template('admin/users.html', role=session.get('role'))

    @app.errorhandler(404)
    def not_found(e):
        return render_template('error.html', message='页面不存在'), 404

    @app.errorhandler(500)
    def server_error(e):
        return render_template('error.html', message='服务器内部错误'), 500

    return app


def init_data():
    from models import db, User, Question, Exam, Score
    from services.exam_service import ExamService

    if User.query.count() > 0:
        return

    admin = User(username='admin', password='admin123', role='admin',
                 realname='系统管理员', email='admin@example.com')
    teacher = User(username='teacher1', password='teacher123', role='teacher',
                   realname='李明', email='teacher@example.com')
    student1 = User(username='student1', password='student123', role='student',
                     realname='张三', email='zhangsan@example.com')
    student2 = User(username='student2', password='student123', role='student',
                     realname='李四', email='lisi@example.com')
    student3 = User(username='student3', password='student123', role='student',
                     realname='王五', email='wangwu@example.com')

    db.session.add_all([admin, teacher, student1, student2, student3])
    db.session.commit()

    questions_data = [
        {'title': 'Python变量命名规则', 'content': '以下哪个是合法的Python变量名？',
         'option_a': '2var', 'option_b': 'var_name', 'option_c': 'class', 'option_d': 'my-var',
         'answer': 'B', 'type': 'single', 'difficulty': 'easy', 'category': 'Python', 'score': 5},
        {'title': 'Python列表方法', 'content': '以下哪个方法可以将元素添加到列表末尾？',
         'option_a': 'add()', 'option_b': 'insert()', 'option_c': 'append()', 'option_d': 'push()',
         'answer': 'C', 'type': 'single', 'difficulty': 'easy', 'category': 'Python', 'score': 5},
        {'title': 'Flask路由装饰器', 'content': 'Flask中用于定义路由的装饰器是哪个？',
         'option_a': '@router', 'option_b': '@app.route', 'option_c': '@route', 'option_d': '@url',
         'answer': 'B', 'type': 'single', 'difficulty': 'easy', 'category': 'Flask', 'score': 5},
        {'title': 'HTTP状态码', 'content': 'HTTP状态码404表示什么？',
         'option_a': '200', 'option_b': '301', 'option_c': '404', 'option_d': '500',
         'answer': 'C', 'type': 'single', 'difficulty': 'easy', 'category': 'HTTP', 'score': 5},
        {'title': 'Python数据类型', 'content': '以下哪个不是Python的内置数据类型？',
         'option_a': 'list', 'option_b': 'dictionary', 'option_c': 'array', 'option_d': 'tuple',
         'answer': 'C', 'type': 'single', 'difficulty': 'medium', 'category': 'Python', 'score': 10},
        {'title': 'SQL SELECT查询', 'content': 'SQL中用于过滤分组数据的子句是哪个？',
         'option_a': 'WHERE', 'option_b': 'FILTER', 'option_c': 'HAVING', 'option_d': 'GROUP BY',
         'answer': 'C', 'type': 'single', 'difficulty': 'medium', 'category': 'SQL', 'score': 10},
        {'title': 'Git创建分支命令', 'content': '以下哪个Git命令用于创建新分支？',
         'option_a': 'git new branch', 'option_b': 'git checkout -b', 'option_c': 'git branch new', 'option_d': 'git create branch',
         'answer': 'B', 'type': 'single', 'difficulty': 'easy', 'category': 'Git', 'score': 5},
        {'title': 'Python循环控制', 'content': 'Python中哪个语句用于跳过当前迭代进入下一次循环？',
         'option_a': 'break', 'option_b': 'pass', 'option_c': 'continue', 'option_d': 'exit',
         'answer': 'C', 'type': 'single', 'difficulty': 'easy', 'category': 'Python', 'score': 5},
        {'title': 'RESTful API方法', 'content': 'HTTP协议中哪个方法通常用于更新资源？',
         'option_a': 'GET', 'option_b': 'POST', 'option_c': 'PUT', 'option_d': 'DELETE',
         'answer': 'C', 'type': 'single', 'difficulty': 'medium', 'category': 'HTTP', 'score': 10},
        {'title': '防止SQL注入', 'content': 'Python中以下哪种方式可以防止SQL注入攻击？',
         'option_a': '字符串拼接', 'option_b': 'f-strings', 'option_c': '参数化查询', 'option_d': 'raw()',
         'answer': 'C', 'type': 'single', 'difficulty': 'medium', 'category': 'SQL', 'score': 10},
        {'title': 'Flask会话存储', 'content': 'Flask中默认的Session数据存储方式是哪种？',
         'option_a': '数据库', 'option_b': 'Cookie', 'option_c': '文件系统', 'option_d': '仅内存',
         'answer': 'B', 'type': 'single', 'difficulty': 'medium', 'category': 'Flask', 'score': 10},
        {'title': 'Python异常处理', 'content': 'Python中用于捕获异常的关键词是哪个？',
         'option_a': 'catch', 'option_b': 'handle', 'option_c': 'except', 'option_d': 'try',
         'answer': 'C', 'type': 'single', 'difficulty': 'easy', 'category': 'Python', 'score': 5},
        {'title': '数据库规范化', 'content': '哪个范式要求消除部分函数依赖？',
         'option_a': '1NF', 'option_b': '2NF', 'option_c': '3NF', 'option_d': 'BCNF',
         'answer': 'B', 'type': 'single', 'difficulty': 'hard', 'category': 'SQL', 'score': 15},
        {'title': 'Python装饰器', 'content': 'Python中定义装饰器使用哪个符号？',
         'option_a': '#', 'option_b': '@', 'option_c': '*', 'option_d': '$',
         'answer': 'B', 'type': 'single', 'difficulty': 'easy', 'category': 'Python', 'score': 5},
        {'title': 'HTTP幂等方法', 'content': '以下哪个HTTP方法具有幂等性？',
         'option_a': 'POST', 'option_b': 'PATCH', 'option_c': 'DELETE', 'option_d': 'None',
         'answer': 'C', 'type': 'single', 'difficulty': 'hard', 'category': 'HTTP', 'score': 15},
    ]

    for qdata in questions_data:
        q = Question(**qdata, creator_id=teacher.id)
        db.session.add(q)

    db.session.commit()

    questions = Question.query.all()
    question_ids = [q.id for q in questions[:10]]

    exam1 = Exam(title='Python基础考试', description='测试Python基础知识掌握情况',
                 duration=60, question_count=10, creator_id=teacher.id,
                 total_score=100, is_published=1)
    exam1.set_question_ids(question_ids)
    db.session.add(exam1)

    exam2 = Exam(title='SQL与数据库考试', description='测试SQL及数据库知识',
                 duration=45, question_count=5, creator_id=teacher.id,
                 total_score=50, is_published=1)
    exam2.set_question_ids([q.id for q in questions if q.category == 'SQL'])
    db.session.add(exam2)

    db.session.commit()

    print("[Init] 数据库初始化完成，已插入测试数据。")


if __name__ == '__main__':
    app = create_app()

    with app.app_context():
        db.create_all()
        init_data()

    print("[Server] 在线考试系统已启动，访问地址：http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
