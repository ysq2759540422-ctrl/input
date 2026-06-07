from flask import Blueprint, request, session
from services.question_service import QuestionService
from utils.helpers import login_required, teacher_required, format_response
from models import db

question_bp = Blueprint('question', __name__, url_prefix='/api/questions')


@question_bp.route('', methods=['POST'])
@teacher_required
def create_question():
    data = request.get_json()
    if not data:
        return format_response(400, '请求体不能为空')

    title = data.get('title', '').strip()
    content = data.get('content', '').strip()
    answer = data.get('answer', '').strip()
    type = data.get('type', 'single')
    difficulty = data.get('difficulty', 'medium')
    category = data.get('category', 'general')
    score = data.get('score', 5)
    option_a = data.get('option_a')
    option_b = data.get('option_b')
    option_c = data.get('option_c')
    option_d = data.get('option_d')

    if not title or not content or not answer:
        return format_response(400, '标题、内容和答案不能为空')

    question, msg = QuestionService.create_question(
        title, content, answer, type, difficulty, category, score,
        option_a, option_b, option_c, option_d, session['user_id']
    )
    if not question:
        return format_response(400, msg)

    return format_response(200, msg, question.to_dict())


@question_bp.route('', methods=['GET'])
@login_required
def list_questions():
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 10))
    difficulty = request.args.get('difficulty')
    category = request.args.get('category')
    type = request.args.get('type')

    questions, total = QuestionService.list_questions(
        page, page_size, difficulty, category, type
    )

    return format_response(200, '成功', {
        'items': [q.to_dict() for q in questions],
        'total': total,
        'page': page,
        'page_size': page_size
    })


@question_bp.route('/<int:question_id>', methods=['GET'])
@login_required
def get_question(question_id):
    question, msg = QuestionService.get_question_by_id(question_id)
    if not question:
        return format_response(404, msg)
    return format_response(200, msg, question.to_dict())


@question_bp.route('/<int:question_id>', methods=['PUT'])
@teacher_required
def update_question(question_id):
    data = request.get_json()
    if not data:
        return format_response(400, '请求体不能为空')

    question, msg = QuestionService.update_question(question_id, **data)
    if not question:
        return format_response(404, msg)

    return format_response(200, msg, question.to_dict())


@question_bp.route('/<int:question_id>', methods=['DELETE'])
@teacher_required
def delete_question(question_id):
    success, msg = QuestionService.delete_question(question_id)
    return format_response(200 if success else 400, msg)


@question_bp.route('/batch', methods=['POST'])
@teacher_required
def batch_create_questions():
    data = request.get_json()
    if not data or 'questions' not in data:
        return format_response(400, '缺少questions数组')

    questions_data = data['questions']
    if not isinstance(questions_data, list) or len(questions_data) == 0:
        return format_response(400, 'questions必须是有效的非空数组')

    created, count = QuestionService.batch_create_questions(
        questions_data, session['user_id']
    )

    return format_response(200, f'成功创建{count}道题目', {
        'created_count': count,
        'ids': [q.id for q in created]
    })


@question_bp.route('/statistics', methods=['GET'])
@login_required
def question_statistics():
    difficulty = request.args.get('difficulty')
    category = request.args.get('category')

    from models import Question
    query = Question.query
    if difficulty:
        query = query.filter_by(difficulty=difficulty)
    if category:
        query = query.filter_by(category=category)

    total = query.count()

    by_difficulty = {}
    for diff in ['easy', 'medium', 'hard']:
        by_difficulty[diff] = Question.query.filter_by(difficulty=diff).count()

    by_category = {}
    categories = db.session.query(Question.category).distinct().all()
    for (cat,) in categories:
        by_category[cat] = Question.query.filter_by(category=cat).count()

    return format_response(200, '成功', {
        'total': total,
        'by_difficulty': by_difficulty,
        'by_category': by_category
    })
