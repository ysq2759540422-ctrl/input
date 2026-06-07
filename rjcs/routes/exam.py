from flask import Blueprint, request, session
from services.exam_service import ExamService
from utils.helpers import login_required, teacher_required, format_response
import json

exam_bp = Blueprint('exam', __name__, url_prefix='/api/exams')


@exam_bp.route('', methods=['POST'])
@teacher_required
def create_exam():
    data = request.get_json()
    if not data:
        return format_response(400, '请求体不能为空')

    title = data.get('title', '').strip()
    description = data.get('description', '').strip()
    duration = data.get('duration', 60)
    question_count = data.get('question_count', 10)
    auto_generate = data.get('auto_generate', True)
    difficulty_distribution = data.get('difficulty_distribution')
    category = data.get('category')

    if not title:
        return format_response(400, '考试标题不能为空')
    if not isinstance(difficulty_distribution, dict):
        difficulty_distribution = {'easy': 0.3, 'medium': 0.5, 'hard': 0.2}

    exam, msg = ExamService.create_exam(
        title, description, duration, question_count, session['user_id'],
        difficulty_distribution, category, auto_generate
    )
    if not exam:
        return format_response(400, msg)

    return format_response(200, msg, exam.to_dict())


@exam_bp.route('', methods=['GET'])
@login_required
def list_exams():
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 10))
    published_only = request.args.get('published_only', 'false').lower() == 'true'
    creator_id = request.args.get('creator_id', type=int)

    exams, total = ExamService.list_exams(page, page_size, published_only, creator_id)

    return format_response(200, '成功', {
        'items': [e.to_dict() for e in exams],
        'total': total,
        'page': page,
        'page_size': page_size
    })


@exam_bp.route('/<int:exam_id>', methods=['GET'])
@login_required
def get_exam(exam_id):
    exam, msg = ExamService.get_exam_by_id(exam_id)
    if not exam:
        return format_response(404, msg)
    return format_response(200, msg, exam.to_dict())


@exam_bp.route('/<int:exam_id>/questions', methods=['GET'])
@login_required
def get_exam_questions(exam_id):
    questions, msg = ExamService.get_exam_questions(exam_id)
    if questions is None:
        return format_response(404, msg)

    result = []
    for q in questions:
        d = q.to_dict()
        if session['role'] != 'admin' and session['role'] != 'teacher':
            d.pop('answer', None)
        result.append(d)

    return format_response(200, msg, result)


@exam_bp.route('/<int:exam_id>', methods=['PUT'])
@teacher_required
def update_exam(exam_id):
    data = request.get_json()
    if not data:
        return format_response(400, '请求体不能为空')

    exam, msg = ExamService.update_exam(exam_id, **data)
    if not exam:
        return format_response(404, msg)

    return format_response(200, msg, exam.to_dict())


@exam_bp.route('/<int:exam_id>', methods=['DELETE'])
@teacher_required
def delete_exam(exam_id):
    success, msg = ExamService.delete_exam(exam_id)
    return format_response(200 if success else 400, msg)


@exam_bp.route('/<int:exam_id>/publish', methods=['POST'])
@teacher_required
def publish_exam(exam_id):
    success, msg = ExamService.publish_exam(exam_id)
    return format_response(200 if success else 400, msg)


@exam_bp.route('/<int:exam_id>/clone', methods=['POST'])
@teacher_required
def clone_exam(exam_id):
    exam, msg = ExamService.clone_exam(exam_id, session['user_id'])
    if not exam:
        return format_response(404, msg)
    return format_response(200, msg, exam.to_dict())
