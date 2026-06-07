from flask import Blueprint, request, session
from services.score_service import ScoreService
from utils.helpers import login_required, teacher_required, format_response
from models import User, Exam

score_bp = Blueprint('score', __name__, url_prefix='/api/scores')


@score_bp.route('/submit', methods=['POST'])
@login_required
def submit_exam():
    data = request.get_json()
    if not data:
        return format_response(400, '请求体不能为空')

    exam_id = data.get('exam_id')
    answers = data.get('answers', {})

    if not exam_id:
        return format_response(400, 'exam_id不能为空')
    if not isinstance(answers, dict):
        return format_response(400, 'answers必须是一个对象')

    score_record, msg = ScoreService.submit_exam(exam_id, session['user_id'], answers)
    if not score_record:
        return format_response(400, msg)

    return format_response(200, msg, score_record.to_dict())


@score_bp.route('/<int:score_id>/grade', methods=['POST'])
@teacher_required
def grade_score(score_id):
    grader_id = session.get('user_id')
    score_record, msg = ScoreService.grade_exam(score_id, grader_id)
    if not score_record:
        return format_response(404, msg)

    return format_response(200, msg, score_record.to_dict())


@score_bp.route('/exam/<int:exam_id>/statistics', methods=['GET'])
@login_required
def exam_statistics(exam_id):
    result, msg = ScoreService.calculate_exam_statistics(exam_id)
    if not result:
        return format_response(404, msg)

    return format_response(200, msg, result)


@score_bp.route('/exam/<int:exam_id>/rankings', methods=['GET'])
@login_required
def exam_rankings(exam_id):
    limit = int(request.args.get('limit', 10))
    rankings = ScoreService.get_rankings(exam_id, limit)
    return format_response(200, '成功', rankings)


@score_bp.route('/exam/<int:exam_id>', methods=['GET'])
@login_required
def list_exam_scores(exam_id):
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 10))

    scores, total = ScoreService.get_exam_scores(exam_id, page, page_size)
    result = []
    for s in scores:
        d = s.to_dict()
        student = User.query.get(s.student_id)
        d['student_name'] = student.username if student else 'Unknown'
        result.append(d)

    return format_response(200, '成功', {
        'items': result,
        'total': total,
        'page': page,
        'page_size': page_size
    })


@score_bp.route('/my', methods=['GET'])
@login_required
def my_scores():
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 10))

    scores, total = ScoreService.get_student_scores(session['user_id'], page, page_size)
    result = []
    for s in scores:
        d = s.to_dict()
        exam = Exam.query.get(s.exam_id)
        d['exam_title'] = exam.title if exam else 'Unknown'
        result.append(d)

    return format_response(200, '成功', {
        'items': result,
        'total': total,
        'page': page,
        'page_size': page_size
    })


@score_bp.route('/exam/<int:exam_id>/batch-grade', methods=['POST'])
@teacher_required
def batch_grade(exam_id):
    graded_count = ScoreService.batch_grade(exam_id)
    return format_response(200, f'成功评分{graded_count}份试卷')


@score_bp.route('/<int:score_id>', methods=['DELETE'])
@teacher_required
def delete_score(score_id):
    success, msg = ScoreService.delete_score(score_id)
    return format_response(200 if success else 400, msg)
