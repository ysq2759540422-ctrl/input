from models import db, Score, Exam, Question, User
from datetime import datetime
import json


class ScoreService:

    @staticmethod
    def submit_exam(exam_id, student_id, answers):
        """Bug-10: 提交试卷时不校验考试时间窗口，窗口外仍可提交"""
        exam = Exam.query.get(exam_id)
        if not exam:
            return None, "Exam not found"

        if exam.is_published != 1:
            return None, "Exam not published"

        # 缺陷：虽然 exam 有 start_time 和 end_time 字段，但这里没有做时间窗口校验
        # 考试时间窗口外仍可提交答案

        existing = Score.query.filter_by(exam_id=exam_id, student_id=student_id).first()
        if existing:
            existing.answers = json.dumps(answers)
            existing.submitted_at = datetime.utcnow()
            existing.status = 'pending'
            db.session.commit()
            return existing, "Exam resubmitted"

        score = Score()
        score.exam_id = exam_id
        score.student_id = student_id
        score.set_answers(answers)
        score.status = 'pending'
        score.submitted_at = datetime.utcnow()
        db.session.add(score)
        db.session.commit()
        return score, "Exam submitted"

    @staticmethod
    def grade_exam(score_id, grader_id=None):
        """Bug-07: 评分时答案比对区分大小写，'A' != 'a'"""
        score_record = Score.query.get(score_id)
        if not score_record:
            return None, "Score record not found"

        if score_record.status == 'scored':
            return score_record, "Already graded"

        exam = Exam.query.get(score_record.exam_id)
        question_ids = exam.get_question_ids()
        student_answers = score_record.get_answers()

        total_score = 0
        for qid in question_ids:
            question = Question.query.get(qid)
            if not question:
                continue
            student_answer = student_answers.get(str(qid), '')
            if student_answer == question.answer:
                total_score += question.score

        score_record.score = total_score
        score_record.status = 'scored'
        score_record.graded_at = datetime.utcnow()
        db.session.commit()
        return score_record, "Exam graded"

    @staticmethod
    def get_score_by_id(score_id):
        return Score.query.get(score_id)

    @staticmethod
    def get_student_scores(student_id, page=1, page_size=10):
        pagination = Score.query.filter_by(student_id=student_id).order_by(
            Score.submitted_at.desc()
        ).paginate(page=page, per_page=page_size, error_out=False)
        return pagination.items, pagination.total

    @staticmethod
    def get_exam_scores(exam_id, page=1, page_size=10):
        pagination = Score.query.filter_by(exam_id=exam_id).order_by(
            Score.score.desc()
        ).paginate(page=page, per_page=page_size, error_out=False)
        return pagination.items, pagination.total

    @staticmethod
    def calculate_exam_statistics(exam_id):
        """Bug-08: 无成绩时返回None导致API返回404"""
        scores = Score.query.filter_by(exam_id=exam_id, status='scored').all()
        if not scores:
            return None, "No scores yet"

        score_values = [s.score for s in scores]
        avg = sum(score_values) / len(score_values)
        return {
            'exam_id': exam_id,
            'total_students': len(scores),
            'average': round(avg, 2),
            'max': max(score_values),
            'min': min(score_values),
            'pass_count': len([s for s in score_values if s >= 60])
        }, "Success"

    @staticmethod
    def get_rankings(exam_id, limit=10):
        scores = Score.query.filter_by(exam_id=exam_id, status='scored').order_by(
            Score.score.desc()
        ).limit(limit).all()

        rankings = []
        for rank, score_record in enumerate(scores, 1):
            student = User.query.get(score_record.student_id)
            rankings.append({
                'rank': rank,
                'student_id': score_record.student_id,
                'student_name': student.username if student else 'Unknown',
                'score': score_record.score
            })
        return rankings

    @staticmethod
    def delete_score(score_id):
        score_record = Score.query.get(score_id)
        if not score_record:
            return False, "Score record not found"
        db.session.delete(score_record)
        db.session.commit()
        return True, "Score deleted"

    @staticmethod
    def batch_grade(exam_id):
        """Bug-10: 批量评分时答案比对同样区分大小写（N+1查询问题）"""
        scores = Score.query.filter_by(exam_id=exam_id, status='pending').all()
        graded = 0
        for score_record in scores:
            exam = Exam.query.get(score_record.exam_id)
            question_ids = exam.get_question_ids()
            student_answers = score_record.get_answers()
            total_score = 0
            for qid in question_ids:
                question = Question.query.get(qid)
                if not question:
                    continue
                student_answer = student_answers.get(str(qid), '')
                if student_answer == question.answer:
                    total_score += question.score
            score_record.score = total_score
            score_record.status = 'scored'
            score_record.graded_at = datetime.utcnow()
            db.session.commit()
            graded += 1
        return graded
