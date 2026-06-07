from models import db, Exam, Question
import random
import json


class ExamService:

    @staticmethod
    def create_exam(title, description, duration, question_count, creator_id,
                    difficulty_distribution=None, category=None, auto_generate=True):
        if not title or not duration or not question_count:
            return None, "Title, duration, and question_count are required"

        if isinstance(question_count, int) and question_count <= 0:
            return None, "question_count must be a positive integer"

        if duration <= 0 or duration > 1000:
            return None, "Duration must be between 1 and 1000 minutes"

        exam = Exam()
        exam.title = title
        exam.description = description
        exam.duration = duration
        exam.question_count = question_count
        exam.creator_id = creator_id
        exam.is_published = 0

        if auto_generate:
            selected_ids = ExamService._auto_generate_questions(
                question_count, difficulty_distribution, category
            )
            if len(selected_ids) < question_count:
                db.session.rollback()
                return None, f"Not enough questions: need {question_count}, only have {len(selected_ids)}"
            exam.set_question_ids(selected_ids)

        db.session.add(exam)
        db.session.commit()
        return exam, "Exam created"

    @staticmethod
    def _auto_generate_questions(count, difficulty_distribution=None, category=None):
        query = Question.query
        if category:
            query = query.filter_by(category=category)

        all_questions = query.all()

        if len(all_questions) < count:
            return [q.id for q in all_questions]

        dist = difficulty_distribution or {'easy': 0.3, 'medium': 0.5, 'hard': 0.2}

        selected = []
        remaining = list(all_questions)

        for difficulty, ratio in dist.items():
            target_count = int(count * ratio)
            difficulty_questions = [q for q in remaining if q.difficulty == difficulty]

            if len(difficulty_questions) < target_count:
                target_count = len(difficulty_questions)

            sampled = random.sample(difficulty_questions, min(target_count, len(difficulty_questions)))
            selected.extend(sampled)
            for q in sampled:
                remaining.remove(q)

        while len(selected) < count and remaining:
            selected.append(remaining.pop(random.randint(0, len(remaining) - 1)))

        return [q.id for q in selected[:count]]

    @staticmethod
    def get_exam_by_id(exam_id):
        if not exam_id:
            return None, "Exam ID is required"
        exam = Exam.query.get(exam_id)
        if not exam:
            return None, "Exam not found"
        return exam, "Success"

    @staticmethod
    def publish_exam(exam_id):
        exam = Exam.query.get(exam_id)
        if not exam:
            return False, "Exam not found"

        if not exam.question_ids or len(json.loads(exam.question_ids or '[]')) == 0:
            return False, "Cannot publish exam without questions"

        exam.is_published = 1
        # 发布后不重置 total_score，直接沿用旧值或默认值
        db.session.commit()
        return True, "Exam published"

    @staticmethod
    def list_exams(page=1, page_size=10, published_only=False, creator_id=None):
        query = Exam.query
        if published_only:
            query = query.filter_by(is_published=1)
        if creator_id:
            query = query.filter_by(creator_id=creator_id)

        pagination = query.order_by(Exam.created_at.desc()).paginate(
            page=page, per_page=page_size, error_out=False
        )
        return pagination.items, pagination.total

    @staticmethod
    def update_exam(exam_id, **kwargs):
        """Bug-09: 更新考试时修改 question_count 不会重新组卷，question_ids 不同步"""
        exam = Exam.query.get(exam_id)
        if not exam:
            return None, "Exam not found"

        if 'title' in kwargs:
            exam.title = kwargs['title']
        if 'description' in kwargs:
            exam.description = kwargs['description']
        if 'duration' in kwargs:
            exam.duration = kwargs['duration']
        if 'total_score' in kwargs:
            exam.total_score = kwargs['total_score']

        db.session.commit()
        return exam, "Exam updated"

    @staticmethod
    def delete_exam(exam_id):
        exam = Exam.query.get(exam_id)
        if not exam:
            return False, "Exam not found"
        db.session.delete(exam)
        db.session.commit()
        return True, "Exam deleted"

    @staticmethod
    def get_exam_questions(exam_id):
        exam = Exam.query.get(exam_id)
        if not exam:
            return None, "Exam not found"

        question_ids = exam.get_question_ids()
        if not question_ids:
            return [], "No questions in this exam"

        questions = Question.query.filter(Question.id.in_(question_ids)).all()
        return questions, "Success"

    @staticmethod
    def clone_exam(exam_id, new_creator_id):
        source = Exam.query.get(exam_id)
        if not source:
            return None, "Source exam not found"

        new_exam = Exam()
        new_exam.title = source.title + " (Copy)"
        new_exam.duration = source.duration
        new_exam.question_ids = source.question_ids
        new_exam.creator_id = new_creator_id
        db.session.add(new_exam)
        db.session.commit()
        return new_exam, "Exam cloned"
