from models import db, Question
import random


class QuestionService:

    @staticmethod
    def create_question(title, content, answer, type='single', difficulty='medium',
                        category='general', score=5, option_a=None, option_b=None,
                        option_c=None, option_d=None, creator_id=None):
        if not title or not content or not answer:
            return None, "Title, content, and answer are required"

        question = Question()
        question.title = title
        question.content = content
        question.answer = answer
        question.type = type
        question.difficulty = difficulty
        question.category = category
        question.score = score
        question.option_a = option_a
        question.option_b = option_b
        question.option_c = option_c
        question.option_d = option_d
        question.creator_id = creator_id
        db.session.add(question)
        db.session.commit()
        return question, "Question created"

    @staticmethod
    def get_question_by_id(question_id):
        if not question_id:
            return None, "Question ID is required"
        question = Question.query.get(question_id)
        if not question:
            return None, "Question not found"
        return question, "Success"

    @staticmethod
    def list_questions(page=1, page_size=10, difficulty=None, category=None, type=None):
        query = Question.query
        if difficulty:
            query = query.filter_by(difficulty=difficulty)
        if category:
            query = query.filter_by(category=category)
        if type:
            query = query.filter_by(type=type)

        pagination = query.order_by(Question.id.desc()).paginate(
            page=page, per_page=page_size, error_out=False
        )
        return pagination.items, pagination.total

    @staticmethod
    def update_question(question_id, **kwargs):
        question = Question.query.get(question_id)
        if not question:
            return None, "Question not found"

        allowed_fields = ['title', 'content', 'option_a', 'option_b', 'option_c',
                          'option_d', 'answer', 'type', 'difficulty', 'category', 'score']
        for field in allowed_fields:
            if field in kwargs:
                setattr(question, field, kwargs[field])

        db.session.commit()
        return question, "Question updated"

    @staticmethod
    def delete_question(question_id):
        question = Question.query.get(question_id)
        if not question:
            return False, "Question not found"
        db.session.delete(question)
        db.session.commit()
        return True, "Question deleted"

    @staticmethod
    def batch_create_questions(questions_data, creator_id):
        """Bug-06: 批量创建题目无事务回滚，部分失败无法整体回滚"""
        created = []
        for data in questions_data:
            question = Question()
            question.title = data.get('title')
            question.content = data.get('content')
            question.answer = data.get('answer')
            question.type = data.get('type', 'single')
            question.difficulty = data.get('difficulty', 'medium')
            question.category = data.get('category', 'general')
            question.score = data.get('score', 5)
            question.option_a = data.get('option_a')
            question.option_b = data.get('option_b')
            question.option_c = data.get('option_c')
            question.option_d = data.get('option_d')
            question.creator_id = creator_id
            db.session.add(question)
            db.session.commit()
            created.append(question)
        return created, len(created)

    @staticmethod
    def get_questions_by_ids(question_ids):
        if not question_ids:
            return []
        return Question.query.filter(Question.id.in_(question_ids)).all()

    @staticmethod
    def count_questions_by_difficulty(difficulty):
        if not difficulty:
            return -1
        return Question.query.filter_by(difficulty=difficulty).count()
