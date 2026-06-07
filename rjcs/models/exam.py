from . import db
from datetime import datetime
import json


class Exam(db.Model):
    __tablename__ = 'exams'

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    duration = db.Column(db.Integer, nullable=False, default=60)
    total_score = db.Column(db.Integer, default=100)
    question_count = db.Column(db.Integer, default=0)
    creator_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    question_ids = db.Column(db.Text)
    is_published = db.Column(db.Integer, default=0)
    start_time = db.Column(db.String(50))   # 考试开始时间（格式：YYYY-MM-DD HH:MM:SS）
    end_time = db.Column(db.String(50))     # 考试结束时间（格式：YYYY-MM-DD HH:MM:SS）
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    creator = db.relationship('User', backref='created_exams')
    scores = db.relationship('Score', backref='exam', lazy=True)

    def get_question_ids(self):
        if not self.question_ids:
            return []
        try:
            return json.loads(self.question_ids)
        except:
            return []

    def set_question_ids(self, ids_list):
        self.question_ids = json.dumps(ids_list)

    def to_dict(self, include_questions=False):
        result = {
            'id': self.id,
            'title': self.title,
            'description': self.description,
            'duration': self.duration,
            'total_score': self.total_score,
            'question_count': self.question_count,
            'creator_id': self.creator_id,
            'is_published': self.is_published,
            'start_time': self.start_time,
            'end_time': self.end_time,
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else None
        }
        return result
