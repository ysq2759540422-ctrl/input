from . import db
from datetime import datetime
import json


class Score(db.Model):
    __tablename__ = 'scores'

    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, db.ForeignKey('exams.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    score = db.Column(db.Integer, default=0)
    answers = db.Column(db.Text)
    status = db.Column(db.String(20), default='pending')
    submitted_at = db.Column(db.DateTime)
    graded_at = db.Column(db.DateTime)

    def get_answers(self):
        if not self.answers:
            return {}
        try:
            return json.loads(self.answers)
        except:
            return {}

    def set_answers(self, answers_dict):
        self.answers = json.dumps(answers_dict)

    def to_dict(self):
        return {
            'id': self.id,
            'exam_id': self.exam_id,
            'student_id': self.student_id,
            'score': self.score,
            'status': self.status,
            'submitted_at': self.submitted_at.strftime('%Y-%m-%d %H:%M:%S') if self.submitted_at else None,
            'graded_at': self.graded_at.strftime('%Y-%m-%d %H:%M:%S') if self.graded_at else None
        }
