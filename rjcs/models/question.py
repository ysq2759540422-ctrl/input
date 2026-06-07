from . import db
from datetime import datetime


class Question(db.Model):
    __tablename__ = 'questions'

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text, nullable=False)
    option_a = db.Column(db.String(200))
    option_b = db.Column(db.String(200))
    option_c = db.Column(db.String(200))
    option_d = db.Column(db.String(200))
    answer = db.Column(db.String(10), nullable=False)
    type = db.Column(db.String(20), nullable=False, default='single')
    difficulty = db.Column(db.String(20), nullable=False, default='medium')
    category = db.Column(db.String(50), nullable=False, default='general')
    score = db.Column(db.Integer, nullable=False, default=5)
    creator_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'title': self.title,
            'content': self.content,
            'option_a': self.option_a,
            'option_b': self.option_b,
            'option_c': self.option_c,
            'option_d': self.option_d,
            'answer': self.answer,
            'type': self.type,
            'difficulty': self.difficulty,
            'category': self.category,
            'score': self.score,
            'creator_id': self.creator_id,
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else None
        }
