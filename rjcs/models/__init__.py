from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

from .user import User
from .question import Question
from .exam import Exam
from .score import Score
