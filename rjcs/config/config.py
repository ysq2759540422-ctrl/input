import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'exam_system_secret_key_2024'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or \
        'sqlite:///' + os.path.join(BASE_DIR, '..', 'exam_system.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    PAGE_SIZE = 10
    MAX_PAGE_SIZE = 100
    UPLOAD_FOLDER = os.path.join(BASE_DIR, '..', 'uploads')
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
    EXAM_DURATION_DEFAULT = 60
    MAX_QUESTIONS_PER_EXAM = 100
    MIN_QUESTIONS_PER_EXAM = 1
