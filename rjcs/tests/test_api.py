"""
在线考试系统 - 自动化测试套件
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from app import create_app, init_data
from models import db, User, Question, Exam, Score


@pytest.fixture
def app():
    test_app = create_app()
    test_app.config['TESTING'] = True
    test_app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
    test_app.config['WTF_CSRF_ENABLED'] = False

    with test_app.app_context():
        db.create_all()
        init_data()
        yield test_app
        db.drop_all()


@pytest.fixture
def client(app):
    return app.test_client()


def get_auth_cookies(client, username='admin', password='admin123'):
    resp = client.post('/auth/login', json={'username': username, 'password': password})
    return resp


class TestAuthModule:
    """测试用例组1：认证模块"""

    def test_login_success(self, client):
        """TC-01: 正常场景 - 使用正确账号密码登录"""
        resp = client.post('/auth/login', json={'username': 'admin', 'password': 'admin123'})
        assert resp.status_code == 200
        data = resp.get_json()
        assert data['code'] == 200
        assert data['data']['username'] == 'admin'
        assert data['data']['role'] == 'admin'

    def test_login_wrong_password(self, client):
        """TC-02: 异常场景 - 使用错误密码登录"""
        resp = client.post('/auth/login', json={'username': 'admin', 'password': 'wrongpass'})
        data = resp.get_json()
        assert data['code'] == 401

    def test_login_user_not_found(self, client):
        """TC-03: 异常场景 - 使用不存在的用户名登录"""
        resp = client.post('/auth/login', json={'username': 'nonexistent', 'password': 'any'})
        data = resp.get_json()
        assert data['code'] == 401

    def test_login_empty_credentials(self, client):
        """TC-04: 异常场景 - 使用空用户名和密码登录"""
        resp = client.post('/auth/login', json={'username': '', 'password': ''})
        data = resp.get_json()
        assert data['code'] == 400

    def test_register_success(self, client):
        """TC-05: 正常场景 - 成功注册新用户"""
        resp = client.post('/auth/register', json={
            'username': 'newuser', 'password': 'newpass123', 'role': 'student'
        })
        data = resp.get_json()
        assert data['code'] == 200
        assert data['data']['username'] == 'newuser'

    def test_register_duplicate_username(self, client):
        """TC-06: 异常场景 - 注册重复用户名"""
        resp = client.post('/auth/login', json={'username': 'admin', 'password': 'admin123'})
        resp2 = client.post('/auth/register', json={
            'username': 'admin', 'password': 'pass', 'role': 'student'
        })
        data = resp2.get_json()
        assert data['code'] == 400

    def test_register_invalid_role(self, client):
        """TC-07: 异常场景 - 使用无效角色注册"""
        resp = client.post('/auth/register', json={
            'username': 'testuser', 'password': 'pass', 'role': 'superadmin'
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_register_empty_username(self, client):
        """TC-08: 异常场景 - 使用空用户名注册"""
        resp = client.post('/auth/register', json={
            'username': '', 'password': 'pass', 'role': 'student'
        })
        data = resp.get_json()
        assert data['code'] == 400


class TestQuestionModule:
    """测试用例组2：题目管理模块"""

    @pytest.fixture(autouse=True)
    def authenticated(self, client):
        client.post('/auth/login', json={'username': 'teacher1', 'password': 'teacher123'})

    def test_create_question_success(self, client):
        """TC-09: 正常场景 - 教师创建题目，填写全部字段"""
        resp = client.post('/api/questions',
            json={
                'title': '测试题目1',
                'content': '2+2等于多少？',
                'option_a': '3', 'option_b': '4', 'option_c': '5', 'option_d': '6',
                'answer': 'B',
                'type': 'single',
                'difficulty': 'easy',
                'category': '数学',
                'score': 5
            })
        data = resp.get_json()
        assert data['code'] == 200
        assert data['data']['title'] == '测试题目1'
        assert data['data']['answer'] == 'B'

    def test_create_question_missing_title(self, client):
        """TC-10: 异常场景 - 创建题目时标题为空"""
        resp = client.post('/api/questions', json={
            'content': '2+2=?', 'answer': 'B', 'score': 5
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_create_question_missing_answer(self, client):
        """TC-11: 异常场景 - 创建题目时答案为空"""
        resp = client.post('/api/questions', json={
            'title': '测试题', 'content': '内容', 'score': 5
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_create_question_invalid_difficulty(self, client):
        """TC-12: 边界场景 - 使用无效难度值创建题目"""
        resp = client.post('/api/questions', json={
            'title': '测试题', 'content': '内容', 'answer': 'A',
            'difficulty': 'extreme', 'score': 5
        })
        data = resp.get_json()
        assert data['code'] == 200

    def test_create_question_negative_score(self, client):
        """TC-13: 边界场景 - 使用负分值创建题目"""
        resp = client.post('/api/questions', json={
            'title': '测试题', 'content': '内容', 'answer': 'A',
            'score': -5
        })
        data = resp.get_json()
        assert data['code'] == 200

    def test_list_questions_pagination(self, client):
        """TC-14: 正常场景 - 分页查询题目列表"""
        resp = client.get('/api/questions?page=1&page_size=5')
        data = resp.get_json()
        assert data['code'] == 200
        assert 'items' in data['data']
        assert data['data']['page'] == 1
        assert data['data']['page_size'] == 5

    def test_list_questions_zero_page(self, client):
        """TC-15: 边界场景 - 使用page=0查询"""
        resp = client.get('/api/questions?page=0')
        data = resp.get_json()
        assert data['code'] == 200

    def test_list_questions_filter_difficulty(self, client):
        """TC-16: 正常场景 - 按难度筛选题目"""
        resp = client.get('/api/questions?difficulty=easy')
        data = resp.get_json()
        assert data['code'] == 200
        for q in data['data']['items']:
            assert q['difficulty'] == 'easy'

    def test_list_questions_filter_category(self, client):
        """TC-17: 正常场景 - 按分类筛选题目"""
        resp = client.get('/api/questions?category=Python')
        data = resp.get_json()
        assert data['code'] == 200

    def test_delete_question_not_found(self, client):
        """TC-18: 异常场景 - 删除不存在的题目"""
        resp = client.delete('/api/questions/99999')
        data = resp.get_json()
        assert data['code'] == 400

    def test_create_question_unauthorized(self, client):
        """TC-19: 安全测试 - 学生角色尝试创建题目"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student1', 'password': 'student123'})
        resp = client.post('/api/questions', json={
            'title': '测试题', 'content': '内容', 'answer': 'A', 'score': 5
        })
        assert resp.status_code == 403


class TestExamModule:
    """测试用例组3：考试管理模块"""

    @pytest.fixture(autouse=True)
    def teacher_auth(self, client):
        client.post('/auth/login', json={'username': 'teacher1', 'password': 'teacher123'})

    def test_create_exam_success(self, client):
        """TC-20: 正常场景 - 创建考试并自动生成题目"""
        resp = client.post('/api/exams', json={
            'title': '期中考试',
            'description': 'Python期中测试',
            'duration': 60,
            'question_count': 3,
            'difficulty_distribution': {'easy': 0.5, 'medium': 0.3, 'hard': 0.2}
        })
        data = resp.get_json()
        assert data['code'] == 200
        assert data['data']['title'] == '期中考试'
        assert data['data']['duration'] == 60

    def test_create_exam_missing_title(self, client):
        """TC-21: 异常场景 - 创建考试时不填写标题"""
        resp = client.post('/api/exams', json={
            'duration': 60, 'question_count': 5
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_create_exam_zero_duration(self, client):
        """TC-22: 边界场景 - 创建考试时考试时长为0"""
        resp = client.post('/api/exams', json={
            'title': '零时长考试', 'duration': 0, 'question_count': 5
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_create_exam_exceeds_db_count(self, client):
        """TC-23: 异常场景 - 请求题目数超过数据库中可用题目数"""
        resp = client.post('/api/exams', json={
            'title': '大题量考试', 'duration': 60, 'question_count': 999
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_create_exam_negative_question_count(self, client):
        """TC-24: 边界场景 - 创建考试时题目数量为负数"""
        resp = client.post('/api/exams', json={
            'title': '负数题目考试', 'duration': 60, 'question_count': -5
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_publish_exam_success(self, client):
        """TC-25: 正常场景 - 发布已有题目的考试"""
        resp = client.post('/api/exams', json={
            'title': '发布测试考试', 'duration': 30, 'question_count': 2
        })
        exam_id = resp.get_json()['data']['id']
        publish_resp = client.post(f'/api/exams/{exam_id}/publish')
        data = publish_resp.get_json()
        assert data['code'] == 200

    def test_list_published_exams(self, client):
        """TC-26: 正常场景 - 仅查询已发布的考试"""
        resp = client.get('/api/exams?published_only=true')
        data = resp.get_json()
        assert data['code'] == 200

    def test_get_exam_questions(self, client):
        """TC-27: 正常场景 - 获取特定考试的题目列表"""
        resp = client.post('/api/exams', json={
            'title': '题目获取测试', 'duration': 45, 'question_count': 2
        })
        exam_id = resp.get_json()['data']['id']
        quest_resp = client.get(f'/api/exams/{exam_id}/questions')
        data = quest_resp.get_json()
        assert data['code'] == 200


class TestScoreModule:
    """测试用例组4：成绩管理模块"""

    @pytest.fixture(autouse=True)
    def published_exam(self, app, client):
        client.post('/auth/login', json={'username': 'teacher1', 'password': 'teacher123'})
        resp = client.post('/api/exams', json={
            'title': '评分测试考试', 'duration': 30, 'question_count': 2
        })
        exam_id = resp.get_json()['data']['id']
        client.post(f'/api/exams/{exam_id}/publish')
        return exam_id

    def test_submit_exam_success(self, client, published_exam):
        """TC-28: 正常场景 - 学生提交考试，答案正确"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student1', 'password': 'student123'})
        resp = client.post('/api/scores/submit', json={
            'exam_id': published_exam,
            'answers': {'1': 'A', '2': 'B'}
        })
        data = resp.get_json()
        assert data['code'] == 200
        assert data['data']['status'] == 'pending'

    def test_submit_exam_missing_exam_id(self, client, published_exam):
        """TC-29: 异常场景 - 提交考试时不提供exam_id"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student2', 'password': 'student123'})
        resp = client.post('/api/scores/submit', json={
            'answers': {'1': 'A'}
        })
        data = resp.get_json()
        assert data['code'] == 400

    def test_grade_exam_case_sensitive_bug(self, client, published_exam):
        """TC-30: 异常场景 - 使用小写答案提交并评分（大小写敏感Bug）"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student3', 'password': 'student123'})
        resp = client.post('/api/scores/submit', json={
            'exam_id': published_exam,
            'answers': {'1': 'a', '2': 'b'}
        })
        score_id = resp.get_json()['data']['id']
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'teacher1', 'password': 'teacher123'})
        grade_resp = client.post(f'/api/scores/{score_id}/grade')
        grade_data = grade_resp.get_json()
        assert grade_data['code'] == 200

    def test_get_student_scores(self, client):
        """TC-31: 正常场景 - 学生查询个人成绩历史"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student1', 'password': 'student123'})
        resp = client.get('/api/scores/my')
        data = resp.get_json()
        assert data['code'] == 200
        assert 'items' in data['data']

    def test_exam_statistics_no_scores(self, client, published_exam):
        """TC-32: 边界场景 - 尚无成绩时的统计查询"""
        resp = client.get(f'/api/scores/exam/{published_exam}/statistics')
        data = resp.get_json()
        assert data['code'] in [200, 404]

    def test_rankings_limit(self, client, published_exam):
        """TC-33: 正常场景 - 获取考试排名前N名"""
        resp = client.get(f'/api/scores/exam/{published_exam}/rankings?limit=5')
        data = resp.get_json()
        assert data['code'] == 200
        assert isinstance(data['data'], list)

    def test_resubmit_exam(self, client, published_exam):
        """TC-34: 正常场景 - 学生重新提交同一考试（覆盖前次提交）"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student1', 'password': 'student123'})
        client.post('/api/scores/submit', json={
            'exam_id': published_exam, 'answers': {'1': 'A'}
        })
        resp = client.post('/api/scores/submit', json={
            'exam_id': published_exam, 'answers': {'1': 'B'}
        })
        data = resp.get_json()
        assert data['code'] == 200

    def test_submit_exam_not_published(self, client):
        """TC-35: 异常场景 - 提交未发布的考试"""
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'teacher1', 'password': 'teacher123'})
        resp = client.post('/api/exams', json={
            'title': '未发布考试', 'duration': 30, 'question_count': 2
        })
        exam_id = resp.get_json()['data']['id']
        client.post('/auth/logout', json={})
        client.post('/auth/login', json={'username': 'student1', 'password': 'student123'})
        submit_resp = client.post('/api/scores/submit', json={
            'exam_id': exam_id, 'answers': {'1': 'A'}
        })
        data = submit_resp.get_json()
        assert data['code'] == 400


if __name__ == '__main__':
    pytest.main([__file__, '-v', '--tb=short'])
