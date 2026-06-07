[README.md](https://github.com/user-attachments/files/28682476/README.md)
# Online Examination System - 在线考试系统

## 项目概述

一个完整的在线考试系统 Web 应用，包含用户管理、题库管理、智能组卷、在线考试、自动评分、成绩查询等核心功能。
本项目专门为「开源软件代码评审与软件测试综合实验」课程设计，包含丰富的可评审代码缺陷和完整测试覆盖。

---

## 技术栈

| 层次 | 技术 |
|------|------|
| 后端框架 | Python 3.x / Flask 2.3 |
| ORM | Flask-SQLAlchemy 3.0 |
| 数据库 | SQLite（自动初始化） |
| 前端 | HTML + Vanilla JavaScript（无框架） |
| 测试 | pytest |
| 报告 | python-docx（Word）、openpyxl（Excel） |

---

## 项目架构

```
rjcs/
├── app.py                      # 应用入口（启动脚本）
├── config/config.py             # 系统配置
├── models/                     # 数据模型层
│   ├── user.py                # 用户模型
│   ├── question.py            # 题目模型
│   ├── exam.py                # 考试模型
│   └── score.py               # 成绩模型
├── routes/                     # 路由层（API接口）
│   ├── auth.py                # 认证路由
│   ├── question.py            # 题目路由
│   ├── exam.py                # 考试路由
│   └── score.py               # 成绩路由
├── services/                   # 业务逻辑层g）
│   ├── auth_service.py         # 认证服务
│   ├── question_service.py     # 题目服务
│   ├── exam_service.py         # 考试服务
│   └── score_service.py        # 成绩服务
├── utils/helpers.py            # 工具函数（权限装饰器）
├── templates/                  # HTML前端模板
└── requirements.txt            # 依赖清单
```

---

## 快速启动

### 1. 安装依赖

```bash
conda activate rjcs
pip install -r requirements.txt
```

### 2. 启动系统

```bash
python app.py
```

系统将在 `http://127.0.0.1:5000` 启动，数据库将自动初始化并插入测试数据。

### 3. 访问系统

- **URL**: http://127.0.0.1:5000
- **默认账号**:
  - 管理员: `admin` / `admin123`（角色: admin）
  - 教师: `teacher1` / `teacher123`（角色: teacher）
  - 学生: `student1` / `student123`（角色: student）
  - 学生: `student2` / `student123`
  - 学生: `student3` / `student123`

### 4. 运行测试

```bash
pytest tests/test_api.py -v
```

---

## 核心业务流程

1. **认证**: 用户注册/登录 → Session会话建立
2. **出题**: 教师登录 → 题库管理 → 创建/编辑/删除题目
3. **组卷**: 教师创建考试 → 设置参数 → 智能随机组卷
4. **考试**: 学生查看已发布考试 → 进入答题 → 提交试卷
5. **评分**: 教师手动/批量评分 → 系统自动比对答案 → 计算得分
6. **查分**: 学生查询个人成绩 → 查看排名统计

---

## 测试覆盖（35条测试用例）

| 模块 | 测试数量 | 说明 |
|------|---------|------|
| 认证模块 | 8 | 登录/注册/注销/权限 |
| 题目模块 | 11 | CRUD/筛选/分页/越权 |
| 考试模块 | 8 | 创建/发布/列表/边界 |
| 成绩模块 | 8 | 提交/评分/排名/重交 |

测试类型覆盖：正常场景、边界场景、异常场景、非法输入场景、安全测试

