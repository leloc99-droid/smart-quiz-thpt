
from flask import Flask, render_template_string, request, redirect, url_for, session
from functools import wraps
from docx import Document
import re
import os
import json
import random
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(32)
TEACHER_PASSWORD = os.environ.get("TEACHER_PASSWORD", "ThayDoiMatKhau2026!")

def teacher_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("teacher_authenticated"):
            return redirect(url_for("teacher_login", next=request.path))
        return view(*args, **kwargs)
    return wrapped

SCHOOL = "HUỲNH TẤN PHÁT – VĨNH LONG"

QUESTIONS = {'Tin học': [],
 'Toán học': [{'id': 'toan1',
               'q': 'Nghiệm của phương trình x + 5 = 12 là gì?',
               'options': ['5', '6', '7', '8'],
               'answer': 2,
               'topic': 'Phương trình',
               'explain': 'Ta có x = 12 - 5 = 7.',
               'review': 'Ôn lại các phép biến đổi phương trình bậc nhất.'},
              {'id': 'toan2',
               'q': 'Giá trị của 3² + 4 là bao nhiêu?',
               'options': ['9', '12', '13', '16'],
               'answer': 2,
               'topic': 'Lũy thừa',
               'explain': '3² + 4 = 9 + 4 = 13.',
               'review': 'Ôn lại thứ tự thực hiện phép tính và lũy thừa.'}],
 'Tiếng Anh': [{'id': 'anh1',
                'q': 'Choose the correct answer: She _____ a student.',
                'options': ['am', 'is', 'are', 'be'],
                'answer': 1,
                'topic': 'Ngữ pháp',
                'explain': 'Với chủ ngữ She, động từ to be ở hiện tại là is.',
                'review': 'Ôn cách dùng am, is, are theo chủ ngữ.'},
               {'id': 'anh2',
                'q': "What is the opposite of 'big'?",
                'options': ['Tall', 'Long', 'Small', 'High'],
                'answer': 2,
                'topic': 'Từ vựng',
                'explain': 'Small có nghĩa là nhỏ, trái nghĩa với big.',
                'review': 'Ôn từ vựng chỉ kích thước và các cặp từ trái nghĩa.'}],
 'Lịch sử': [],
 'GDKT&PL': [],
 'Địa lí': [],
 'Vật lí': [{'id': 'li1',
             'q': 'Đơn vị đo lực trong hệ SI là gì?',
             'options': ['Jun', 'Oát', 'Niutơn', 'Vôn'],
             'answer': 2,
             'topic': 'Đại lượng và đơn vị',
             'explain': 'Đơn vị đo lực trong hệ SI là niutơn (N).',
             'review': 'Ôn lại các đại lượng vật lí và đơn vị đo tương ứng.'}],
 'Hóa học': [{'id': 'hoa1',
              'q': 'Kí hiệu hóa học của oxi là gì?',
              'options': ['Ox', 'O', 'Oi', 'H'],
              'answer': 1,
              'topic': 'Nguyên tố hóa học',
              'explain': 'Kí hiệu hóa học của nguyên tố oxi là O.',
              'review': 'Ôn lại kí hiệu hóa học của các nguyên tố thường gặp.'}]}

STYLE = """
<style>
* { box-sizing: border-box; }
body {
    font-family: Arial, sans-serif;
    background: #f0f4f8;
    margin: 0;
    padding: 22px 14px;
    color: #1e293b;
}
.container { max-width: 850px; margin: auto; }
.header {
    background: #075fea;
    color: white;
    padding: 28px 18px;
    border-radius: 16px;
    text-align: center;
}
.header h1 { margin: 0 0 10px; font-size: 32px; }
.header h2 { margin: 0 0 16px; font-size: 22px; }
.card {
    background: white;
    padding: 24px;
    margin-top: 20px;
    border-radius: 16px;
    box-shadow: 0 4px 15px #00000010;
}
.subjects {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 12px;
}
.subject, .btn {
    display: block;
    text-decoration: none;
    border: 0;
    border-radius: 10px;
    padding: 14px;
    text-align: center;
    font-size: 16px;
    cursor: pointer;
}
.subject {
    color: #174ea6;
    border: 1px solid #cbd5e1;
    font-weight: bold;
}
.subject:hover { background: #eff6ff; }
.btn { background: #155eef; color: white; }
.btn.secondary { background: #475467; }
.btn.green { background: #16803c; }
.option {
    display: block;
    padding: 13px;
    margin: 10px 0;
    border: 1px solid #cbd5e1;
    border-radius: 9px;
    cursor: pointer;
}
.option:hover { background: #eff6ff; }
.option input { margin-right: 8px; }
.result {
    background: #ecfdf3;
    padding: 16px;
    border-radius: 10px;
}
.analysis {
    background: #fff7e6;
    padding: 16px;
    border-radius: 10px;
    margin-top: 18px;
}
.footer {
    text-align: center;
    color: #64748b;
    margin: 20px 0;
    font-size: 13px;
}
.navbar { display:flex; flex-wrap:wrap; gap:0; background:#0878ed; border-radius:10px; margin:0 0 16px; overflow:hidden; }
.navbar a { color:#fff; text-decoration:none; padding:13px 14px; font-weight:700; font-size:14px; display:block; }
.navbar a:hover { background:#075fc2; }
.navbar .who { margin-left:auto; color:#fff; padding:13px 14px; font-weight:700; font-size:14px; }
@media (max-width: 480px) {
    .header h1 { font-size: 25px; }
    .header h2 { font-size: 19px; }
    .card { padding: 17px; }
}
</style>
"""

QUESTION_BANK_FILE = "question_bank.json"

if os.path.exists(QUESTION_BANK_FILE):
    try:
        with open(QUESTION_BANK_FILE, "r", encoding="utf-8") as f:
            saved_questions = json.load(f)

        for subject, items in saved_questions.items():
            if subject in QUESTIONS and isinstance(items, list):
                existing_ids = {q["id"] for q in QUESTIONS[subject]}
                for q in items:
                    if q.get("id") not in existing_ids:
                        QUESTIONS[subject].append(q)
                        existing_ids.add(q.get("id"))
    except (json.JSONDecodeError, OSError) as e:
        print("Không đọc được ngân hàng câu hỏi đã lưu:", e)




HEADER = """
<div class="header">
    <h1>SMART QUIZ THPT</h1>
    <h2>HUỲNH TẤN PHÁT – VĨNH LONG</h2>
    <div>Ứng dụng ôn tập trắc nghiệm thông minh</div>
    <p>Học tập mọi lúc – mọi nơi</p>
</div>
"""

NAVBAR = """
<nav class="navbar">
  <a href="{{ url_for('home') }}">⌂ Trang chủ</a>
  <a href="{{ url_for('home') }}">▤ Tạo đề</a>
  {% if teacher_authenticated or session.get('judge_name') %}
    <a href="{{ url_for('question_bank') }}">▦ Ngân hàng câu hỏi</a>
    <a href="{{ url_for('import_word') }}">▣ Nhập đề Word</a>
  {% endif %}
  {% if teacher_authenticated %}
    <span class="who">♙ {{ teacher_name or 'Giáo viên' }}</span>
  {% else %}
    {% if session.get('judge_name') %}<span class="who">♙ Giám khảo: {{ session.get('judge_name') }}</span>{% endif %}
    <a href="{{ url_for('teacher_login') }}">♙ Giáo viên</a>
  {% endif %}
</nav>
"""

HOME_TEMPLATE = """
<!DOCTYPE html><html lang="vi"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SMART QUIZ THPT - HUỲNH TẤN PHÁT</title>
""" + STYLE + """
</head><body><div class="container">
""" + NAVBAR + HEADER + """
<div class="card">
    <h2>Chào mừng thầy cô, các em học sinh và quý giám khảo!</h2>
    <p>Giám khảo có thể trải nghiệm phần mềm ngay, không cần mật khẩu.</p>
    <form method="post" action="{{ url_for('judge_entry') }}" style="display:flex;gap:10px;flex-wrap:wrap;margin:14px 0 20px">
        <input name="judge_name" type="text" maxlength="80" placeholder="Nhập họ tên (không bắt buộc)" value="{{ session.get('judge_name','') }}" style="flex:1;min-width:220px;padding:12px;border:1px solid #cbd5e1;border-radius:8px">
        <button class="btn" type="submit">Bắt đầu trải nghiệm</button>
    </form>
    <p>Hoặc chọn môn học bên dưới để bắt đầu làm bài.</p>
    <div class="subjects">
        {% for subject in subjects %}
        <a class="subject" href="{{ url_for('quiz', subject=subject) }}">
            {{ subject }}<br><small>Bắt đầu ôn tập →</small>
        </a>
        {% endfor %}
    </div>
</div>

<div class="card">
    {% if teacher_authenticated %}
    <h2>Khu vực giáo viên</h2>
    <p>Thầy cô đã đăng nhập. Có thể nhập đề Word vào ngân hàng câu hỏi.</p>
    <a class="btn secondary" href="{{ url_for('import_word') }}">Nhập đề Word</a>
    <p style="margin-top:12px"><a href="{{ url_for('teacher_logout') }}">Đăng xuất giáo viên</a></p>
    {% else %}
    <h2>Dành cho giáo viên</h2>
    <p>Giáo viên đăng nhập để thêm đề vào ngân hàng câu hỏi. Học sinh có thể ôn tập mà không cần tài khoản.</p>
    <a class="btn secondary" href="{{ url_for('teacher_login') }}">Đăng nhập giáo viên</a>
    {% endif %}
</div>

<div class="footer">Sản phẩm nghiên cứu khoa học kỹ thuật học sinh THPT</div>
</div></body></html>
"""
QUIZ_TEMPLATE = """
<!DOCTYPE html><html lang="vi"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ôn tập {{ subject }} - SMART QUIZ THPT</title>
""" + STYLE + """
</head><body><div class="container">
""" + NAVBAR + HEADER + """
<div class="card">
    <a class="btn secondary" href="{{ url_for('home') }}">← Chọn môn khác</a>
    <h2>Ôn tập môn {{ subject }}</h2>

    {% if exam_notice %}
    <div class="analysis"><p><b>{{ exam_notice }}</b></p></div>
    {% endif %}
    {% if retry_mode %}
    <p><b>Luyện tập lại các câu trả lời sai</b></p>
    {% elif questions %}
    <p>Số câu hỏi: {{ questions|length }}</p>
    {% endif %}

    {% if validation_error %}
    <div class="analysis" role="alert" style="border:2px solid #b42318;background:#fff1f0;color:#b42318">
        <b>Chưa thể nộp bài:</b> {{ validation_error }}
    </div>
    {% endif %}

    {% if result is none and questions %}
    <form method="post" id="quiz-form">
        <input type="hidden" name="retry_mode" value="{{ 1 if retry_mode else 0 }}">
        {% for question in questions %}
        {% set qi = loop.index0 %}
        <div style="margin: 22px 0">
            <h3>Câu {{ loop.index }}: {{ question.q }}</h3>
            {% if question.get("type") == "true_false" %}
                <p>Chọn Đúng hoặc Sai cho từng nhận định:</p>
                {% for statement in question.statements %}
                <div class="option">
                    <b>{{ ["a", "b", "c", "d"][loop.index0] }}.</b>
                    {{ statement.text }}
                    <label style="display:inline-block;margin:10px 12px 0 0">
                        <input type="radio" name="q_{{ question.id }}_{{ loop.index0 }}" value="true">
                        Đúng
                    </label>
                    <label style="display:inline-block;margin:10px 0 0">
                        <input type="radio" name="q_{{ question.id }}_{{ loop.index0 }}" value="false">
                        Sai
                    </label>
                </div>
                {% endfor %}
            {% else %}
                {% for option in question.options %}
                <label class="option">
                    <input type="radio" name="q_{{ question.id }}"
                           value="{{ loop.index0 }}">
                    {{ ["A", "B", "C", "D"][loop.index0] }}. {{ option }}
                </label>
                {% endfor %}
            {% endif %}
        </div>
        {% endfor %}
        <p style="color:#b42318;font-weight:bold">Lưu ý: Mỗi câu trắc nghiệm và từng nhận định Đúng/Sai đều phải được chọn. Hệ thống không tự chọn đáp án thay học sinh.</p>
        <button class="btn" type="submit">Nộp bài và xem kết quả</button>
    </form>
    <script>
    (function () {
        const form = document.getElementById('quiz-form');
        if (!form) return;
        form.addEventListener('submit', function (event) {
            const groups = new Map();
            form.querySelectorAll('input[type="radio"]').forEach(function (radio) {
                if (!groups.has(radio.name)) groups.set(radio.name, []);
                groups.get(radio.name).push(radio);
            });
            const unanswered = [];
            groups.forEach(function (radios, name) {
                if (!radios.some(function (radio) { return radio.checked; })) {
                    unanswered.push(name);
                    radios.forEach(function (radio) {
                        const label = radio.closest('label') || radio.closest('.option');
                        if (label) label.style.border = '2px solid #b42318';
                    });
                }
            });
            if (unanswered.length) {
                event.preventDefault();
                alert('Bạn chưa trả lời đầy đủ ' + unanswered.length + ' câu hỏi/nhận định. Vui lòng chọn đáp án cho tất cả trước khi nộp bài. Hệ thống không tự chọn đáp án.');
                const first = form.querySelector('input[name="' + unanswered[0] + '"]');
                if (first) first.scrollIntoView({behavior: 'smooth', block: 'center'});
            }
        });
    })();
    </script>

    {% elif result is none %}
    <div class="analysis"><p>Chưa thể tạo đề thi. Thầy cô cần nhập đủ câu hỏi hợp lệ vào ngân hàng rồi thử lại.</p></div>
    {% else %}
    <div class="result">
        {% if retry_mode %}
        <h2>Kết quả luyện lại câu sai</h2>
        <p>Điểm phần luyện lại: <b>{{ "%.2f"|format(result.score) }}/{{ "%.2f"|format(result.total) }} điểm</b></p>
        <p style="margin-bottom:0">Đây là điểm của riêng {{ questions|length }} câu được luyện lại, <b>không phải điểm của toàn bộ đề 28 câu</b>.</p>
        {% else %}
        <h2>Kết quả bài làm</h2>
        <p>Điểm đạt theo trọng số: <b>{{ "%.2f"|format(result.score) }}/{{ "%.2f"|format(result.total) }}</b></p>
        <p>Điểm quy đổi:
            <b>{{ "%.2f"|format(result.score * 10 / result.total) if result.total else "0.00" }}/10</b>
        </p>
        {% endif %}
    </div>

    <div class="analysis">
        <h3>Phân tích kiến thức cần ôn tập</h3>
        {% if result.wrong %}
            {% if retry_mode %}
            <p>Trong lượt luyện lại này, còn {{ result.wrong|length }} câu chưa chính xác. Với câu Đúng/Sai, hệ thống hiển thị riêng kết quả từng nhận định a, b, c, d bên dưới.</p>
            {% else %}
            <p>Ứng dụng ghi nhận {{ result.wrong|length }} câu chưa chính xác.</p>
            {% endif %}
            {% for topic, items in result.topics.items() %}
                <p><b>{{ topic }}</b></p>
                {% for item in items %}
                    <p>• {{ item.review }}</p>
                {% endfor %}
            {% endfor %}
            <p>Đây là gợi ý tự động dựa trên chủ đề của các câu trả lời sai.</p>
        {% else %}
            <p>Chúc mừng em! Tất cả câu trả lời đều chính xác.
            Hãy tiếp tục luyện tập để củng cố kiến thức.</p>
        {% endif %}
    </div>

    <h3>Đáp án và giải thích</h3>
    {% for item in result.details %}
    <div class="card">
        <b>{% if retry_mode %}Câu luyện lại {{ loop.index }}{% else %}{{ loop.index }}{% endif %}. {{ item.q }}</b>
        <p>Chủ đề: {{ item.topic }}</p>
        <p>Bài làm: {{ item.selected }}</p>
        <p>Đáp án đúng: <b>{{ item.correct_answer }}</b></p>
        <p style="color: {{ 'green' if item.correct else 'red' }}">
            {{ 'Chính xác' if item.correct else 'Chưa chính xác' }}
        </p>
        <p>Giải thích: {{ item.explain }}</p>
        {% if item.statements %}
            {% for st in item.statements %}
            <div class="option">
                <b>{{ st.label }}.</b> {{ st.text }}<br>
                Bài làm: {{ st.selected }} |
                Đáp án: <b>{{ st.answer }}</b> |
                <span style="color: {{ 'green' if st.correct else 'red' }}">
                    {{ 'Đúng' if st.correct else 'Chưa đúng' }}
                </span>
                <p>Giải thích: {{ st.explain }}</p>
            </div>
            {% endfor %}
        {% endif %}
    </div>
    {% endfor %}

    <p>
        {% if result.wrong %}
        <a class="btn green"
           href="{{ url_for('quiz', subject=subject, retry=1) }}">
           Luyện tập lại câu sai
        </a>
        {% endif %}
    </p>
    <p><a class="btn" href="{{ url_for('quiz', subject=subject) }}">
        Làm lại toàn bộ bài
    </a></p>
    {% endif %}
</div>
<div class="footer">SMART QUIZ THPT | {{ school }}</div>
</div></body></html>
"""


@app.route("/judge-entry", methods=["POST"])
def judge_entry():
    # Chế độ trải nghiệm dành cho giám khảo: chỉ lưu tên hiển thị, không cấp quyền quản trị.
    judge_name = re.sub(r"\s+", " ", request.form.get("judge_name", "").strip())[:80]
    if judge_name:
        session["judge_name"] = judge_name
    else:
        session.pop("judge_name", None)
    return redirect(url_for("home"))


@app.route("/")
def home():
    return render_template_string(
        HOME_TEMPLATE,
        subjects=list(QUESTIONS.keys()),
        teacher_authenticated=bool(session.get("teacher_authenticated")),
        teacher_name=session.get("teacher_name", "")
    )

@app.route("/quiz/<path:subject>", methods=["GET", "POST"])
def quiz(subject):
    if subject not in QUESTIONS:
        return redirect(url_for("home"))

    retry_mode = (
        request.args.get("retry") == "1"
        or request.form.get("retry_mode") == "1"
    )

    exam_notice = None
    if retry_mode:
        wrong_ids = session.get("wrong_by_subject", {}).get(subject, [])
        questions = [q for q in QUESTIONS[subject] if q["id"] in wrong_ids]
        if not questions:
            return redirect(url_for("quiz", subject=subject))
    else:
        # Mỗi lượt ôn tập của mọi môn gồm 24 câu trắc nghiệm và 4 câu Đúng/Sai.
        # Khi nộp bài, dùng đúng bộ câu hỏi đã được lưu trong session để chấm cùng đề đã làm.
        mc_bank = [
            q for q in QUESTIONS[subject]
            if q.get("type") != "true_false"
            and len(q.get("options", [])) == 4
            and isinstance(q.get("answer"), int)
        ]
        tf_bank = [
            q for q in QUESTIONS[subject]
            if q.get("type") == "true_false"
            and len(q.get("statements", [])) == 4
            and all(st.get("answer") is not None for st in q.get("statements", []))
        ]
        if request.method == "POST":
            exam_ids = session.get("active_exam_ids", {}).get(subject, [])
            questions = [q for q in QUESTIONS[subject] if q.get("id") in exam_ids]
        else:
            if len(mc_bank) >= 24 and len(tf_bank) >= 4:
                selected = random.sample(mc_bank, 24) + random.sample(tf_bank, 4)
                random.shuffle(selected)
                questions = selected
                session["active_exam_ids"] = {
                    **session.get("active_exam_ids", {}),
                    subject: [q["id"] for q in questions]
                }
            else:
                questions = []
                exam_notice = (
                    f"Ngân hàng môn {subject} hiện có {len(mc_bank)} câu trắc nghiệm "
                    f"và {len(tf_bank)} câu Đúng/Sai hợp lệ. "
                    "Cần tối thiểu 24 câu trắc nghiệm và 4 câu Đúng/Sai "
                    "(mỗi câu có đủ 4 nhận định và đáp án) để tạo đề."
                )

    result = None
    validation_error = None

    if request.method == "POST":
        # Kiểm tra ở máy chủ: tuyệt đối không tự hiểu nhận định bỏ trống là Đúng/Sai.
        missing = []
        for q in questions:
            if q.get("type") == "true_false":
                for idx, statement in enumerate(q.get("statements", [])):
                    raw = request.form.get(f'q_{q["id"]}_{idx}')
                    if raw not in ("true", "false"):
                        missing.append(f'Câu {q.get("q", "")}: nhận định {chr(ord("a") + idx)}')
            else:
                raw = request.form.get("q_" + q["id"])
                try:
                    selected_index = int(raw)
                    if selected_index < 0 or selected_index >= len(q.get("options", [])):
                        raise ValueError
                except (TypeError, ValueError):
                    missing.append(f'Câu {q.get("q", "")}')
        if missing:
            return render_template_string(
                QUIZ_TEMPLATE,
                subject=subject,
                questions=questions,
                result=None,
                retry_mode=retry_mode,
                exam_notice=exam_notice,
                validation_error=f'{len(missing)} câu/nhận định chưa được chọn. Hãy chọn đầy đủ trước khi nộp bài.',
                school=SCHOOL,
                teacher_authenticated=bool(session.get("teacher_authenticated")),
                teacher_name=session.get("teacher_name", "")
            )

        score = 0
        details = []
        wrong = []
        topics = {}

        total_points = 0

        for question in questions:
            if question.get("type") == "true_false":
                statement_details = []
                correct_count = 0
                for idx, statement in enumerate(question["statements"]):
                    raw = request.form.get(f'q_{question["id"]}_{idx}')
                    selected_bool = raw == "true"
                    answered = raw in ("true", "false")
                    # Chuẩn hóa đáp án cũ trong question_bank.json: có thể là bool,
                    # chuỗi "Sai"/"Đúng", hoặc chuỗi "False"/"True".
                    stored_answer = statement.get("answer")
                    if isinstance(stored_answer, bool):
                        answer_bool = stored_answer
                    else:
                        answer_bool = _normalize_tf_answer(str(stored_answer))
                    # Không dùng bool("False") hoặc bool("Sai"), vì chuỗi không rỗng
                    # trong Python đều thành True và gây chấm sai nhận định Sai.
                    # So sánh bằng nhãn hiển thị đã chuẩn hóa (Đúng/Sai), tránh
                    # lệch kiểu dữ liệu giữa dữ liệu cũ và dữ liệu vừa nhập từ Word.
                    selected_label = ("Đúng" if raw == "true" else "Sai") if answered else None
                    answer_label = ("Đúng" if answer_bool is True else "Sai") if answer_bool is not None else None
                    is_correct = bool(answered and answer_label is not None and selected_label == answer_label)
                    if is_correct:
                        correct_count += 1
                    statement_details.append({
                        "label": chr(ord("a") + idx),
                        "text": statement["text"],
                        "selected": selected_label if answered else "Chưa trả lời",
                        "answer": answer_label if answer_label is not None else "Chưa nhận diện",
                        "correct": is_correct,
                        "explain": statement.get("explain", "")
                    })
                # Quy tắc thi tốt nghiệp THPT: câu Đúng/Sai gồm 4 nhận định.
                # Số nhận định đúng 0/1/2/3/4 tương ứng 0/0.1/0.25/0.5/1 điểm.
                tf_scores = {0: 0.0, 1: 0.1, 2: 0.25, 3: 0.5, 4: 1.0}
                question_points = tf_scores.get(correct_count, 0.0)
                score += question_points
                total_points += 1.0
                fully_correct = correct_count == len(question["statements"])
                if not fully_correct:
                    wrong.append(question["id"])
                    topics.setdefault(question["topic"], []).append(question)
                details.append({
                    "q": question["q"],
                    "topic": question["topic"],
                    "selected": f"Đúng {correct_count}/{len(question['statements'])} ý (đạt {question_points:.2f} điểm)",
                    "correct_answer": " | ".join(
                        f'{d["label"]}. {d["answer"]}' for d in statement_details
                    ),
                    "correct": fully_correct,
                    "explain": question.get("explain", ""),
                    "statements": statement_details
                })
            else:
                selected = request.form.get("q_" + question["id"])
                try:
                    selected_index = int(selected)
                except (TypeError, ValueError):
                    selected_index = -1

                correct = selected_index == question["answer"]
                if correct:
                    score += 0.25
                else:
                    wrong.append(question["id"])
                    topics.setdefault(question["topic"], []).append(question)
                total_points += 0.25

                if 0 <= selected_index < len(question["options"]):
                    selected_text = question["options"][selected_index]
                else:
                    selected_text = "Chưa trả lời"

                details.append({
                    "q": question["q"],
                    "topic": question["topic"],
                    "selected": selected_text,
                    "correct_answer": question["options"][question["answer"]],
                    "correct": correct,
                    "explain": question["explain"],
                    "statements": []
                })

        session["wrong_by_subject"] = {
            **session.get("wrong_by_subject", {}),
            subject: wrong
        }

        result = {
            "score": round(score, 2),
            "total": round(total_points, 2),
            "details": details,
            "wrong": wrong,
            "topics": topics
        }

    return render_template_string(
        QUIZ_TEMPLATE,
        subject=subject,
        questions=questions,
        result=result,
        retry_mode=retry_mode,
        exam_notice=exam_notice,
        validation_error=validation_error,
        school=SCHOOL,
        teacher_authenticated=bool(session.get("teacher_authenticated")),
        teacher_name=session.get("teacher_name", "")
    )

TEACHER_LOGIN_TEMPLATE = """
<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Đăng nhập giáo viên - SMART QUIZ THPT</title>
""" + STYLE + """
</head><body><div class="container">""" + NAVBAR + HEADER + """
<div class="card"><h2>Đăng nhập giáo viên</h2>
<p>Học sinh không cần đăng nhập để làm bài. Chức năng này chỉ dành cho giáo viên quản lý đề.</p>
{% if error %}<p style="color:#b42318">{{ error }}</p>{% endif %}
<form method="post">
<label for="teacher_name">Tên giáo viên</label>
<input id="teacher_name" name="teacher_name" type="text" required maxlength="80" autocomplete="name" placeholder="Ví dụ: Nguyễn Văn An"
style="display:block;width:100%;padding:12px;margin:10px 0 16px;border:1px solid #cbd5e1;border-radius:8px">
<label for="password">Mật khẩu giáo viên</label>
<input id="password" name="password" type="password" required autocomplete="current-password"
style="display:block;width:100%;padding:12px;margin:10px 0 16px;border:1px solid #cbd5e1;border-radius:8px">
<input type="hidden" name="next" value="{{ next }}">
<button class="btn" type="submit">Đăng nhập</button>
</form><p><a href="{{ url_for('home') }}">← Quay lại trang chủ</a></p>
</div></div></body></html>
"""

@app.route("/teacher-login", methods=["GET", "POST"])
def teacher_login():
    error = None
    next_url = request.args.get("next", request.form.get("next", ""))
    if request.method == "POST":
        if not TEACHER_PASSWORD:
            error = "Chưa cấu hình mật khẩu giáo viên trên máy chủ. Vui lòng liên hệ người quản trị."
        elif request.form.get("password", "") == TEACHER_PASSWORD:
            teacher_name = re.sub(r"\s+", " ", request.form.get("teacher_name", "").strip())[:80]
            if not teacher_name:
                error = "Thầy cô vui lòng nhập tên hiển thị."
                return render_template_string(TEACHER_LOGIN_TEMPLATE, error=error, next=next_url,
                    teacher_authenticated=False, teacher_name="")
            session.clear()
            session["teacher_authenticated"] = True
            session["teacher_name"] = teacher_name
            # Chỉ cho chuyển hướng nội bộ, tránh open redirect.
            if next_url.startswith("/") and not next_url.startswith("//"):
                return redirect(next_url)
            return redirect(url_for("home"))
        else:
            error = "Mật khẩu chưa đúng. Thầy cô vui lòng thử lại."
    return render_template_string(TEACHER_LOGIN_TEMPLATE, error=error, next=next_url,
        teacher_authenticated=bool(session.get("teacher_authenticated")),
        teacher_name=session.get("teacher_name", ""))

@app.route("/teacher-logout")
def teacher_logout():
    session.pop("teacher_authenticated", None)
    session.pop("teacher_name", None)
    return redirect(url_for("home"))

QUESTION_BANK_TEMPLATE = """
<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ngân hàng câu hỏi - SMART QUIZ THPT</title>
""" + STYLE + """
</head><body><div class="container">
""" + NAVBAR + HEADER + """
<div class="card">
  <h2>▦ Ngân hàng câu hỏi</h2>
  <p>Xin chào {{ teacher_name }}. Thầy cô có thể chọn môn và tìm câu hỏi trong ngân hàng.</p>
  <form method="get" style="display:flex;gap:10px;flex-wrap:wrap;margin:16px 0">
    <select name="subject" style="padding:12px;border:1px solid #cbd5e1;border-radius:8px;min-width:180px">
      {% for s in subjects %}<option value="{{ s }}" {% if s == selected_subject %}selected{% endif %}>{{ s }}</option>{% endfor %}
    </select>
    <input name="q" value="{{ search_text }}" placeholder="Tìm nội dung câu hỏi..." style="flex:1;min-width:200px;padding:12px;border:1px solid #cbd5e1;border-radius:8px">
    <button class="btn" type="submit">Tìm kiếm</button>
  </form>
  <p><b>{{ questions|length }}</b> câu hiển thị / {{ total_subject }} câu trong môn {{ selected_subject }}.</p>
  {% for question in questions %}
    <div class="card" style="border:1px solid #dbe3ee;box-shadow:none">
      <h3>Câu {{ loop.index }}. {{ question.q }}</h3>
      {% if question.get('type') == 'true_false' %}
        {% for st in question.statements %}
          <p>{{ ['a','b','c','d'][loop.index0] }}. {{ st.text }}
          <b style="color:#087443">— {{ 'Đúng' if st.answer is sameas true else ('Sai' if st.answer is sameas false else st.answer) }}</b></p>
        {% endfor %}
      {% else %}
        {% for option in question.options %}
          <p {% if loop.index0 == question.answer %}style="font-weight:bold;color:#087443"{% endif %}>
            {{ ['A','B','C','D'][loop.index0] }}. {{ option }}
            {% if loop.index0 == question.answer %} ✓ Đáp án{% endif %}
          </p>
        {% endfor %}
      {% endif %}
      <p style="color:#64748b;font-size:13px">Chủ đề: {{ question.get('topic','Chưa phân loại') }} · Mã: {{ question.get('id','') }}</p>
    </div>
  {% else %}
    <p>Không tìm thấy câu hỏi phù hợp.</p>
  {% endfor %}
  <p><a class="btn secondary" href="{{ url_for('home') }}">← Về trang chủ</a></p>
</div></div></body></html>
"""

@app.route("/question-bank")
def question_bank():
    # Giám khảo đã nhập tên được xem ngân hàng; chỉ giáo viên mới có quyền quản trị.
    if not session.get("teacher_authenticated") and not session.get("judge_name"):
        return redirect(url_for("home"))
    subjects = list(QUESTIONS.keys())
    selected_subject = request.args.get("subject", subjects[0] if subjects else "")
    if selected_subject not in QUESTIONS:
        selected_subject = subjects[0] if subjects else ""
    search_text = request.args.get("q", "").strip()
    all_questions = QUESTIONS.get(selected_subject, [])
    if search_text:
        needle = search_text.casefold()
        def question_text(item):
            parts = [str(item.get("q", ""))]
            parts.extend(str(x) for x in item.get("options", []))
            parts.extend(str(st.get("text", "")) for st in item.get("statements", []))
            return " ".join(parts).casefold()
        filtered = [item for item in all_questions if needle in question_text(item)]
    else:
        filtered = list(all_questions)
    return render_template_string(
        QUESTION_BANK_TEMPLATE,
        subjects=subjects,
        selected_subject=selected_subject,
        search_text=search_text,
        questions=filtered,
        total_subject=len(all_questions),
        teacher_authenticated=bool(session.get("teacher_authenticated")),
        teacher_name=session.get("teacher_name") or session.get("judge_name", "Giám khảo")
    )


@app.route("/import-word", methods=["GET", "POST"])
def import_word():
    # Giám khảo được thử tải và xem trước đề Word; chỉ giáo viên mới được lưu vào ngân hàng thật.
    is_teacher = bool(session.get("teacher_authenticated"))
    if not is_teacher and not session.get("judge_name"):
        return redirect(url_for("home"))
    if request.method == "POST":
        subject = request.form.get("subject", "")
        file = request.files.get("word_file")
        action = request.form.get("action", "preview")
        if subject not in QUESTIONS:
            return "Môn học không hợp lệ. Thầy quay lại trang chủ nhé.", 400
        if not file or not file.filename.lower().endswith(".docx"):
            return "Thầy vui lòng chọn tệp Word .docx (Word 2007 trở lên).", 400
        try:
            document = Document(file)
            questions = read_word_questions(document)
            if not questions:
                return render_template_string(IMPORT_TEMPLATE, subjects=list(QUESTIONS.keys()),
                    selected_subject=subject, teacher_name=session.get("teacher_name", "Giáo viên"),
                    error="Chưa nhận diện được câu hỏi. Thầy kiểm tra số câu, định dạng A/B/C/D và từ khóa DẠNG: DUNGSAI.")
            if action == "save":
                if not is_teacher:
                    return "Chế độ giám khảo chỉ được kiểm tra và xem trước đề Word. Muốn lưu vào ngân hàng chính thức, vui lòng đăng nhập giáo viên.", 403
                added = save_imported_questions(subject, questions)
                mc_count = sum(1 for q in questions if q.get("type") != "true_false" and q.get("answer") is not None)
                tf_count = sum(1 for q in questions if q.get("type") == "true_false" and len(q.get("statements", [])) == 4)
                return render_template_string(IMPORT_RESULT_TEMPLATE, subject=subject, added=added,
                    mc_count=mc_count, tf_count=tf_count, total=len(questions), school=SCHOOL)
            mc_count = sum(1 for q in questions if q.get("type") != "true_false")
            tf_count = sum(1 for q in questions if q.get("type") == "true_false")
            valid_tf = sum(1 for q in questions if q.get("type") == "true_false" and len(q.get("statements", [])) == 4 and all(st.get("answer") is not None for st in q["statements"]))
            valid_mc = sum(1 for q in questions if q.get("type") != "true_false" and all(k in q.get("options", {}) for k in "ABCD") and q.get("answer") in ("A", "B", "C", "D"))
            return render_template_string(IMPORT_PREVIEW_TEMPLATE, subject=subject, subjects=list(QUESTIONS.keys()),
                teacher_name=session.get("teacher_name", "Giáo viên"), questions=questions, mc_count=mc_count, tf_count=tf_count, valid_mc=valid_mc,
                valid_tf=valid_tf, total=len(questions), school=SCHOOL, can_save=is_teacher)
        except Exception as e:
            return render_template_string(IMPORT_TEMPLATE, subjects=list(QUESTIONS.keys()),
                selected_subject=subject, teacher_name=session.get("teacher_name", "Giáo viên"), error="Không đọc được tệp Word: " + str(e))

    return render_template_string(IMPORT_TEMPLATE, subjects=list(QUESTIONS.keys()),
        selected_subject="Tin học", teacher_name=session.get("teacher_name", "Giáo viên"), error=None)


IMPORT_TEMPLATE = r"""
<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Nhập đề Word - SMART QUIZ THPT</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#f1f6fc;color:#172b4d;font-family:Arial,sans-serif;font-size:18px;line-height:1.55}
.topbar{background:linear-gradient(110deg,#0864ef,#0875f9);color:white;padding:18px 3%;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap}
.brand{font-size:26px;font-weight:800}.brand small{display:block;font-size:17px;font-weight:600;margin-top:3px}.nav{font-size:17px;font-weight:700}
.wrap{max-width:1480px;margin:20px auto;padding:0 20px}.panel{background:#fff;border-radius:22px;padding:30px;box-shadow:0 5px 22px #183b6410}
h1{font-size:36px;margin:0 0 6px;color:#17468f}.intro{font-size:20px;color:#526782;margin:0 0 25px}.formgrid{display:grid;grid-template-columns:250px minmax(260px,720px);gap:20px;align-items:center;margin:18px 0}.formgrid label{font-weight:700;font-size:20px}select,input[type=file]{font:inherit;width:100%;padding:14px;border:2px solid #d8e2ef;border-radius:10px;background:#fff;color:#172b4d}
button,.btn{font:inherit;font-weight:700;border:0;border-radius:11px;padding:15px 24px;cursor:pointer;text-decoration:none;display:inline-block;text-align:center}.blue{background:#0868f5;color:white}.green{background:#17ad67;color:white}.muted{background:#e8eef5;color:#263d58}.actions{display:flex;gap:16px;align-items:center;justify-content:space-between;flex-wrap:wrap;margin:22px 0}.back{color:#6335d7;font-weight:700}.notice{padding:17px 20px;border-radius:13px;background:#fff6df;border:1px solid #ffe2a0;margin:22px 0}.notice strong{display:block;color:#ad310e;font-size:21px}.examples{display:grid;grid-template-columns:1fr 1fr;gap:16px}.example{padding:20px;border-radius:14px;font-size:17px}.mc{background:#eaf5ff;border:1px solid #c5e4ff}.tf{background:#e5fff1;border:1px solid #b8f3d4}.example h2{font-size:23px;margin:0 0 12px}.sample{background:white;border:1px solid #dbe5f0;border-radius:10px;padding:16px}.sample b{color:#0a4ca2}.tf .sample b{color:#bd271b}.error{background:#fff0f0;border:1px solid #ffcaca;color:#9e1c1c;padding:14px;border-radius:10px}.small{font-size:15px;color:#5b6d82}@media(max-width:800px){.formgrid{grid-template-columns:1fr;gap:8px}.examples{grid-template-columns:1fr}.wrap{padding:0 10px}.panel{padding:20px}h1{font-size:29px}.topbar{padding:16px}.brand{font-size:21px}}
</style></head><body>
<header class="topbar"><div class="brand">SMART QUIZ THPT<small>HUỲNH TẤN PHÁT – VĨNH LONG</small></div><div class="nav"><a style="color:white;text-decoration:none" href="{{ url_for('home') }}">⌂ Trang chủ</a>　 <a style="color:white;text-decoration:none" href="{{ url_for('home') }}">▤ Tạo đề</a>　 <a style="color:white;text-decoration:none" href="{{ url_for('question_bank') }}">▦ Ngân hàng câu hỏi</a>　 <a style="color:white;text-decoration:none" href="{{ url_for('import_word') }}">▣ Nhập đề Word</a>　 ♙ {{ teacher_name or session.get('judge_name') or 'Người trải nghiệm' }}</div></header>
<div class="wrap"><main class="panel"><h1>▣ Nhập đề thi từ Word</h1><p class="intro">Chọn môn học và tải file đề thi định dạng .docx. Hệ thống sẽ đọc câu hỏi để thầy cô kiểm tra trước khi nhập vào ngân hàng.</p>
{% if error %}<div class="error">{{ error }}</div>{% endif %}
<form method="post" enctype="multipart/form-data"><div class="formgrid"><label for="subject">Chọn môn học:</label><select id="subject" name="subject" required>{% for s in subjects %}<option value="{{ s }}" {% if s==selected_subject %}selected{% endif %}>{{ s }}</option>{% endfor %}</select></div>
<div class="formgrid"><label for="word_file">Chọn đề Word (.docx):</label><div><input id="word_file" type="file" name="word_file" accept=".docx" required><div class="small">Chỉ hỗ trợ file .docx (Word 2007 trở lên).</div></div></div>
<div class="actions"><a class="back" href="{{ url_for('home') }}">⌂ Về trang chủ</a><button class="blue" type="submit" name="action" value="preview">☷ Kiểm tra tệp Word</button></div></form>
<div class="notice"><strong>💡 Hướng dẫn định dạng file Word</strong>Ngân hàng có thể gồm 40 câu trắc nghiệm và 12 câu Đúng/Sai. Mỗi lượt làm bài sẽ lấy ngẫu nhiên 24 câu trắc nghiệm và 4 câu Đúng/Sai.</div>
<div class="examples"><section class="example mc"><h2>① Dạng trắc nghiệm A, B, C, D</h2><div class="sample"><b>Câu 1.</b> Thiết bị nào là thiết bị nhập dữ liệu?<br>A. Màn hình<br>B. Bàn phím<br>C. Máy chiếu<br>D. Loa<br><b>Đáp án: B</b></div></section>
<section class="example tf"><h2>② Dạng Đúng/Sai</h2><div class="sample"><b>DẠNG: DUNGSAI</b><br>Câu 25. Xét các nhận định sau về Python:<br>a) Python là ngôn ngữ lập trình. [Đúng]<br>b) HTML dùng tạo cấu trúc trang web. [Đúng]<br>c) CSS là hệ quản trị cơ sở dữ liệu. [Sai]<br>d) Bàn phím là thiết bị xuất dữ liệu. [Sai]<br><b>Đáp án: a) Đúng; b) Đúng; c) Sai; d) Sai</b></div></section></div>
</main></div></body></html>
"""

IMPORT_PREVIEW_TEMPLATE = r"""
<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Kiểm tra đề Word</title>
<style>body{font-family:Arial,sans-serif;background:#f1f6fc;color:#172b4d;font-size:18px;line-height:1.5;margin:0}.wrap{max-width:1200px;margin:25px auto;padding:20px}.panel{background:white;padding:28px;border-radius:18px;box-shadow:0 5px 20px #0001}h1{color:#17468f}.stats{display:flex;gap:12px;flex-wrap:wrap}.stat{padding:14px 18px;background:#eaf3ff;border-radius:10px;font-weight:bold}.question{border:1px solid #dbe5f0;padding:18px;margin:14px 0;border-radius:12px}.tf{background:#f0fff6}.warn{color:#a33b00;background:#fff5df;padding:12px;border-radius:9px}.btn{padding:13px 20px;border:0;border-radius:9px;font-size:17px;font-weight:bold;cursor:pointer}.blue{background:#0868f5;color:#fff}.green{background:#16a765;color:#fff}.actions{display:flex;gap:12px;flex-wrap:wrap;margin-top:20px}select,input{font:inherit;padding:10px}a{color:#5d35c8}</style></head><body><div class="wrap"><div class="panel"><h1>Kiểm tra đề Word đã nhận dạng</h1><p>Môn học: <b>{{ subject }}</b></p><div class="stats"><div class="stat">Tổng câu: {{ total }}</div><div class="stat">Trắc nghiệm: {{ mc_count }}</div><div class="stat">Đúng/Sai: {{ tf_count }}</div></div>
{% if mc_count < 24 or tf_count < 4 %}<p class="warn">Ngân hàng cần tối thiểu 24 câu trắc nghiệm và 4 câu Đúng/Sai hợp lệ để tạo một lượt làm bài. Hiện file có {{ mc_count }} câu trắc nghiệm và {{ tf_count }} câu Đúng/Sai. Thầy có thể kiểm tra nhận dạng trước khi lưu.</p>{% endif %}
<p>Câu trắc nghiệm hợp lệ có đủ A/B/C/D và đáp án nhận diện: {{ valid_mc }}. Câu Đúng/Sai hợp lệ có 4 nhận định và đáp án đủ: {{ valid_tf }}.</p>
{% for q in questions %}<section class="question {% if q.type=='true_false' %}tf{% endif %}"><h3>Câu {{ q.number }}. {{ q.q }}</h3>{% if q.type=='true_false' %}<b>DẠNG ĐÚNG/SAI</b>{% for st in q.statements %}<p>{{ st.label }}) {{ st.text }} — Đáp án: {% if st.answer is none %}Chưa nhận diện{% elif st.answer %}Đúng{% else %}Sai{% endif %}</p>{% endfor %}{% else %}{% for label in ['A','B','C','D'] %}{% if q.options.get(label) %}<p>{{ label }}. {{ q.options[label] }}</p>{% endif %}{% endfor %}<b>Đáp án nhận diện: {{ q.answer or 'Chưa nhận diện' }}</b>{% endif %}</section>{% endfor %}
{% if can_save %}<form method="post" enctype="multipart/form-data"><input type="hidden" name="action" value="save"><label>Môn học: <select name="subject">{% for s in subjects %}<option value="{{ s }}" {% if s==subject %}selected{% endif %}>{{ s }}</option>{% endfor %}</select></label><p>Chọn lại đúng file Word vừa kiểm tra để xác nhận lưu vào ngân hàng chính thức.</p><input type="file" name="word_file" accept=".docx" required><div class="actions"><button class="green" type="submit">Nhập vào ngân hàng</button><a href="{{ url_for('import_word') }}">Chọn file khác</a><a href="{{ url_for('home') }}">Về trang chủ</a></div></form>{% else %}<div class="notice"><strong>Chế độ trải nghiệm dành cho giám khảo</strong>Thầy cô có thể tải tệp Word để kiểm tra hệ thống nhận dạng và xem trước kết quả. Để bảo vệ dữ liệu dự án, chức năng lưu vào ngân hàng chính thức chỉ dành cho giáo viên đăng nhập.</div><div class="actions"><a href="{{ url_for('import_word') }}">Thử tệp Word khác</a><a href="{{ url_for('home') }}">Về trang chủ</a></div>{% endif %}</div></div></body></html>
"""

IMPORT_RESULT_TEMPLATE = r"""
<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Kết quả nhập Word</title><style>body{font:18px Arial;background:#f1f6fc;color:#172b4d;padding:24px}.box{max-width:850px;margin:25px auto;background:white;padding:30px;border-radius:18px}h1{color:#14703d}.btn{display:inline-block;padding:14px 20px;background:#0868f5;color:white;text-decoration:none;border-radius:9px;margin:8px 8px 0 0}</style></head><body><div class="box"><h1>Đã hoàn tất bước nhập Word</h1><p>Môn học: <b>{{ subject }}</b></p><p>Số câu đã nhận dạng: {{ total }} ({{ mc_count }} trắc nghiệm, {{ tf_count }} Đúng/Sai).</p><p>Số câu mới đã thêm vào ngân hàng: <b>{{ added }}</b></p><p><strong>Lưu ý:</strong> Đây là ngân hàng câu hỏi. Mỗi lượt làm bài sẽ lấy ngẫu nhiên 24 câu trắc nghiệm và 4 câu Đúng/Sai nếu ngân hàng đủ câu hợp lệ.</p><a class="btn" href="{{ url_for('import_word') }}">Nhập đề khác</a><a class="btn" href="{{ url_for('home') }}">Về trang chủ</a></div></body></html>
"""


def _normalize_tf_answer(value):
    """Chuẩn hóa các cách ghi đáp án Đúng/Sai trong file Word."""
    if value is None:
        return None
    v = str(value).strip().lower()
    v = re.sub(r"^[\s\[\](){}:;,.\-–]+|[\s\[\](){}:;,.\-–]+$", "", v)
    # Bỏ dấu tiếng Việt để nhận cả Dung/Sai và Đúng/Sai.
    v_ascii = v.replace("đ", "d")
    if v in ("đúng", "true", "t") or v_ascii in ("dung",):
        return True
    if v in ("sai", "false", "f"):
        return False
    return None


def read_word_questions(document):
    """Đọc câu A/B/C/D và câu Đúng/Sai từ Word.

    Câu Đúng/Sai nên có dòng DẠNG: DUNGSAI trước phần câu hỏi. Đáp án
    có thể đặt ở cuối nhận định bằng [Đúng]/[Sai], hoặc ở dòng riêng như
    'Đáp án: a) Đúng;' sau mỗi nhận định, hoặc gom cả bốn ý trên một dòng.
    """
    questions = []
    current = None
    mode = "mc"

    def finish_current():
        nonlocal current
        if not current:
            return
        if current.get("type") == "true_false":
            if current.get("statements"):
                questions.append(current)
        elif current.get("options"):
            questions.append(current)
        current = None

    def apply_tf_answer_line(line):
        """Cập nhật đáp án từ dòng khóa đáp án; trả True nếu đã xử lý."""
        if not re.match(r"^\s*(?:Đáp\s*án|Dap\s*an)\s*:", line, re.I):
            return False
        # Chấp nhận a) Đúng; b) Sai; c. Đúng; d: Sai.
        pattern = r"\b([a-dA-D])\s*[).:]?\s*(?:[-–]\s*)?(Đúng|Dung|Sai|True|False)\b"
        for label, value in re.findall(pattern, line, re.I):
            answer = _normalize_tf_answer(value)
            for st in current.get("statements", []):
                if st.get("label") == label.lower():
                    st["answer"] = answer
        return True

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        normalized = re.sub(r"\s+", " ", text).strip().upper()

        if re.match(r"^DẠNG\s*:\s*DUNGSAI\s*$", normalized) or re.match(r"^DANG\s*:\s*DUNGSAI\s*$", normalized):
            finish_current()
            mode = "true_false"
            continue
        if re.match(r"^DẠNG\s*:\s*(?:TRACNGHIEM|TRẮC NGHIỆM)\s*$", normalized) or re.match(r"^DANG\s*:\s*TRACNGHIEM\s*$", normalized):
            finish_current()
            mode = "mc"
            continue

        qmatch = re.match(r"^\s*Câu\s*(\d+)\s*[.:]?\s*(.*)$", text, re.I)
        if qmatch:
            finish_current()
            current = {
                "number": int(qmatch.group(1)),
                "q": qmatch.group(2).strip(),
                "options": {},
                "answer": None,
                "type": "true_false" if mode == "true_false" else "multiple_choice",
                "statements": [] if mode == "true_false" else None,
            }
            continue
        if current is None:
            continue

        if current.get("type") == "true_false":
            # Kiểm tra dòng đáp án trước khi xét dòng nhận định.
            if apply_tf_answer_line(text):
                continue

            sm = re.match(r"^\s*([a-dA-D])\s*[).:]\s*(.*)$", text)
            if sm:
                label = sm.group(1).lower()
                body = sm.group(2).strip()
                am = re.search(r"\[\s*(Đúng|Dung|Sai|True|False)\s*\]", body, re.I)
                answer = _normalize_tf_answer(am.group(1)) if am else None
                if am:
                    body = (body[:am.start()] + body[am.end():]).strip()
                # Nếu dòng nhận định bị lặp nhãn, cập nhật nhận định cũ thay vì tạo bản trùng.
                existing = next((st for st in current["statements"] if st.get("label") == label), None)
                if existing:
                    existing["text"] = body or existing["text"]
                    if answer is not None:
                        existing["answer"] = answer
                else:
                    current["statements"].append({"label": label, "text": body, "answer": answer, "explain": ""})
                continue

            if re.match(r"^(Giải thích|Giai thich)\s*:", text, re.I):
                continue
            if current["statements"]:
                current["statements"][-1]["text"] = (current["statements"][-1]["text"] + " " + text).strip()
            else:
                current["q"] = (current["q"] + " " + text).strip()
            continue

        # Câu trắc nghiệm A/B/C/D.
        if re.match(r"^Đáp\s*án\s*:", text, re.I):
            am = re.search(r"Đáp\s*án\s*:\s*([ABCD])\b", text, re.I)
            if am:
                current["answer"] = am.group(1).upper()
            continue
        markers = []
        full_text = paragraph.text
        offset = 0
        for run in paragraph.runs:
            run_text = run.text
            for mm in re.finditer(r"(?<![A-Za-z])([ABCD])\s*[.)]\s*", run_text, re.I):
                markers.append({"label": mm.group(1).upper(), "start": offset + mm.start(), "end": offset + mm.end(), "underline": bool(run.underline)})
            offset += len(run_text)
        if markers:
            for i, marker in enumerate(markers):
                end_pos = markers[i + 1]["start"] if i + 1 < len(markers) else len(full_text)
                option_text = full_text[marker["end"]:end_pos].strip(" \t\r\n.")
                label = marker["label"]
                current["options"][label] = (current["options"].get(label, "") + " " + option_text).strip()
                if marker["underline"]:
                    current["answer"] = label
        elif not current["options"] and not re.match(r"^Đáp\s*án\s*:", text, re.I):
            current["q"] = (current["q"] + " " + text).strip()

    finish_current()
    return questions


def save_imported_questions(subject, questions):
    saved_questions = {}
    if os.path.exists(QUESTION_BANK_FILE):
        try:
            with open(QUESTION_BANK_FILE, "r", encoding="utf-8") as f:
                saved_questions = json.load(f)
        except (json.JSONDecodeError, OSError):
            saved_questions = {}
    if subject not in QUESTIONS:
        raise ValueError("Môn học không hợp lệ.")
    saved_questions.setdefault(subject, [])
    existing_ids = {q.get("id") for q in QUESTIONS[subject]}
    saved_ids = {q.get("id") for q in saved_questions[subject]}
    label_to_index = {"A": 0, "B": 1, "C": 2, "D": 3}
    added = 0
    for q in questions:
        new_id = f"word_{subject}_{len(saved_questions[subject]) + 1}_{q['number']}"
        if new_id in existing_ids or new_id in saved_ids:
            # Generate a distinct suffix rather than silently dropping a question.
            suffix = 2
            base = new_id
            while new_id in existing_ids or new_id in saved_ids:
                new_id = f"{base}_{suffix}"
                suffix += 1
        if q.get("type") == "true_false":
            statements = q.get("statements", [])
            if len(statements) != 4 or any(st.get("answer") is None or not st.get("text") for st in statements):
                continue
            item = {
                "id": new_id, "type": "true_false", "q": q.get("q", "").strip(),
                "statements": [{"text": st["text"].strip(), "answer": bool(st["answer"]), "explain": st.get("explain", "")} for st in statements],
                "topic": "Nhập từ Word", "explain": "Câu hỏi được nhập từ đề Word.",
                "review": "Ôn lại các nhận định và kiến thức liên quan."
            }
        else:
            if not q.get("q") or not all(label in q.get("options", {}) and q["options"][label].strip() for label in "ABCD"):
                continue
            answer_label = q.get("answer")
            if answer_label not in label_to_index:
                continue
            item = {
                "id": new_id, "q": q["q"].strip(),
                "options": [q["options"][label].strip() for label in "ABCD"],
                "answer": label_to_index[answer_label], "topic": "Nhập từ Word",
                "explain": "Câu hỏi được nhập từ đề Word.",
                "review": "Ôn lại câu hỏi và kiến thức liên quan."
            }
        saved_questions[subject].append(item)
        saved_ids.add(new_id)
        added += 1
    with open(QUESTION_BANK_FILE, "w", encoding="utf-8") as f:
        json.dump(saved_questions, f, ensure_ascii=False, indent=2)
    # Update in-memory questions for immediate use in this local session.
    for item in saved_questions.get(subject, []):
        if item.get("id") not in {q.get("id") for q in QUESTIONS[subject]}:
            QUESTIONS[subject].append(item)
    return added



if __name__ == "__main__": 

  app.run(host="127.0.0.1", port=5003, debug=False)



