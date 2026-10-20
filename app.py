
from flask import Flask, render_template_string, request, redirect, url_for, session
from docx import Document
import re
import os
import json
app = Flask(__name__)
app.secret_key = "smart-quiz-thpt-local-key"

SCHOOL = "HUỲNH TẤN PHÁT – VĨNH LONG"

QUESTIONS = {
    "Tin học": [
        {
            "id": "tin1",
            "q": "Python là gì?",
            "options": ["Một hệ điều hành", "Một ngôn ngữ lập trình",
                        "Một thiết bị phần cứng", "Một trình duyệt web"],
            "answer": 1,
            "topic": "Ngôn ngữ lập trình",
            "explain": "Python là ngôn ngữ lập trình bậc cao.",
            "review": "Ôn lại khái niệm ngôn ngữ lập trình và đặc điểm của Python."
        },
        {
            "id": "tin2",
            "q": "HTML chủ yếu được dùng để làm gì?",
            "options": ["Tạo cấu trúc trang web", "Diệt virus",
                        "Thay thế hệ điều hành", "Tăng dung lượng RAM"],
            "answer": 0,
            "topic": "Thiết kế web",
            "explain": "HTML dùng để xây dựng cấu trúc nội dung trang web.",
            "review": "Ôn lại vai trò của HTML và phân biệt HTML với CSS, JavaScript."
        },
        {
            "id": "tin3",
            "q": "Thiết bị nào sau đây là thiết bị nhập dữ liệu?",
            "options": ["Màn hình", "Loa", "Bàn phím", "Máy chiếu"],
            "answer": 2,
            "topic": "Phần cứng máy tính",
            "explain": "Bàn phím cho phép người dùng nhập dữ liệu vào máy tính.",
            "review": "Ôn lại cách phân loại thiết bị nhập và thiết bị xuất."
        }
    ],
    "Toán học": [
        {
            "id": "toan1",
            "q": "Nghiệm của phương trình x + 5 = 12 là gì?",
            "options": ["5", "6", "7", "8"],
            "answer": 2,
            "topic": "Phương trình",
            "explain": "Ta có x = 12 - 5 = 7.",
            "review": "Ôn lại các phép biến đổi phương trình bậc nhất."
        },
        {
            "id": "toan2",
            "q": "Giá trị của 3² + 4 là bao nhiêu?",
            "options": ["9", "12", "13", "16"],
            "answer": 2,
            "topic": "Lũy thừa",
            "explain": "3² + 4 = 9 + 4 = 13.",
            "review": "Ôn lại thứ tự thực hiện phép tính và lũy thừa."
        }
    ],
    "Tiếng Anh": [
        {
            "id": "anh1",
            "q": "Choose the correct answer: She _____ a student.",
            "options": ["am", "is", "are", "be"],
            "answer": 1,
            "topic": "Ngữ pháp",
            "explain": "Với chủ ngữ She, động từ to be ở hiện tại là is.",
            "review": "Ôn cách dùng am, is, are theo chủ ngữ."
        },
        {
            "id": "anh2",
            "q": "What is the opposite of 'big'?",
            "options": ["Tall", "Long", "Small", "High"],
            "answer": 2,
            "topic": "Từ vựng",
            "explain": "Small có nghĩa là nhỏ, trái nghĩa với big.",
            "review": "Ôn từ vựng chỉ kích thước và các cặp từ trái nghĩa."
        }
    ],
    "Ngữ văn": [
        {
            "id": "van1",
            "q": "Tác giả của Truyện Kiều là ai?",
            "options": ["Nguyễn Du", "Nguyễn Trãi", "Nam Cao", "Tố Hữu"],
            "answer": 0,
            "topic": "Tác giả và tác phẩm",
            "explain": "Truyện Kiều là kiệt tác của đại thi hào Nguyễn Du.",
            "review": "Ôn lại tác giả, hoàn cảnh sáng tác và giá trị Truyện Kiều."
        }
    ],
    "Vật lí": [
        {
            "id": "li1",
            "q": "Đơn vị đo lực trong hệ SI là gì?",
            "options": ["Jun", "Oát", "Niutơn", "Vôn"],
            "answer": 2,
            "topic": "Đại lượng và đơn vị",
            "explain": "Đơn vị đo lực trong hệ SI là niutơn (N).",
            "review": "Ôn lại các đại lượng vật lí và đơn vị đo tương ứng."
        }
    ],
    "Hóa học": [
        {
            "id": "hoa1",
            "q": "Kí hiệu hóa học của oxi là gì?",
            "options": ["Ox", "O", "Oi", "H"],
            "answer": 1,
            "topic": "Nguyên tố hóa học",
            "explain": "Kí hiệu hóa học của nguyên tố oxi là O.",
            "review": "Ôn lại kí hiệu hóa học của các nguyên tố thường gặp."
        }
    ]
}

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

HOME_TEMPLATE = """
<!DOCTYPE html><html lang="vi"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SMART QUIZ THPT - HUỲNH TẤN PHÁT</title>
""" + STYLE + """
</head><body><div class="container">
""" + HEADER + """
<div class="card">
    <h2>Chào mừng thầy cô và các em học sinh!</h2>
    <p>Chọn môn học để bắt đầu ôn tập.</p>
    <div class="subjects">
        {% for subject in subjects %}
        <a class="subject" href="{{ url_for('quiz', subject=subject) }}">
            {{ subject }}<br><small>Bắt đầu ôn tập →</small>
        </a>
        {% endfor %}
    </div>
</div>

<div class="card">
    <h2>Nhập đề thi từ Word</h2>
    <p>Nhập câu hỏi trắc nghiệm A, B, C, D từ đề thi McMIX.</p>
    <a class="btn secondary" href="{{ url_for('import_word') }}">Nhập đề Word</a>
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
""" + HEADER + """
<div class="card">
    <a class="btn secondary" href="{{ url_for('home') }}">← Chọn môn khác</a>
    <h2>Ôn tập môn {{ subject }}</h2>

    {% if retry_mode %}
    <p><b>Luyện tập lại các câu trả lời sai</b></p>
    {% else %}
    <p>Số câu hỏi: {{ questions|length }}</p>
    {% endif %}

    {% if result is none %}
    <form method="post">
        <input type="hidden" name="retry_mode" value="{{ 1 if retry_mode else 0 }}">
        {% for question in questions %}
        {% set qi = loop.index0 %}
        <div style="margin: 22px 0">
            <h3>Câu {{ loop.index }}: {{ question.q }}</h3>
            {% for option in question.options %}
            <label class="option">
                <input type="radio" name="q_{{ question.id }}"
                       value="{{ loop.index0 }}">
                {{ ["A", "B", "C", "D"][loop.index0] }}. {{ option }}
            </label>
            {% endfor %}
        </div>
        {% endfor %}
        <button class="btn" type="submit">Nộp bài và xem kết quả</button>
    </form>

    {% else %}
    <div class="result">
        <h2>Kết quả bài làm</h2>
        <p>Số câu đúng: <b>{{ result.score }}/{{ result.total }}</b></p>
        <p>Điểm số:
            <b>{{ "%.2f"|format(result.score * 10 / result.total) }}/10</b>
        </p>
    </div>

    <div class="analysis">
        <h3>Phân tích kiến thức cần ôn tập</h3>
        {% if result.wrong %}
            <p>Ứng dụng ghi nhận {{ result.wrong|length }} câu chưa chính xác.</p>
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
        <b>{{ loop.index }}. {{ item.q }}</b>
        <p>Chủ đề: {{ item.topic }}</p>
        <p>Bài làm: {{ item.selected }}</p>
        <p>Đáp án đúng: <b>{{ item.correct_answer }}</b></p>
        <p style="color: {{ 'green' if item.correct else 'red' }}">
            {{ 'Chính xác' if item.correct else 'Chưa chính xác' }}
        </p>
        <p>Giải thích: {{ item.explain }}</p>
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


@app.route("/")
def home():
    return render_template_string(
        HOME_TEMPLATE,
        subjects=list(QUESTIONS.keys())
    )


@app.route("/quiz/<path:subject>", methods=["GET", "POST"])
def quiz(subject):
    if subject not in QUESTIONS:
        return redirect(url_for("home"))

    retry_mode = (
        request.args.get("retry") == "1"
        or request.form.get("retry_mode") == "1"
    )

    if retry_mode:
        wrong_ids = session.get("wrong_by_subject", {}).get(subject, [])
        questions = [
            q for q in QUESTIONS[subject] if q["id"] in wrong_ids
        ]
        if not questions:
            return redirect(url_for("quiz", subject=subject))
    else:
        questions = QUESTIONS[subject]

    result = None

    if request.method == "POST":
        score = 0
        details = []
        wrong = []
        topics = {}

        for question in questions:
            selected = request.form.get("q_" + question["id"])
            try:
                selected_index = int(selected)
            except (TypeError, ValueError):
                selected_index = -1

            correct = selected_index == question["answer"]
            if correct:
                score += 1
            else:
                wrong.append(question["id"])
                topics.setdefault(question["topic"], []).append(question)

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
                "explain": question["explain"]
            })

        session["wrong_by_subject"] = {
            **session.get("wrong_by_subject", {}),
            subject: wrong
        }

        result = {
            "score": score,
            "total": len(questions),
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
        school=SCHOOL
    )

@app.route("/import-word", methods=["GET", "POST"])
def import_word():
    if request.method == "POST":
        subject = request.form.get("subject", "")
        file = request.files.get("word_file")
        action = request.form.get("action", "preview")
        if subject not in QUESTIONS:
            return "Môn học không hợp lệ. Thầy quay lại trang chủ nhé."
        if action == "save":
            if not file or not file.filename.lower().endswith(".docx"):
                return "Thầy vui lòng chọn lại tệp Word .docx để lưu."

            try:
                document = Document(file)
                questions = read_word_questions(document)
                added = save_imported_questions(subject, questions)
                return render_template_string("""
                    <h2>Kết quả nhập ngân hàng</h2>
                    <p>Môn học: {{ subject }}</p>
                    <p>Số câu mới đã lưu: {{ added }}</p>
                    <p><a href="{{ url_for('import_word') }}">Nhập đề khác</a></p>
                    <p><a href="{{ url_for('home') }}">Về trang chủ</a></p>
                """, subject=subject, added=added)
            except Exception as e:
                return "Không lưu được câu hỏi: " + str(e)

        if subject not in QUESTIONS:
            return "Môn học không hợp lệ. Thầy quay lại trang chủ nhé."

        if not file or not file.filename.lower().endswith(".docx"):
            return "Thầy vui lòng chọn tệp Word có đuôi .docx."

        try:
            document = Document(file)
            questions = read_word_questions(document)
            return render_template_string("""
    <h2>Kết quả đọc đề Word</h2>
    <p>Môn học: {{ subject }}</p>
    <p>Số đoạn văn bản: {{ count }}</p>
    <p>Số câu hỏi nhận diện được: {{ questions|length }}</p>

    {% for q in questions %}
    <hr>
    <h3>Câu {{ q.number }}. {{ q.q }}</h3>
    {% for label in ['A', 'B', 'C', 'D'] %}
        {% if label in q.options %}
        <p>{{ label }}. {{ q.options[label] }}</p>
        {% endif %}
    {% endfor %}
    <p><strong>Đáp án nhận diện:</strong> {{ q.answer or 'Chưa nhận diện' }}</p>
    {% endfor %}

    <hr>
    <p><strong>Đã đọc {{ questions|length }} câu hỏi.</strong></p>
    <p>Chưa lưu vào ngân hàng. Cần xác nhận ở bước tiếp theo.</p>
    
    <form method="post" enctype="multipart/form-data">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="subject" value="{{ subject }}">
    <p>
        <label>Chọn lại tệp Word để xác nhận lưu:</label>
        <input type="file" name="word_file" accept=".docx" required>
    </p>
    <button type="submit">Xác nhận nhập vào ngân hàng</button>
    </form>

    <p><a href="{{ url_for('import_word') }}">Thử tệp khác</a></p>
    <p><a href="{{ url_for('home') }}">Về trang chủ</a></p>
""", subject=subject, count=len(document.paragraphs),
     questions=questions)
        except Exception as e:
            return "Không đọc được tệp Word. Thầy kiểm tra lại tệp nhé: " + str(e)

    return render_template_string("""
        <h2>Nhập đề thi từ Word</h2>
        <form method="post" enctype="multipart/form-data">
            <p>
                <label>Chọn môn học:</label>
                <select name="subject" required>
                    {% for s in subjects %}
                    <option value="{{ s }}">{{ s }}</option>
                    {% endfor %}
                </select>
            </p>
            <p>
                <label>Chọn đề Word (.docx):</label>
                <input type="file" name="word_file"
                       accept=".docx" required>
            </p>
            <button type="submit">Kiểm tra tệp Word</button>
        </form>
        <p><a href="{{ url_for('home') }}">Về trang chủ</a></p>
    """, subjects=list(QUESTIONS.keys()))


def read_word_questions(document):
    questions = []
    current = None

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        match = re.match(r"^\s*Câu\s*(\d+)\s*[.:]?\s*(.*)$", text, re.I)

        if match:
            if current and current["options"]:
                questions.append(current)

            current = {
                "number": int(match.group(1)),
                "q": match.group(2).strip(),
                "options": {},
                "answer": None,
            }
            continue

        if current is None:
            continue

        # Nhận diện nhãn A., B., C., D. trong từng đoạn Word.
        markers = []
        full_text = paragraph.text
        offset = 0

        for run in paragraph.runs:
            run_text = run.text
            for label_match in re.finditer(
                r"(?<![A-Za-z])([ABCD])\s*\.",
                run_text,
                re.I
            ):
                markers.append({
                    "label": label_match.group(1).upper(),
                    "start": offset + label_match.start(),
                    "end": offset + label_match.end(),
                    "underline": bool(run.underline),
                })
            offset += len(run_text)

        for i, marker in enumerate(markers):
            end = markers[i + 1]["start"] if i + 1 < len(markers) else len(full_text)
            option_text = full_text[marker["end"]:end].strip(" \t\r\n.")
            label = marker["label"]

            current["options"][label] = (
                current["options"].get(label, "") + " " + option_text
            ).strip()

            if marker["underline"]:
                current["answer"] = label

        # Nếu đoạn không có nhãn A-D, nối tiếp vào câu hỏi hoặc phương án.
        if not markers and not current["options"]:
            current["q"] = (current["q"] + " " + text).strip()

    if current and current["options"]:
        questions.append(current)

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

    existing_ids = {
        q["id"] for q in QUESTIONS[subject]
    }

    saved_questions.setdefault(subject, [])
    saved_ids = {
        q.get("id") for q in saved_questions[subject]
    }

    label_to_index = {"A": 0, "B": 1, "C": 2, "D": 3}
    added = 0

    for q in questions:
        if not q.get("q") or not all(
            label in q.get("options", {})
            for label in ["A", "B", "C", "D"]
        ):
            continue

        answer_label = q.get("answer")
        if answer_label not in label_to_index:
            continue

        new_id = f"word_{subject}_{len(saved_questions[subject]) + 1}_{q['number']}"
        if new_id in existing_ids or new_id in saved_ids:
            continue

        item = {
            "id": new_id,
            "q": q["q"],
            "options": [
                q["options"]["A"],
                q["options"]["B"],
                q["options"]["C"],
                q["options"]["D"],
            ],
            "answer": label_to_index[answer_label],
            "topic": "Nhập từ Word",
            "explain": "Câu hỏi được nhập từ đề Word.",
            "review": "Ôn lại câu hỏi và kiến thức liên quan."
        }

        saved_questions[subject].append(item)
        saved_ids.add(new_id)
        added += 1

    with open(QUESTION_BANK_FILE, "w", encoding="utf-8") as f:
        json.dump(saved_questions, f, ensure_ascii=False, indent=2)

    return added




if __name__ == "__main__": 

  app.run(host="127.0.0.1", port=5000, debug=False)



