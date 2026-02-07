from flask import Flask, render_template_string, request, session, redirect, url_for
import pandas as pd
import random
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'super-secret-key-change-this-2025'

# ────────────────────────────────────────────────
# ────────────────────────────────────────────────
BASE_DIR    = r"C:\Users\Santhosh kumar D\OneDrive\Desktop\kt"
EXCEL_FILE  = os.path.join(BASE_DIR, "questions.xlsx")
RESULT_FILE = os.path.join(BASE_DIR, "result.xlsx")
# ────────────────────────────────────────────────

def load_questions():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(
            f"questions.xlsx not found in:\n{BASE_DIR}\n\n"
            "Format required (no header row):\n"
            "Question text\tOption A\tOption B\tOption C\tOption D\tAnswer (A/B/C/D)"
        )

    df = pd.read_excel(EXCEL_FILE, header=None)

    if df.shape[1] < 6:
        raise ValueError("Excel must have 6 columns: question + 4 options + answer letter")

    questions = []
    for _, row in df.iterrows():
        try:
            q = str(row[0]).strip()
            if not q:
                continue

            opts = [str(row[i]).strip() for i in range(1, 5)]
            correct = str(row[5]).strip().upper()

            if correct not in 'ABCD' or not all(opts):
                continue

            questions.append({
                'question': q,
                'options': opts,
                'correct': correct
            })
        except:
            continue

    if not questions:
        raise ValueError("No valid questions found in Excel. Check format/content.")

    print(f"Loaded {len(questions)} questions successfully.")
    return questions


# ────────────────────────────────────────────────
#               HTML TEMPLATES
# ────────────────────────────────────────────────

LOGIN_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>iMatiz Technology </title>
    <style>
        body {font-family:Arial,sans-serif; background:#f8f9fa; display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}
        .card {background:white;padding:50px 40px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.15);max-width:420px;text-align:center;}
        h1 {color:#2c3e50;margin-bottom:20px;}
        .msg {color:#dc3545; font-weight:bold; margin-bottom:15px;}
        input,button {width:100%;padding:14px;font-size:18px;margin:12px 0;border-radius:6px;box-sizing:border-box;}
        button {background:#28a745;color:white;border:none;cursor:pointer;}
        button:hover {background:#218838;}
    </style>
</head>
<body>
<div class="card">
    <h1>iMatiz Technology</h1>
    {% if kicked_msg %}<div class="msg">{{ kicked_msg | safe }}</div>{% endif %}
    <form method="post">
        <input type="text" name="name" placeholder="Enter your full name" required autofocus>
        <button type="submit">Start Test</button>
    </form>
</div>
</body>
</html>
"""

QUESTION_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Question {{ qnum }} / {{ total }}</title>
    <style>
        body {font-family:Arial,sans-serif;background:#f8f9fa;margin:0;padding:20px;}
        .container {max-width:820px;margin:auto;background:white;padding:30px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.12);}
        .timer {font-size:2.6rem;font-weight:bold;color:#dc3545;text-align:center;margin-bottom:1rem;}
        h2 {text-align:center;color:#2c3e50;margin-bottom:1.5rem;}
        .question {font-size:1.4rem;line-height:1.6;margin:2rem 0;color:#34495e;}
        label {display:block;margin:1.3rem 0;padding:1rem;font-size:1.22rem;border-radius:8px;cursor:pointer;transition:0.18s;}
        label:hover {background:#f1f3f5;}
        input[type=radio] {transform:scale(1.5);margin-right:14px;}
        button {display:block;margin:2.5rem auto 0;padding:0.9rem 3rem;font-size:1.25rem;background:#007bff;color:white;border:none;border-radius:8px;cursor:pointer;}
        button:hover {background:#0069d9;}
    </style>
</head>
<body onload="startTimer();">

<div class="container">
    <div class="timer" id="timer">10 seconds</div>
    <h2>Question {{ qnum }} of {{ total }}</h2>
    <div class="question">{{ question }}</div>

    <form method="post" id="form">
        {% for opt in options %}
        <label>
            <input type="radio" name="ans" value="{{ 'ABCD'[loop.index0] }}" required>
            {{ opt }}
        </label>
        {% endfor %}
        <button type="submit">Next →</button>
    </form>
</div>

<script>
let time = 10;
let timer = setInterval(() => {
    time--;
    document.getElementById("timer").innerText = time + " seconds";
    if (time <= 0) {
        clearInterval(timer);
        document.getElementById("form").submit();
    }
}, 1000);

document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
        alert("Tab switch or minimize detected.\\nTest terminated - partial score saved.");
        window.location.href = "/tab_cheat_end";
    }
});
</script>
</body>
</html>
"""

RESULT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Test Completed</title>
    <style>
        body {font-family:Arial,sans-serif;background:#f8f9fa;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}
        .card {background:white;padding:50px;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,0.15);text-align:center;max-width:500px;width:90%;}
        h1 {color:#28a745;}
        .score {font-size:5rem;font-weight:bold;color:#007bff;margin:1.5rem 0;}
        a {display:inline-block;margin-top:2.5rem;padding:14px 40px;background:#007bff;color:white;text-decoration:none;border-radius:8px;font-size:1.3rem;}
        a:hover {background:#0069d9;}
    </style>
</head>
<body>
<div class="card">
    <h1>Test Completed</h1>
    <p style="font-size:1.6rem;">{{ name }}</p>
    <div class="score">{{ score }} / {{ total }}</div>
    <p style="color:#555;font-size:1.3rem;">Your result has been recorded.</p>
    <a href="/">Back to Login</a>
</div>
</body>
</html>
"""

# ────────────────────────────────────────────────
#                   ROUTES
# ────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    kicked_msg = ""
    if request.args.get('terminated') == 'yes':
        kicked_msg = "Previous session was terminated due to tab switch or timeout.<br>Please start again."

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if not name:
            return "<h2 style='color:red;text-align:center'>Please enter your name</h2>", 400

        try:
            questions = load_questions()
        except Exception as e:
            return f"""
            <h2 style="color:red;text-align:center">Error loading questions</h2>
            <pre style="background:#f8d7da;padding:15px;border-radius:6px;max-width:800px;margin:20px auto;">{str(e)}</pre>
            <p style="text-align:center"><a href="/">Try again</a></p>
            """, 500

        random.shuffle(questions)

        session['name']     = name
        session['questions'] = questions
        session['current']   = 0
        session['answers']   = {}

        return redirect('/test')

    return render_template_string(LOGIN_HTML, kicked_msg=kicked_msg)


@app.route('/test', methods=['GET', 'POST'])
def test():
    if 'questions' not in session:
        return redirect('/')

    if session['current'] >= len(session['questions']):
        return redirect('/result')

    if request.method == 'POST':
        ans = request.form.get('ans')
        session['answers'][str(session['current'])] = ans
        session['current'] += 1
        return redirect('/test')

    q = session['questions'][session['current']]
    return render_template_string(QUESTION_HTML,
                                 qnum=session['current'] + 1,
                                 total=len(session['questions']),
                                 question=q['question'],
                                 options=q['options'])


@app.route('/result')
def result():
    if 'questions' not in session:
        return redirect('/')

    questions = session['questions']
    answers   = session['answers']

    # Number of answered questions = how many times "Next" was clicked
    answered_count = len(answers)

    # Score (correct among answered)
    score = 0
    for i in range(len(questions)):
        if str(i) in answers and answers[str(i)] == questions[i]['correct']:
            score += 1

    name  = session.get('name', 'Unknown')
    total = len(questions)
    unanswered = total - answered_count  # <-- this is the real unanswered count

    row = {
        'Employee Name': name,
        'Correct Answers': score,
        'Answered Questions': answered_count,
        'Unanswered Questions': unanswered,
        'Total Questions': total,
        'Date & Time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'Status': 'Completed'
    }

    if os.path.exists(RESULT_FILE):
        df = pd.read_excel(RESULT_FILE)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        df = pd.DataFrame([row])

    df.to_excel(RESULT_FILE, index=False)

    session.clear()

    return render_template_string(RESULT_HTML, name=name, score=score, total=total)


@app.route('/tab_cheat_end')
def tab_cheat_end():
    if 'questions' not in session:
        return redirect('/')

    questions = session['questions']
    answers   = session['answers']

    # Answered = number of submitted answers
    answered_count = len(answers)

    # Score only from answered ones
    score = 0
    for i in range(len(questions)):
        if str(i) in answers and answers[str(i)] == questions[i]['correct']:
            score += 1

    name  = session.get('name', 'Unknown')
    total = len(questions)
    unanswered = total - answered_count

    row = {
        'Employee Name': f"{name} (Terminated - Tab Switch)",
        'Correct Answers': score,
        'Answered Questions': answered_count,
        'Unanswered Questions': unanswered,
        'Total Questions': total,
        'Date & Time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'Status': 'Terminated'
    }

    if os.path.exists(RESULT_FILE):
        df = pd.read_excel(RESULT_FILE)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        df = pd.DataFrame([row])

    df.to_excel(RESULT_FILE, index=False)

    session.clear()

    return redirect('/?terminated=yes')


if __name__ == '__main__':
    print("\n" + "═"*70)
    print("   iMatiz")
    # print(f"   Questions: {EXCEL_FILE}")
    # print(f"   Results saved to: {RESULT_FILE}")
    # print("═"*70)
    print("\nOpen →  http://127.0.0.1:5000")
    print("Network → http://<your-ip>:5000   (ipconfig → IPv4 Address)")
    print()

    app.run(host='0.0.0.0', port=5000, debug=True)