from flask import Flask, render_template_string, request, session, redirect, url_for
import pandas as pd
import random
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'super-secret-key-change-this-2025'

# ────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE  = os.path.join(BASE_DIR, "questions.xlsx")
RESULT_FILE = os.path.join(BASE_DIR, "result.xlsx")

ALLOWED_EMPLOYEES = ["Santhosh", "Janani G", "Aishwrya","Satish","Zaiba","Guru Divya","Aarthi","Vashanth","Abinaya","Suchithra","Dhana priya","Dhanya","Nivetha","Shreyas","Siri","Ananaya","Ashwini","Gopika","Sridharshini","Kavi Keerthna","Ramya","Priya","Keerthana","Nagraj","Vimal"]

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
            if not q: continue
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
    <title>iMatiz Technology</title>
    <style>
        body {font-family:Arial,sans-serif; background:#f8f9fa; display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}
        .card {background:white;padding:50px 40px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.15);max-width:420px;text-align:center;}
        h1 {color:#2c3e50;margin-bottom:20px;}
        .msg {color:#dc3545; font-weight:bold; margin-bottom:15px;}
        select,button {width:100%;padding:14px;font-size:18px;margin:12px 0;border-radius:6px;box-sizing:border-box;}
        button {background:#28a745;color:white;border:none;cursor:pointer;}
        button:hover {background:#218838;}
        .blocked {color:#dc3545; font-weight:bold; margin:25px 0; line-height:1.5; display:none;}
    </style>
</head>
<body>
<div class="card">
    <h1>iMatiz Technology</h1>
    {% if kicked_msg %}<div class="msg">{{ kicked_msg | safe }}</div>{% endif %}
    
    <div id="blocked-msg" class="blocked">
        This test was terminated earlier (tab switch / timeout).<br>
        You are no longer allowed to restart the test in this browser.<br>
        Contact admin / HR if needed.
    </div>

    <form method="post" id="login-form">
        <select name="name" id="name-select" required autofocus>
            <option value="" disabled selected>Select your name</option>
            {% for emp in employees %}
            <option value="{{ emp }}">{{ emp }}</option>
            {% endfor %}
        </select>
        <button type="submit" id="start-btn">Start Test</button>
    </form>
</div>

<script>
    const nameSelect = document.getElementById('name-select');
    const blockedMsg = document.getElementById('blocked-msg');
    const form = document.getElementById('login-form');

    function checkLock(name) {
        if (!name) return;
        const isLocked = localStorage.getItem('quiz_locked_' + name) === '1';
        if (isLocked) {
            blockedMsg.style.display = 'block';
            form.style.display = 'none';
        } else {
            blockedMsg.style.display = 'none';
            form.style.display = 'block';
        }
    }

    nameSelect.addEventListener('change', () => checkLock(nameSelect.value));
    // Check on load in case name is pre-filled (rare)
    if (nameSelect.value) checkLock(nameSelect.value);
</script>
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

let tabSwitchDetected = false;

function markTerminated() {
    localStorage.setItem('quiz_locked_' + '{{ name|e }}', '1');
}

document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden" && !tabSwitchDetected) {
        tabSwitchDetected = true;
        markTerminated();
        alert("Tab switch / minimize detected.\\nTest terminated.\\nYou cannot restart.");
        window.location.href = "/tab_cheat_end";
    }
});

window.addEventListener("blur", () => {
    if (!tabSwitchDetected) {
        tabSwitchDetected = true;
        markTerminated();
        alert("Window lost focus.\\nTest terminated.\\nYou cannot restart.");
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
    <title>Test Results - iMatiz</title>
    <style>
        body {font-family:Arial,sans-serif;background:#f8f9fa;margin:40px;}
        .container {max-width:1100px;margin:auto;background:white;padding:30px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.15);}
        h1 {text-align:center;color:#2c3e50;}
        table {width:100%;border-collapse:collapse;margin-top:25px;}
        th, td {padding:14px;text-align:left;border-bottom:1px solid #ddd;}
        th {background:#007bff;color:white;}
        tr:nth-child(even) {background:#f8f9fa;}
        .score {font-weight:bold;color:#28a745;}
        .terminated {color:#dc3545;font-weight:bold;}
        .back {display:inline-block;margin-top:30px;padding:14px 40px;background:#007bff;color:white;text-decoration:none;border-radius:8px;}
        .back:hover {background:#0069d9;}
    </style>
</head>
<body>
<div class="container">
    <h1>Test Results</h1>
    {% if results|length == 0 %}
        <p style="text-align:center;color:#666;">No results recorded yet.</p>
    {% else %}
    <table>
        <tr>
            <th>Employee Name</th>
            <th>Score</th>
            <th>Answered</th>
            <th>Unanswered</th>
            <th>Total</th>
            <th>Date & Time</th>
            <th>Status</th>
        </tr>
        {% for r in results %}
        <tr>
            <td>{{ r['Employee Name'] }}</td>
            <td class="score">{{ r['Correct Answers'] }}</td>
            <td>{{ r['Answered Questions'] }}</td>
            <td>{{ r['Unanswered Questions'] }}</td>
            <td>{{ r['Total Questions'] }}</td>
            <td>{{ r['Date & Time'] }}</td>
            <td {% if 'Terminated' in r['Status'] %}class="terminated"{% endif %}>
                {{ r['Status'] }}
            </td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}
    <center><a href="/" class="back">Back to Login</a></center>
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
        kicked_msg = "Previous session was terminated due to tab switch or timeout.<br>Contact Santhosh."

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if name not in ALLOWED_EMPLOYEES:
            return "<h2 style='color:red;text-align:center'>Invalid employee name</h2>", 403

        # Prevent re-login in same session (active test)
        if 'name' in session and session['name'] == name:
            return redirect('/test')

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

    return render_template_string(LOGIN_HTML,
                                 employees=ALLOWED_EMPLOYEES,
                                 kicked_msg=kicked_msg)


@app.route('/test', methods=['GET', 'POST'])
def test():
    if 'questions' not in session:
        return redirect('/')

    if session['current'] >= len(session['questions']):
        return redirect('/result')

    if request.method == 'POST':
        ans = request.form.get('ans')
        if ans:
            session['answers'][str(session['current'])] = ans
        session['current'] += 1
        return redirect('/test')

    q = session['questions'][session['current']]
    return render_template_string(QUESTION_HTML,
                                 qnum=session['current'] + 1,
                                 total=len(session['questions']),
                                 question=q['question'],
                                 options=q['options'],
                                 name=session.get('name', ''))   # ← required for lock


@app.route('/result')
def result():
    if 'questions' not in session:
        return redirect('/')

    # Save result
    questions = session['questions']
    answers   = session['answers']
    answered_count = len(answers)
    score = sum(1 for i in range(len(questions)) if str(i) in answers and answers[str(i)] == questions[i]['correct'])

    name  = session.get('name', 'Unknown')
    total = len(questions)
    unanswered = total - answered_count

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

    # Show all results
    if os.path.exists(RESULT_FILE):
        df = pd.read_excel(RESULT_FILE)
        df['Date & Time'] = pd.to_datetime(df['Date & Time'])
        df = df.sort_values('Date & Time', ascending=False)
        results = df.to_dict('records')
    else:
        results = []

    return render_template_string(RESULT_HTML, results=results)


@app.route('/tab_cheat_end')
def tab_cheat_end():
    if 'questions' not in session:
        return redirect('/')

    questions = session['questions']
    answers   = session['answers']
    answered_count = len(answers)
    score = sum(1 for i in range(len(questions)) if str(i) in answers and answers[str(i)] == questions[i]['correct'])

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
    print("   iMatiz Technology - Employee Quiz")
    print("   Allowed users:", ", ".join(ALLOWED_EMPLOYEES))
    print("   Note: Terminated users cannot restart in same browser")
    print("\nOpen →  http://127.0.0.1:5000")
    print("Network → http://<your-ip>:5000   (ipconfig → IPv4 Address)")
    print()
    app.run(host='0.0.0.0', port=5000, debug=True)
