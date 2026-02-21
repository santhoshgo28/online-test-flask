from flask import Flask, render_template_string, request, session, redirect, url_for
import pandas as pd
import random
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'super-secret-key-change-this-2025'

# ────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "questions.xlsx")
RESULT_FILE = os.path.join(BASE_DIR, "result.xlsx")

ALLOWED_EMPLOYEES = [
    "Santhosh", "Rajkumar Sir ", "Ram Sir ", "janani G Hegde2", "Amrutha N M",
    "AishwaryA G N", "Satish", "Zaiba Khanum", "GuruDivya L", "Aarthi R",
    "Vashanth Kumar", "Abinaya 1", "Suchithra PS", "Dhanapriya R",
    "Dhanya Shree U", "Nivetha S", "Shreyas CM", "Siri H G", "Ananaya GC",
    "Ashwini Sindhe", "Gopika ", "Bhagya Shree U", "Sri Dharshini",
    "Kavikeerthana Palmani", "Ramya shree", "PriyaDharshini", "Keerthana L",
    "NAGARAJAN R", "Kiruthika Saravanan", "Vimalkarthik"
]

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
# HTML TEMPLATES
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
    <h1>iMatiz Technology Quiz</h1>
    {% if kicked_msg %}<div class="msg">{{ kicked_msg | safe }}</div>{% endif %}
   
    <div id="blocked-msg" class="blocked">
        This test was terminated earlier (tab switch / timeout).<br>
        You are no longer allowed to restart in this browser.<br>
        Contact Rajkumar if needed.
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
        .buttons {display:flex; justify-content:center; gap:30px; margin-top:2.5rem;}
        button {padding:0.9rem 2.5rem;font-size:1.25rem;border:none;border-radius:8px;cursor:pointer;}
        .next {background:#007bff;color:white;}
        .next:hover {background:#0069d9;}
        .skip {background:#6c757d;color:white;}
        .skip:hover {background:#5a6268;}
    </style>
</head>
<body onload="startTimer();">
<div class="container">
    <div class="timer" id="timer">30 seconds</div>
    <h2>Question {{ qnum }} of {{ total }}</h2>
    <div class="question">{{ question }}</div>
   
    <form method="post" id="form">
        {% for opt in options %}
        <label>
            <input type="radio" name="ans" value="{{ 'ABCD'[loop.index0] }}">
            {{ opt }}
        </label>
        {% endfor %}
       
        <div class="buttons">
            <button type="submit" name="action" value="next" class="next">Next →</button>
            <button type="button" onclick="window.location.href='/test?skip=1'" class="skip">Skip</button>
        </div>
    </form>
</div>
<script>
let time = 30;
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

FINISHED_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Quiz Completed - iMatiz</title>
    <style>
        body {font-family:Arial,sans-serif;background:#f8f9fa;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}
        .card {background:white;padding:60px 40px;border-radius:16px;box-shadow:0 10px 40px rgba(0,0,0,0.2);text-align:center;max-width:520px;}
        h1 {color:#28a745;font-size:2.8rem;margin-bottom:20px;}
        p {font-size:1.4rem;color:#555;line-height:1.6;margin:30px 0;}
        .btn {display:inline-block;padding:18px 50px;background:#007bff;color:white;font-size:1.4rem;text-decoration:none;border-radius:10px;transition:0.2s;}
        .btn:hover {background:#0062cc; transform: translateY(-2px);}
    </style>
</head>
<body>
<div class="card">
    <h1>Quiz Completed!</h1>
    <p>Thank you, <strong>{{ name }}</strong>!<br>You have successfully finished the iMatiz Technology Quiz.</p>
    <a href="/result" class="btn">View My Results →</a>
</div>
</body>
</html>
"""

RESULT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Your Results - iMatiz</title>
    <style>
        body {font-family:Arial,sans-serif;background:#f8f9fa;margin:40px;}
        .container {max-width:1000px;margin:auto;background:white;padding:30px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.15);}
        h1, h2 {text-align:center;color:#2c3e50;}
        .greeting {text-align:center; color:#2c3e50; margin-bottom:30px;}
        table {width:100%;border-collapse:collapse;margin:25px 0;}
        th, td {padding:14px;text-align:left;border-bottom:1px solid #ddd;}
        th {background:#007bff;color:white;}
        tr:nth-child(even) {background:#f8f9fa;}
        .score {font-weight:bold;color:#28a745;font-size:1.3em;}
        .terminated {color:#dc3545;font-weight:bold;}
        .meta {color:#555; margin:10px 0; text-align:center;}
        .back {display:inline-block;padding:14px 40px;background:#007bff;color:white;text-decoration:none;border-radius:8px;}
        .back:hover {background:#0069d9;}
    </style>
</head>
<body>
<div class="container">
    <h1>Your Quiz Results</h1>
    <div class="greeting">Hello <strong>{{ employee_name }}</strong></div>
    {% if results|length == 0 %}
        <p style="text-align:center;color:#666;">No previous quiz attempts found.</p>
    {% else %}
    <table>
        <tr>
            <th>Date & Time</th>
            <th>Score</th>
            <th>Answered</th>
            <th>Skipped</th>
            <th>Total</th>
            <th>Status</th>
        </tr>
        {% for r in results %}
        <tr>
            <td>{{ r['Date & Time'] }}</td>
            <td class="score">{{ r['Correct Answers'] }} / {{ r['Total Questions'] }}</td>
            <td>{{ r['Answered Questions'] }}</td>
            <td>{{ r['Skipped Questions'] }}</td>
            <td>{{ r['Total Questions'] }}</td>
            <td {% if 'Terminated' in r['Status'] %}class="terminated"{% endif %}>
                {{ r['Status'] }}
            </td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}
    <center>
        <a href="/" class="back">Back to Login</a>
    </center>
</div>
</body>
</html>
"""

# ────────────────────────────────────────────────
# ROUTES
# ────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    kicked_msg = ""
    if request.args.get('terminated') == 'yes':
        kicked_msg = "Previous session was terminated due to tab switch or timeout.<br>Contact Rajkumar ."
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if name not in ALLOWED_EMPLOYEES:
            return "<h2 style='color:red;text-align:center'>Invalid employee name</h2>", 403
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
        session['name'] = name
        session['questions'] = questions
        session['current'] = 0
        session['answers'] = {}  # None = skipped / timeout, str = answered
        return redirect('/test')
    return render_template_string(LOGIN_HTML,
                                 employees=ALLOWED_EMPLOYEES,
                                 kicked_msg=kicked_msg)

@app.route('/test', methods=['GET', 'POST'])
def test():
    if 'questions' not in session:
        return redirect('/')

    if session['current'] >= len(session['questions']):
        return redirect('/finished')   # ← Changed: go to thank-you screen instead of result

    # Handle skip via GET ?skip=1
    if request.method == 'GET' and request.args.get('skip') == '1':
        session['answers'][str(session['current'])] = None
        session['current'] += 1
        return redirect('/test')

    if request.method == 'POST':
        action = request.form.get('action', 'next')
        if action == 'next':
            ans = request.form.get('ans')
            session['answers'][str(session['current'])] = ans if ans else None
        session['current'] += 1
        return redirect('/test')

    q = session['questions'][session['current']]
    return render_template_string(QUESTION_HTML,
                                 qnum=session['current'] + 1,
                                 total=len(session['questions']),
                                 question=q['question'],
                                 options=q['options'],
                                 name=session.get('name', ''))

@app.route('/finished')
def finished():
    if 'questions' not in session:
        return redirect('/')
    return render_template_string(FINISHED_HTML,
                                 name=session.get('name', 'Participant'))

@app.route('/result')
def result():
    if 'questions' not in session:
        return redirect('/')

    # Calculate & save BEFORE clearing session
    name = session.get('name', 'Unknown')
    questions = session.get('questions', [])
    answers = session.get('answers', {})
    total = len(questions)
    correct = 0
    answered = 0
    skipped = 0

    for i in range(total):
        user_ans = answers.get(str(i))
        if user_ans is None:
            skipped += 1
        else:
            answered += 1
            if user_ans == questions[i]['correct']:
                correct += 1

    row = {
        'Employee Name': name,
        'Correct Answers': correct,
        'Answered Questions': answered,
        'Skipped Questions': skipped,
        'Total Questions': total,
        'Date & Time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'Status': 'Completed'
    }

    # Save result
    if os.path.exists(RESULT_FILE):
        df = pd.read_excel(RESULT_FILE)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        df = pd.DataFrame([row])
    df.to_excel(RESULT_FILE, index=False)

    # Load user's previous results
    if os.path.exists(RESULT_FILE):
        df = pd.read_excel(RESULT_FILE)
        df['Employee Name'] = df['Employee Name'].astype(str).str.strip()
        user_results = df[df['Employee Name'] == name].copy()
        user_results['Date & Time'] = pd.to_datetime(user_results['Date & Time'])
        user_results = user_results.sort_values('Date & Time', ascending=False)
        results = user_results.to_dict('records')
    else:
        results = []

    # Safe to clear now
    session.clear()

    return render_template_string(RESULT_HTML,
                                 results=results,
                                 employee_name=name)

@app.route('/tab_cheat_end')
def tab_cheat_end():
    if 'questions' not in session:
        return redirect('/')

    name = session.get('name', 'Unknown')
    questions = session.get('questions', [])
    answers = session.get('answers', {})
    total = len(questions)
    correct = 0
    answered = 0
    skipped = 0

    for i in range(total):
        user_ans = answers.get(str(i))
        if user_ans is None:
            skipped += 1
        else:
            answered += 1
            if user_ans == questions[i]['correct']:
                correct += 1

    row = {
        'Employee Name': f"{name} (Terminated - Tab Switch)",
        'Correct Answers': correct,
        'Answered Questions': answered,
        'Skipped Questions': skipped,
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
    print(" iMatiz Technology - Employee Quiz")
    print(" Allowed users:", ", ".join(ALLOWED_EMPLOYEES))
    print(" Now shows 'Quiz Completed' screen before showing results")
    print(" Results are private per employee")
    print("\nOpen → http://127.0.0.1:5000")
    print("Network → http://<your-ip>:5000")
    print()
    app.run(host='0.0.0.0', port=5000, debug=True)
