from flask import Flask, render_template, request, redirect, url_for, session
import os
from openpyxl import Workbook, load_workbook
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'devsecret')

EXCEL_FILE = 'presences_questions.xlsx'
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

# --- Initialiser le fichier Excel si n'existe pas ---
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()

    # Feuille Présences
    ws1 = wb.active
    ws1.title = "Présences"
    ws1.append(["Timestamp","Nom","Prénom","Sexe","Quartier","Téléphone","Événement"])

    # Feuille Questions
    ws2 = wb.create_sheet(title="Questions")
    ws2.append(["Timestamp","Nom","Question","Événement"])

    wb.save(EXCEL_FILE)

# --- Routes ---
@app.route('/')
def index():
    return redirect(url_for('presence'))

@app.route('/presence', methods=['GET','POST'])
def presence():
    if request.method == 'POST':
        nom = request.form['nom']
        prenom = request.form['prenom']
        sexe = request.form['sexe']
        quartier = request.form['quartier']
        telephone = request.form['telephone']
        evenement = request.form.get('event','General')
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Présences"]
        ws.append([timestamp, nom, prenom, sexe, quartier, telephone, evenement])
        wb.save(EXCEL_FILE)
        return render_template('merci.html', message="Présence enregistrée !")
    return render_template('presence.html')

@app.route('/questions', methods=['GET','POST'])
def questions():
    if request.method == 'POST':
        nom = request.form['nom']
        question = request.form['question']
        evenement = request.form.get('event','General')
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Questions"]
        ws.append([timestamp, nom, question, evenement])
        wb.save(EXCEL_FILE)
        return render_template('merci.html', message="Question envoyée !")
    return render_template('questions.html')

@app.route('/admin', methods=['GET','POST'])
def admin():
    if request.method == 'POST':
        password = request.form['password']
        if password == ADMIN_PASSWORD:
            session['admin'] = True
            return redirect(url_for('admin_dashboard'))
        else:
            return render_template('admin.html', error="Mot de passe incorrect")
    return render_template('admin.html')

@app.route('/admin/dashboard')
def admin_dashboard():
    if not session.get('admin'):
        return redirect(url_for('admin'))

    wb = load_workbook(EXCEL_FILE)

    # Récupérer les présences
    presences = []
    for row in wb["Présences"].iter_rows(values_only=True):
        if row[0] == "Timestamp":
            continue
        presences.append({
            "Timestamp": row[0],
            "Nom": row[1],
            "Prénom": row[2],
            "Sexe": row[3],
            "Quartier": row[4],
            "Téléphone": row[5],
            "Événement": row[6]
        })

    # Récupérer les questions
    questions = []
    for row in wb["Questions"].iter_rows(values_only=True):
        if row[0] == "Timestamp":
            continue
        questions.append({
            "Timestamp": row[0],
            "Nom": row[1],
            "Question": row[2],
            "Événement": row[3]
        })

    return render_template('admin_dashboard.html', presences=presences, questions=questions)

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin', None)
    return redirect(url_for('admin'))

if __name__ == "__main__":
    app.run(debug=True)
