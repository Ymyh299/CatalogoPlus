from flask import Flask, render_template, redirect, url_for, session, flash, request, send_file, jsonify
import json
from functools import wraps
from werkzeug.security import check_password_hash
import mysql.connector
import subprocess
import os 
import time
import pandas as pd


app = Flask(__name__)
app.secret_key = "anag25000"

DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "QualquerUma123",   
    "database": "catalogoplus"
}


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "indesign")
CSV_PATH = os.path.join(DATA_DIR, "data_merge.csv")
SCRIPT_PATH = os.path.join(DATA_DIR, "merge_script.jsx")
PDF_PATH = os.path.join(DATA_DIR, "output", "resultado.pdf")

INDESIGN_EXE = r"C:\Program Files\Adobe\Adobe InDesign 2026\InDesign.exe"

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("username")
        senha = request.form.get("password")

        # Conectar ao banco
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)

        cursor.execute("SELECT * FROM usuarios WHERE user = %s", (usuario,))
        user = cursor.fetchone()

        cursor.close()
        conn.close()

        if not user:
            flash("Usuário ou senha incorretos.", "erro")
            return redirect(url_for("login"))

        # Verificar senha usando bcrypt (Werkzeug)
        if not check_password_hash(user["password_hash"], senha):
            flash("Usuário ou senha incorretos.", "erro")
            return redirect(url_for("login"))

        # Login OK
        session["user_id"] = user["id"]
        session["usuario"] = user["user"]

        flash("Login realizado com sucesso!", "sucesso")
        return redirect(url_for("index"))

    return render_template("login.html")


@app.route("/")
@login_required
def index():
    return render_template("index.html", usuario=session["usuario"])

# @app.route("/painel")
# @login_required
# def painel():
#     return render_template("painel.html", usuario=session["usuario"])




@app.route("/visualizar")
@login_required
def visualizar():

    return render_template("visualizer.html", usuario=session["usuario"])




@app.route('/foto/<ref>')
def foto(ref):
    caminho = f"C:/Users/Administrador/Documents/fotosref/{ref}.jpg"
    return send_file(caminho)




@app.route('/painel', methods=["GET", "POST"])
@login_required
def painel():
    layout_escolhido = request.form.get('layout_escolhido')
    print(f"Layout escolhido: {layout_escolhido}")
    session['layout_escolhido'] = layout_escolhido
    return render_template("painel.html", usuario=session["usuario"])



@app.route("/opcoes", methods=["POST"])
def opcoes():
    session['referencias'] = request.form.get("referencias")
    return render_template("option.html")




def executar_indesign():
    vbs_path = r"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\indesign\rodar_indesign.vbs"
    try:
        # Usar subprocess.run() que espera o processo terminar
        resultado = subprocess.run(["wscript", vbs_path], shell=False, capture_output=True, text=True)
        
        if resultado.returncode == 0:
            print("InDesign executado com sucesso")
            return True
        else:
            print("Erro na execução do VBS:", resultado.stderr)
            return False
            
    except Exception as e:
        print("Erro ao executar vbs:", e)
        return False
    

@app.route("/gerar_planilha", methods=["POST"])
def gerar_planilha():
    dados_json = request.form.get("dados_json")
    dados = json.loads(dados_json)
    layout = session.get("layout_escolhido")
    referencias = session.get("referencias")

    # monta planilha
    df = pd.DataFrame([{
        "layout": layout,
        "referencias": referencias,
        **dados
    }])

    # salva o XLSX
    df.to_excel("resultado.xlsx", index=False)

    # (Opcional) gerar PDF aqui também

    return redirect(url_for("visualizar"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)