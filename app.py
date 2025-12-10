from flask import Flask, render_template, redirect, url_for, session, flash, request, send_file, jsonify
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




@app.route('/painel')
@login_required
def painel():
    layout_escolhido = request.form.get('layout_escolhido')
    # layout_escolhido será "img1.png", "layout-moderno", etc.
    print(f"Layout escolhido: {layout_escolhido}")
    
    # Salvar no banco de dados ou sessão
    session['layout_escolhido'] = layout_escolhido
    
    return render_template("painel.html", usuario=session["usuario"])



@app.route("/opcoes")
@login_required
def options():
    
    return render_template("option.html", usuario=session["usuario"])



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
    

@app.route("/criarplanilha", methods=["POST"])
def criar_planilha():
    referencias_texto = request.form.get("referencia")

    if not referencias_texto:
        return jsonify({"erro": "Nenhuma referência informada."}), 400

    referencias = referencias_texto.strip().split()

    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor()

    query = """
        SELECT name, price, promotional, price_promotional
        FROM products
        WHERE code = %s
    """

    dados = []

    for ref in referencias:
        cursor.execute(query, (ref,))
        result = cursor.fetchone()

        if result:
            name, price, promotional, price_promotional = result
            
            if promotional == 1:
                real = "R$"
                traco = "\\"
                de_ou_por = "DE:"
                por = "POR:"
            else:
                real = ""
                traco = ""
                de_ou_por = "POR:"
                por = ""
                price_promotional = ""

            foto_path = rf"Z:\\FOTOS\fotos_PDF\ref\{ref}.jpg"

            dados.append({
                "Referência": ref,
                "Nome": name,
                "De/Por": de_ou_por,
                "Preço Original": price,
                "Traço": traco,
                "Real": real,
                "Preço Promocional": price_promotional,
                "Por": por,
                "@Fotos": foto_path
            })

    cursor.close()
    conn.close()
    

    if not dados:
        return jsonify({"erro": "Nenhuma referência encontrada."}), 404

    # 💾 Gera CSV compatível com InDesign
    df = pd.DataFrame(dados)
    df.to_csv(CSV_PATH, index=False, sep=",", encoding="UTF-16")

    # 🚀 Executa o InDesign automaticamente
    sucesso = executar_indesign()

    if sucesso:
        time.sleep(10)
        return jsonify({
            "mensagem": "✅ Planilha gerada e InDesign executado com sucesso!",
            "quantidade": len(dados),
            "arquivo_csv": CSV_PATH
        })
    else:
        return jsonify({"erro": "Falha ao executar o InDesign."}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)