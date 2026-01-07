from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
from datetime import datetime
import os
from openpyxl import Workbook

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "database.db")

# =========================
# DATABASE
# =========================
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS materiais (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                material TEXT NOT NULL,
                cor TEXT NOT NULL,
                data TEXT NOT NULL
            )
        """)

        db.execute("""
            CREATE TABLE IF NOT EXISTS producao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido TEXT NOT NULL,
                modelo TEXT NOT NULL,
                cor TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                etapa TEXT NOT NULL,
                data TEXT NOT NULL
            )
        """)


# inicializa o banco ao carregar o app (Flask 3 / Gunicorn safe)
init_db()

# =========================
# ROTAS PRINCIPAIS
# =========================
@app.route("/")
def index():
    return render_template("index.html")


# =========================
# MATERIAL (BOBINA)
# =========================
@app.route("/material", methods=["GET", "POST"])
def material():
    db = get_db()

    if request.method == "POST":
        material = request.form["material"]
        cor = request.form["cor"]

        db.execute(
            "INSERT INTO materiais (material, cor, data) VALUES (?, ?, ?)",
            (material, cor, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        db.commit()
        return redirect(url_for("material"))

    materiais = db.execute(
        "SELECT * FROM materiais ORDER BY id DESC"
    ).fetchall()

    return render_template("materials.html", materiais=materiais)


# =========================
# CORTE
# =========================
@app.route("/producao/corte", methods=["GET", "POST"])
def corte():
    db = get_db()

    if request.method == "POST":
        db.execute(
            """INSERT INTO producao
               (pedido, modelo, cor, quantidade, etapa, data)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                request.form["pedido"],
                request.form["modelo"],
                request.form["cor"],
                request.form["quantidade"],
                "CORTE",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
        )
        db.commit()
        return redirect(url_for("corte"))

    registros = db.execute(
        "SELECT * FROM producao WHERE etapa = 'CORTE' ORDER BY id DESC"
    ).fetchall()

    return render_template("cut.html", registros=registros)


# =========================
# COSTURA - ENTRADA
# =========================
@app.route("/producao/costura-entrada", methods=["GET", "POST"])
def costura_entrada():
    db = get_db()

    if request.method == "POST":
        db.execute(
            """INSERT INTO producao
               (pedido, modelo, cor, quantidade, etapa, data)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                request.form["pedido"],
                request.form["modelo"],
                request.form["cor"],
                request.form["quantidade"],
                "COSTURA_ENTRADA",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
        )
        db.commit()
        return redirect(url_for("costura_entrada"))

    registros = db.execute(
        "SELECT * FROM producao WHERE etapa = 'COSTURA_ENTRADA' ORDER BY id DESC"
    ).fetchall()

    return render_template("sewing_in.html", registros=registros)


# =========================
# COSTURA - SAÍDA (ITENS PRONTOS)
# =========================
@app.route("/producao/costura-saida", methods=["GET", "POST"])
def costura_saida():
    db = get_db()

    if request.method == "POST":
        db.execute(
            """INSERT INTO producao
               (pedido, modelo, cor, quantidade, etapa, data)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                request.form["pedido"],
                request.form["modelo"],
                request.form["cor"],
                request.form["quantidade"],
                "COSTURA_SAIDA",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
        )
        db.commit()
        return redirect(url_for("costura_saida"))

    registros = db.execute(
        "SELECT * FROM producao WHERE etapa = 'COSTURA_SAIDA' ORDER BY id DESC"
    ).fetchall()

    return render_template("sewing_out.html", registros=registros)


# =========================
# RELATÓRIOS
# =========================
@app.route("/relatorio")
def relatorio():
    db = get_db()
    dados = db.execute(
        "SELECT * FROM producao ORDER BY data DESC"
    ).fetchall()

    return render_template("reports.html", dados=dados)


# =========================
# EXPORTAR EXCEL
# =========================
@app.route("/exportar")
def exportar():
    db = get_db()
    dados = db.execute("SELECT * FROM producao ORDER BY data").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.append(["Pedido", "Modelo", "Cor", "Quantidade", "Etapa", "Data"])

    for d in dados:
        ws.append([
            d["pedido"],
            d["modelo"],
            d["cor"],
            d["quantidade"],
            d["etapa"],
            d["data"]
        ])

    file_path = os.path.join(BASE_DIR, "relatorio_producao.xlsx")
    wb.save(file_path)

    return send_file(file_path, as_attachment=True)


# =========================
# START LOCAL (opcional)
# =========================
if __name__ == "__main__":
    app.run(debug=True)
