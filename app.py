from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
from datetime import datetime
import os
from openpyxl import Workbook

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "database.db")


# ---------------- DATABASE ----------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS materiais (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                material TEXT,
                cor TEXT,
                data TEXT
            )
        """)

        db.execute("""
            CREATE TABLE IF NOT EXISTS producao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido TEXT,
                modelo TEXT,
                cor TEXT,
                quantidade INTEGER,
                etapa TEXT,
                data TEXT
            )
        """)


@app.before_first_request
def setup():
    init_db()


# ---------------- ROTAS ----------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/material", methods=["GET", "POST"])
def material():
    db = get_db()

    if request.method == "POST":
        material = request.form["material"]
        cor = request.form["cor"]

        db.execute(
            "INSERT INTO materiais (material, cor, data) VALUES (?, ?, ?)",
            (material, cor, datetime.now())
        )
        db.commit()
        return redirect(url_for("material"))

    materiais = db.execute(
        "SELECT * FROM materiais ORDER BY id DESC"
    ).fetchall()

    return render_template("materials.html", materiais=materiais)


@app.route("/producao/<etapa>", methods=["GET", "POST"])
def producao(etapa):
    db = get_db()

    if request.method == "POST":
        pedido = request.form["pedido"]
        modelo = request.form["modelo"]
        cor = request.form["cor"]
        quantidade = request.form["quantidade"]

        db.execute(
            "INSERT INTO producao (pedido, modelo, cor, quantidade, etapa, data) VALUES (?, ?, ?, ?, ?, ?)",
            (pedido, modelo, cor, quantidade, etapa, datetime.now())
        )
        db.commit()
        return redirect(url_for("producao", etapa=etapa))

    registros = db.execute(
        "SELECT * FROM producao WHERE etapa = ? ORDER BY id DESC",
        (etapa,)
    ).fetchall()

    return render_template("producao.html", registros=registros, etapa=etapa)


@app.route("/relatorio")
def relatorio():
    db = get_db()
    dados = db.execute(
        "SELECT * FROM producao ORDER BY data DESC"
    ).fetchall()
    return render_template("reports.html", dados=dados)


@app.route("/exportar")
def exportar():
    db = get_db()
    dados = db.execute("SELECT * FROM producao").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.append(["Pedido", "Modelo", "Cor", "Quantidade", "Etapa", "Data"])

    for d in dados:
        ws.append([d["pedido"], d["modelo"], d["cor"], d["quantidade"], d["etapa"], d["data"]])

    file_path = os.path.join(BASE_DIR, "relatorio.xlsx")
    wb.save(file_path)

    return send_file(file_path, as_attachment=True)


# ---------------- START ----------------
if __name__ == "__main__":
    init_db()
    app.run()
