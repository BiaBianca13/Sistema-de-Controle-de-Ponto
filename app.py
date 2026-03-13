import os
import sqlite3
import openpyxl
import re
import psycopg2.extras

from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = "secret"

@app.template_filter("tempo")
def tempo(minutos):

    minutos = int(minutos)

    horas = minutos // 60
    resto = minutos % 60

    if horas == 0:
        return f"{resto} min"
    if resto == 0:
        return f"{horas}h"

    return f"{horas}h {resto} min"

DB = "ponto.db"

import os
import psycopg2

DATABASE_URL = os.environ.get("DATABASE_URL")

if DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://")

def p():
    if DATABASE_URL:
        return "%s"
    return "?"

def conectar():
    if DATABASE_URL:
        conn = psycopg2.connect(DATABASE_URL)
        conn.cursor_factory = psycopg2.extras.RealDictCursor
        return conn
    else:
        conn = sqlite3.connect(DB)
        conn.row_factory = sqlite3.Row
        return conn
    
def criar_banco():
    conn = conectar()
    cur = conn.cursor()

    if DATABASE_URL:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS funcionarios(
            id SERIAL PRIMARY KEY,
            nome TEXT UNIQUE,
            entrada TEXT,
            saida TEXT,
            almoco INTEGER,
            valor_hora REAL
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS registros(
            id SERIAL PRIMARY KEY,
            funcionario TEXT,
            data TEXT,
            batidas TEXT,
            minutos_trabalhados INTEGER,
            minutos_extra INTEGER,
            atraso INTEGER,
            debito INTEGER,
            falta INTEGER
        )
        """)
    else:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS funcionarios(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT UNIQUE,
            entrada TEXT,
            saida TEXT,
            almoco INTEGER,
            valor_hora REAL
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS registros(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            funcionario TEXT,
            data TEXT,
            batidas TEXT,
            minutos_trabalhados INTEGER,
            minutos_extra INTEGER,
            atraso INTEGER,
            debito INTEGER,
            falta INTEGER
        )
        """)

    conn.commit()
    conn.close()


criar_banco()


def extrair_horas(texto):
    if not texto:
        return []
    linhas = str(texto).split("\n")
    horas = []

    for l in linhas:
        l = l.strip()
        if re.match(r"\d{2}:\d{2}", l):
            horas.append(l)
    return horas

def calcular(batidas):
    if not batidas or len(batidas) < 1:
        return 0, None, "falta"

    horarios = [datetime.strptime(h, "%H:%M") for h in batidas]
    horarios.sort()

    filtrado = [horarios[0]]

    for h in horarios[1:]:
        diff = (h - filtrado[-1]).total_seconds() / 60
        if diff > 2:
            filtrado.append(h)

    entrada_real = filtrado[0]
    saida_real = filtrado[-1]

    intervalo = timedelta()

    meio = filtrado[1:-1]

    for i in range(0, len(meio) - 1, 2):
        inicio = meio[i]
        fim = meio[i + 1]

        if fim > inicio:
            intervalo += (fim - inicio)

    trabalhado = (saida_real - entrada_real) - intervalo

    if trabalhado.total_seconds() < 0:
        minutos = 0
    else:
        minutos = int(trabalhado.total_seconds() / 60)

    status = "ok"

    if len(filtrado) == 3:
        status = "saida_faltando"

    elif len(filtrado) == 2:
        status = "sem_volta_almoco"

    return minutos, entrada_real, status

def ler_excel(caminho):
    wb = openpyxl.load_workbook(caminho)
    ws = wb.active

    dados = []

    periodo = None

    for col in range(1, 10):
        valor = ws.cell(2, col).value
        if valor and "/" in str(valor):
            periodo = valor
            break

    if periodo:
        datas = re.findall(r"\d{2}/\d{2}/\d{4}", str(periodo))
        dia, mes, ano = datas[0].split("/")
        mes = int(mes)
        ano = int(ano)
    else:
        hoje = date.today()
        mes = hoje.month
        ano = hoje.year

    for r in range(1, ws.max_row):

        if ws.cell(r, 1).value == "ID":

            nome = ws.cell(r, 12).value

            linha_dias = r + 1
            linha_semana = r + 2
            linha_horas = r + 3

            for col in range(1, 32):

                dia = ws.cell(linha_dias, col).value
                semana = ws.cell(linha_semana, col).value

                if isinstance(dia, int):

                    if semana in ["SAB", "DOM"]:
                        continue

                    batidas = extrair_horas(ws.cell(linha_horas, col).value)

                    dados.append({
                        "nome": nome,
                        "data": date(ano, mes, dia),
                        "batidas": batidas
                    })
    return dados

@app.route("/")
def dashboard():
    conn = conectar()
    cur = conn.cursor()

    cur.execute("SELECT * FROM registros ORDER BY funcionario, data")
    dados = cur.fetchall()

    conn.close()

    funcionarios = {}
    resumo = {}

    for d in dados:
        nome = d["funcionario"]

        if nome not in funcionarios:
            funcionarios[nome] = []
            resumo[nome] = {
                "extra": 0,
                "debito": 0,
                "atraso": 0,
                "falta": 0,
                "saldo": 0
            }

        funcionarios[nome].append(d)
        resumo[nome]["extra"] += d["minutos_extra"]
        resumo[nome]["debito"] += d["debito"]
        resumo[nome]["atraso"] += d["atraso"]
        resumo[nome]["falta"] += d["falta"]

    for nome in resumo:
        resumo[nome]["saldo"] = resumo[nome]["extra"] - resumo[nome]["atraso"]

    return render_template(
        "dashboard.html",
        funcionarios=funcionarios,
        resumo=resumo
    )

@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        arquivo = request.files.get("arquivo")

        if not arquivo or arquivo.filename == "":
            flash("Selecione um arquivo.")
            return redirect("/upload")

        os.makedirs("uploads", exist_ok=True)

        caminho = os.path.join("uploads", arquivo.filename)
        arquivo.save(caminho)

        registros = ler_excel(caminho)

        conn = conectar()
        cur = conn.cursor()

        cur.execute("DELETE FROM registros")

        for r in registros:

            batidas = r["batidas"]

            extra = 0
            debito = 0
            atraso = 0
            falta = 0
            minutos = 0

            cur.execute(
                f"SELECT * FROM funcionarios WHERE UPPER(nome)=UPPER({p()})",
                (r["nome"],)
            )

            func = cur.fetchone()

            if len(batidas) < 2:
                falta = 1

            else:

                minutos, entrada_real, status = calcular(batidas)

                if func:

                    entrada_cfg = datetime.strptime(func["entrada"], "%H:%M")
                    saida_cfg = datetime.strptime(func["saida"], "%H:%M")

                    if r["data"].weekday() == 4:
                        saida_cfg = datetime.strptime("17:00", "%H:%M")

                    almoco_min = int(func["almoco"])

                    horarios = [datetime.strptime(h, "%H:%M") for h in batidas]
                    horarios.sort()

                    filtrado = [horarios[0]]

                    for h in horarios[1:]:
                        diff = (h - filtrado[-1]).total_seconds() / 60
                        if diff > 10:
                            filtrado.append(h)

                    horarios = filtrado

                    entrada_real = horarios[0]
                    saida_real = horarios[-1]

                    extra = 0
                    atraso = 0

                    if entrada_real < entrada_cfg:
                        extra += int((entrada_cfg - entrada_real).total_seconds() / 60)

                    if entrada_real > entrada_cfg:
                        atraso += int((entrada_real - entrada_cfg).total_seconds() / 60)

                    if status != "saida_faltando":

                        if saida_real > saida_cfg:
                            extra += int((saida_real - saida_cfg).total_seconds() / 60)

                        if saida_real < saida_cfg:
                            atraso += int((saida_cfg - saida_real).total_seconds() / 60)

                        if len(horarios) >= 4:

                            almoco_saida = horarios[1]
                            almoco_volta = horarios[2]

                            intervalo = int(
                                (almoco_volta - almoco_saida).total_seconds() / 60
                            )

                            if intervalo > almoco_min:
                                atraso += intervalo - almoco_min

            cur.execute(f"""
                INSERT INTO registros(
                    funcionario,
                    data,
                    batidas,
                    minutos_trabalhados,
                    minutos_extra,
                    atraso,
                    debito,
                    falta
                ) VALUES ({p()},{p()},{p()},{p()},{p()},{p()},{p()},{p()})
            """, (
                r["nome"],
                str(r["data"]),
                "\n".join(batidas),
                minutos,
                extra,
                atraso,
                debito,
                falta
            ))

        conn.commit()
        conn.close()

        flash("Arquivo importado com sucesso.")
        return redirect("/")

    return render_template("upload.html")

@app.route("/funcionarios")
def funcionarios():

    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT * FROM funcionarios")
    lista = cur.fetchall()
    conn.close()
    return render_template("funcionarios.html", lista=lista)

@app.route("/funcionarios/add", methods=["POST"])
def add_func():
    nome = request.form["nome"]
    entrada = request.form["entrada"]
    saida = request.form["saida"]
    almoco = request.form["almoco"]

    conn = conectar()   
    cur = conn.cursor()

    try:
        if DATABASE_URL:
            cur.execute(f"""
            INSERT INTO funcionarios
            (nome,entrada,saida,almoco,valor_hora)
            VALUES({p()},{p()},{p()},{p()},{p()})
            ON CONFLICT (nome) DO NOTHING
            """,(nome,entrada,saida,almoco,0))
        else:
            cur.execute(f"""
            INSERT OR IGNORE INTO funcionarios
            (nome,entrada,saida,almoco,valor_hora)
            VALUES({p()},{p()},{p()},{p()},{p()})
            """,(nome,entrada,saida,almoco,0))
    except Exception as e:
        print(e)


    conn.commit()
    conn.close()

    return redirect("/funcionarios")

@app.route("/funcionarios/delete/<int:id>")
def deletar_funcionario(id):

    conn = conectar()
    cur = conn.cursor()

    cur.execute(f"DELETE FROM funcionarios WHERE id={p()}", (id,))

    conn.commit()
    conn.close()

    return redirect("/funcionarios")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
