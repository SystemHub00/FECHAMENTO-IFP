import csv
import os
import uuid
from flask import Flask, render_template, redirect, url_for

app = Flask(__name__)

PROTOCOLO_CSV = 'protocolos.csv'
PROTOCOLO_SET = set()

# Carrega protocolos já existentes
if os.path.exists(PROTOCOLO_CSV):
    with open(PROTOCOLO_CSV, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if row:
                PROTOCOLO_SET.add(row[0])

def gerar_protocolo():
    while True:
        protocolo = str(uuid.uuid4())
        if protocolo not in PROTOCOLO_SET:
            PROTOCOLO_SET.add(protocolo)
            with open(PROTOCOLO_CSV, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([protocolo])
            return protocolo

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/inscricao')
def inscricao():
    return render_template('inscricao.html')

@app.route('/endereco')
def endereco():
    return render_template('endereco.html')

@app.route('/curso')
def curso():
    return render_template('curso.html')

@app.route('/revisao')
def revisao():
    return render_template('revisao.html')

@app.route('/confirmacao')
def confirmacao():
    protocolo = gerar_protocolo()
    return render_template('confirmacao.html', protocolo=protocolo)

if __name__ == '__main__':
    app.run(debug=True)
