from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
MODEL_FOLDER = 'models'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)
if not os.path.exists(MODEL_FOLDER):
    os.makedirs(MODEL_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MODEL_FOLDER'] = MODEL_FOLDER

# Cria o arquivo modelo se não existir
def create_model_file():
    model_path = os.path.join(app.config['MODEL_FOLDER'], 'modelo_planilha.xlsx')
    if not os.path.exists(model_path):
        columns = ['ESPECIE', 'CNPJ', 'ACUMULADOR', 'CFOP', 'NF', 'DATA ENTRADA', 'VALOR', 'IRRF', 'CSRF']
        df = pd.DataFrame(columns=columns)
        df.to_excel(model_path, index=False)

create_model_file()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            output_filename = process_file(filepath)
            flash('File processed successfully!')
            return redirect(url_for('download', filename=output_filename))
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename), as_attachment=True)

@app.route('/download_model')
def download_model():
    model_path = os.path.join(app.config['MODEL_FOLDER'], 'modelo_planilha.xlsx')
    return send_file(model_path, as_attachment=True)

def process_file(filepath):
    clientes = pd.read_excel(filepath)
    clientes = clientes.drop_duplicates(['CNPJ','NF','DATA ENTRADA','VALOR'])
    
    clientes['DATA ENTRADA'] = pd.to_datetime(clientes['DATA ENTRADA'], dayfirst=True)
    clientes['DATA ENTRADA'] = clientes['DATA ENTRADA'].dt.strftime('%d/%m/%Y')
    clientes['VALOR'] = clientes['VALOR'].astype("float")
    clientes['IRRF'] = clientes['IRRF'].astype("float")
    clientes['CSRF'] = clientes['CSRF'].astype("float")
    
    data = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f"Notas_com_Retencoes_{data}.txt"
    output_filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
    
    with open(output_filepath, 'w') as arquivo:
        for index, row in clientes.iterrows():
            ESPECIE = row['ESPECIE']
            CNPJ = str(row['CNPJ']).zfill(14)
            ACUMULADOR = row['ACUMULADOR']
            CFOP = row['CFOP']
            NF = row['NF']
            ENTRADA = row['DATA ENTRADA']
            VALOR = '{:.2f}'.format(row['VALOR']).replace('.', ',')
            IRRF = row['IRRF']
            CSRF = row['CSRF']
            VALOR_IRRF = str(IRRF).replace('.', ',')
            VALOR_CSRF = str(CSRF).replace('.', ',')
            CODIGO_RECOLHIMENTO_IRRF = '170806'
            CODIGO_RECOLHIMENTO_CSRF = '595207'
            
            arquivo.write(f'|1000|{ESPECIE}|{CNPJ}||{ACUMULADOR}|{CFOP}||{NF}|U||{ENTRADA}|{ENTRADA}|{VALOR}||||||||||||||||||||||||||{VALOR}||||||||||||||||||||||||||||||||||||||||||||||||||||||||||\n')
            
            if IRRF > 0:
                arquivo.write(f'|1020|16||{VALOR}|1,50|{VALOR_IRRF}|||||{VALOR}|{CODIGO_RECOLHIMENTO_IRRF}||||\n')
            
            if CSRF > 0:
                arquivo.write(f'|1020|25||{VALOR}|4,65|{VALOR_CSRF}|||||{VALOR}|{CODIGO_RECOLHIMENTO_CSRF}||||\n')
    
    return output_filename

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=8080)
