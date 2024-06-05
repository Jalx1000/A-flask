from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import os
import json

app = Flask(__name__)

PREDEFINED_HEADERS = ['Lista', 'Nombre', 'Correo',
                      'Número', 'Empresa', 'Categoría', 'Curso', 'Año', 'Mes']


@app.route('/')
def index():
    return render_template('index.html', headers=PREDEFINED_HEADERS)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.xlsx'):
        df = pd.read_excel(file)
        available_columns = df.columns.tolist()
        return jsonify({'columns': available_columns}), 200

    return jsonify({'error': 'Invalid file type'}), 400


@app.route('/process', methods=['POST'])
def process_file():
    data = request.form.to_dict()
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.xlsx'):
        df = pd.read_excel(file)
        new_df = pd.DataFrame(columns=PREDEFINED_HEADERS)

        for header in PREDEFINED_HEADERS:
            selected_column = data.get(header)
            if selected_column and selected_column in df.columns:
                new_df[header] = df[selected_column]
            else:
                new_df[header] = ""

        output_filename = 'processed_file.xlsx'
        new_df.to_excel(output_filename, index=False)

        return send_file(output_filename, as_attachment=True)

    return jsonify({'error': 'Invalid file type'}), 400


if __name__ == '__main__':
    app.run(debug=True)
