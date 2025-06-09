from flask import Flask, request, render_template, send_file, redirect, flash
import os
from werkzeug.utils import secure_filename
from extractor import process_single_order

app = Flask(__name__)
app.secret_key = 'clave_secreta'
app.config['UPLOAD_FOLDER'] = 'uploads'

TEMPLATE_PATH = 'template.docx'
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        pdf_file = request.files.get('file_pdf')
        excel_file = request.files.get('file_excel')

        if not pdf_file or pdf_file.filename == '' or not allowed_file(pdf_file.filename):
            flash('Debe subir un archivo PDF válido.')
            return redirect(request.url)

        if not excel_file or excel_file.filename == '':
            flash('Debe subir un archivo Excel.')
            return redirect(request.url)

        pdf_filename = secure_filename(pdf_file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
        pdf_file.save(pdf_path)

        excel_filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        excel_file.save(excel_path)

        # Aquí solo procesamos el PDF, puedes usar excel_path si lo necesitas
        success, output_filename = process_single_order(pdf_path, TEMPLATE_PATH)

        if success:
            return send_file(output_filename, as_attachment=True)
        else:
            flash('Error procesando el informe.')
            return redirect(request.url)

    return render_template('home.html')


if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    app.run(host='0.0.0.0', port=5000, debug=True)
