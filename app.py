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
        file = request.files.get('file')
        if not file or file.filename == '':
            flash('Debe seleccionar un archivo PDF.')
            return redirect(request.url)

        if allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            success, output_filename = process_single_order(filepath, TEMPLATE_PATH)
            if success:
                return send_file(output_filename, as_attachment=True)
            else:
                flash('Error procesando el PDF.')
                return redirect(request.url)

    return render_template('index.html')

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    app.run(host='0.0.0.0', port=5000, debug=True)
