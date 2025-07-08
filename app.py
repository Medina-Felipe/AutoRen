from flask import *
import os
from scripts.config import Config
from scripts.auth import login_user
from scripts.process_report import generate_report_from_data

app = Flask(__name__)
app.config.from_object(Config)

@app.route("/", methods=["GET", "POST"])
def login_page():
    if request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]
        user = login_user(email, password)
        if user:
            session["user"] = user["email"]
            return redirect("/dashboard")
        else:
            return render_template("login.html", error="Credenciales invalidas")
    return render_template("login.html")

@app.route("/logout")
def logout():
    if "user" not in session:
            return redirect("/")
    upload_dir = "uploads"
    output_dir = "output"
    
    if os.path.exists(upload_dir):
        for filename in os.listdir(upload_dir):
            file_path = os.path.join(upload_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)

    if os.path.exists(output_dir):
        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
    session.pop("generated_file", None)
    session.pop("user", None)
    session.pop("numero_informe", None)
    session.pop("observaciones", None)
    return redirect("/")

@app.route("/upload-files", methods=["POST"])
def upload_files():
    if "user" not in session:
        return redirect("/")

    pdf_file = request.files.get('file_pdf')
    excel_file = request.files.get('file_excel')

    numero_informe = request.form.get("numero_informe")
    observaciones = request.form.get("observaciones")

    session["numero_informe"] = numero_informe
    session["observaciones"] = observaciones

    if not pdf_file or not pdf_file.filename.lower().endswith('.pdf'):
        flash("Debes subir un archivo PDF.")
        return redirect("/dashboard")
    
    if not excel_file or not excel_file.filename.lower().endswith('.xlsx'):
        flash("Debes subir un archivo Excel.")
        return redirect("/dashboard")

    upload_dir = "uploads"
    os.makedirs(upload_dir, exist_ok=True)

    pdf_file.save(os.path.join(upload_dir, "orden_ingreso.pdf"))
    excel_file.save(os.path.join(upload_dir, "excel_datos.xlsx"))

    flash("Archivos y campos subidos correctamente.")
    
    return redirect("/dashboard")

@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect("/")
    
    upload_dir = "uploads"
    pdf_path = os.path.join(upload_dir, "orden_ingreso.pdf")
    excel_path = os.path.join(upload_dir, "excel_datos.xlsx")

    pdf_file = "orden_ingreso.pdf" if os.path.exists(pdf_path) else None
    excel_file = "excel_datos.xlsx" if os.path.exists(excel_path) else None

    return render_template("dashboard.html", pdf_file=pdf_file, excel_file=excel_file)

@app.route("/generate-report", methods=["POST"])
def generate_report():
    if "user" not in session:
        return redirect("/")

    pdf_path = "uploads/orden_ingreso.pdf"
    excel_path = "uploads/excel_datos.xlsx"

    if not os.path.exists(pdf_path) or not os.path.exists(excel_path):
        flash("Debes subir ambos archivos antes de generar el informe.")
        return redirect("/dashboard")

    manual_data = {
        '[NumInforme]': session.get("numero_informe"),
        '[Observaciones]': session.get("observaciones")
    }

    output_path = generate_report_from_data(pdf_path, excel_path, manual_data)
    
    if output_path:
        filename = os.path.basename(output_path)
        session["generated_file"] = filename  
    else:
        flash("Hubo un error al generar el informe.")

    return redirect("/dashboard")

@app.route("/download-report")
def download_report():
    if "user" not in session:
        return redirect("/")

    filename = session.get("generated_file")
    if not filename:
        flash("No hay informe para descargar.")
        return redirect("/dashboard")

    path = os.path.join("output", filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    else:
        flash("Archivo no encontrado.")
        return redirect("/dashboard")

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    app.run(host='0.0.0.0', port=5000, debug=True)