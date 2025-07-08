from datetime import datetime

import pdfplumber
import re

def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    def extract(patron):
        match = re.search(patron, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    data = {
        '[Nombre]': extract(r'Nombre:\s*(.+)'),
        '[FechaOrdenIngreso]': extract(r'Fecha:\s*(\d{2}/\d{2}/\d{4})'),

        '[CentroFacultad]': extract(r'Centro o\s*Facultad:\s*(.*(?:\n.*Ingeniería\s+Química)?)'),

        '[Institucion]': extract(r'Institución:\s*(.+)'),
        '[NroCotizacion]': extract(r'N° de\s*Cotización:\s*(.+)'),

        '[ProyectoAsociado]': extract(r'Proyectos.*?:\s*([^\n]+(?:\n[^\n]+)*)'),

        '[Encargado]': extract(r'Encargado de Análisis:\s*(.+)'),

        '[Equipo]': extract(r'Equipo:\s*(.+)'),
        '[Analisis]': extract(r'Análisis:\s*(.+)'),

        '[NumMuestras]': extract(r'N° de Muestras:\s*(\d+)'),
        '[Observaciones]': extract(r'Observaciones:\s*(.+)'),

        '[FechaEmision]': datetime.now().strftime("%d/%m/%Y")
    }

    return data

if __name__ == "__main__":
    pdf_path = "C:\\Users\\Adolfo\\Desktop\\Orden de Ingreso.pdf"
    data = extract_data_from_pdf(pdf_path)
    print(data)
