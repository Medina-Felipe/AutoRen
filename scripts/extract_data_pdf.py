from datetime import datetime

import pdfplumber
import re

def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    def extract(patron, multiline=False):
        flags = re.IGNORECASE | re.DOTALL if multiline else re.IGNORECASE
        match = re.search(patron, text, flags)
        if match:
            result = match.group(1).strip()
            return re.sub(r'\s*\n\s*', ' ', result)  # Reemplaza saltos de línea con espacio
        return ""

    data = {
        '[Nombre]': extract(r'Nombre:\s*(.+)'),
        '[FechaOrdenIngreso]': extract(r'Fecha:\s*(\d{2}/\d{2}/\d{4})'),
        '[FechaEmision]': datetime.now().strftime("%d/%m/%Y"),
        '[Rut]': extract(r'Rut:\s*\n*\s*([\d\.]+-[Kk0-9])', multiline=True),
        '[Institucion]': extract(r'Institución:\s*(.+)'),
        '[Giro]': extract(r'Giro:\s*(.+?)\s*(?:Proyecto Asociado:)', multiline=True),
        '[ProyectoAsociado]': extract(r'Proyecto\s*Asociado:\s*(.+?)\s*(?:Email:)', multiline=True),
        '[Email]': extract(r'Email:\s*([^\s\n]+)'),
        '[Contacto]': extract(r'Contacto:\s*(\+\d{8,15})'),
        '[Ciudad]': extract(r'Ciudad:\s*([^\n]+)'),
        '[Equipo]': extract(r'Equipo:\s*(.+)'),
        '[Analisis]': extract(r'Análisis:\s*(.+)'),
        '[NumMuestras]': extract(r'N° de Muestras:\s*(\d+)'),
        '[NroCotizacion]': extract(r'N° de Cotización:\s*(\d+)'),
        '[Observaciones]': extract(r'Observaciones:\s*(.+?)\s*(?:Firma|Cláusula)', multiline=True)
    }

    return data

if __name__ == "__main__":
    pdf_path = "C:\\Users\\Adolfo\\Desktop\\A.pdf"
    data = extract_data_from_pdf(pdf_path)
    print(data)
