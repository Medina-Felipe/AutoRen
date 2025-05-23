import pdfplumber
import re

def extract_data_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Leer todas las páginas
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            
            # Función para extraer datos usando expresiones regulares
            def extract_field(pattern):
                match = re.search(pattern, text)
                return match.group(1).strip() if match else None

            # Extraer los datos
            data = {
                'nombre': extract_field(r'Nombre o Razón Social:\s*(.*?)(?:\n|$)'),
                'rut': extract_field(r'Rut:\s*(.*?)(?:\n|$)'),
                'giro': extract_field(r'Giro:\s*(.*?)(?:\n|$)'),
                'direccion': extract_field(r'Dirección Comercial:\s*(.*?)(?:\n|$)'),
                'ciudad': extract_field(r'Ciudad:\s*(.*?)(?:\n|$)'),
                'contacto': extract_field(r'Contacto:\s*(.*?)(?:\n|$)'),
                'email': extract_field(r'e-mail:\s*(.*?)(?:\n|$)'),
                'proyecto': extract_field(r'Proyecto Asociado:\s*(.*?)(?:\n|$)')
            }
            
            return data

    except Exception as e:
        print(f"Error al procesar el PDF: {str(e)}")
        return None

if __name__ == "__main__":
    pdf_path = "Orden de Ingreso.pdf"
    data = extract_data_from_pdf(pdf_path)
    
    if data:
        print("Datos extraídos del PDF:")
        for key, value in data.items():
            print(f"{key}: {value}")
    else:
        print("No se pudieron extraer los datos del PDF")

