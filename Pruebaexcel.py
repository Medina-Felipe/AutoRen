import pandas as pd
from docxtpl import DocxTemplate

excel_file = "data.xlsx"
df = pd.read_excel(excel_file)

template_file = "template.docx"
doc = DocxTemplate(template_file)

context = {"table": df.to_dict('records')}

doc.render(context)

output_file = "output.docx"
doc.save(output_file)