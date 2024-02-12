#pip install python-docx
from docx import Document
import pandas as pd

def fill_invitation(template_path, output_path, data):
    doc = Document(template_path)
    
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                #paragraph.text = paragraph.text.replace(key, value)#if no formating issue, each paragraph has a clear placeholder
                for run in paragraph.runs:
                    run.text = run.text.replace(key, value)
    doc.save(output_path)

def generate_invitation_from_csv(csv_path, template_path):
    df = pd.read_csv(csv_path)
    for idx, row in df.iterrows():
        data = {
            '[Salutation]': row['salutation'],
            '[First Name]': row['first_name'],
            '[Last Name]': row['last_name'],
            '[Last Contacted]': row['last_contacted'],
            '[Company Name]': row['company_name']
        }
        output_path = f'invitation_{idx + 1}_{row['first_name']}_{row['last_name']}_.docx'
        fill_invitation(template_path, output_path, data)

if __name__ == '__main__':
    csv_path = 'contacts.csv'
    template_path = 'template.docx'
    generate_invitation_from_csv(csv_path, template_path)