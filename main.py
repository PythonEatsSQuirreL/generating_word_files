#pip install python-docx
from docx import Document

def fill_invitation(template_path, output_path, data):
    doc = Document(template_path)
    
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                #paragraph.text = paragraph.text.replace(key, value)#if no formating issue, each paragraph has a clear placeholder
                for run in paragraph.runs:
                    run.text = run.text.replace(key, value)
    doc.save(output_path)
    
if __name__ == '__main__':
    data = {
        '[Salutation]': 'Mr.',
        '[First Name]': 'Mike',
        '[Last Name]': 'Smith',
        '[Last Contacted]': 'October 1, 2023',
        '[Company Name]': 'Smith Inc.'
    }
    
    template_path = 'template.docx'
    output_path = 'filled.docx'
    
    fill_invitation(template_path, output_path, data)