import pymupdf
import os
from docx import Document
from docx.shared import Pt, RGBColor

### 1 ###
def extract_text_images_from_pdf(pdf_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    all_text = []
    for page in doc:
        text = page.get_text()

        if text.strip():
            all_text.append(text)

        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            img_path = os.path.join(output_dir, f"page_{page.number + 1}_img_{img_index + 1}.png")
            with open(img_path, "wb") as img_file:
                img_file.write(image_bytes)

    with open(os.path.join(output_dir, 'extract_text.txt'), 'w', encoding='utf-8') as file:
        for text in all_text:
            file.write(text + '\n')


def extract_text_from_docx(docx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    doc = Document(docx_path)

    all_text = []
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            all_text.append(paragraph.text)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text:
                    all_text.append(cell.text)

    with open(os.path.join(output_dir, 'extract_text.txt'), 'w', encoding='utf-8') as file:
        for text in all_text:
            file.write(text + '\n')

    return all_text

    


### 2 ###
def extract_formatting_from_docx(docx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    doc = Document(docx_path)
    formatting_details = []

    with open(os.path.join(output_dir, 'extract_formatting_details_text.txt'), 'w', encoding='utf-8') as file:
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                para_details = {'text': paragraph.text, 'runs': []}
                file.write(f"Text: {paragraph.text}\n")
                for run in paragraph.runs:
                    font_name = run.font.name if run.font.name else "Default"
                    font_size = run.font.size.pt if run.font.size else "Default"
                    bold = "Yes" if run.bold else "No"
                    italic = "Yes" if run.italic else "No"
                    text_color = run.font.color.rgb if run.font.color and run.font.color.rgb else "Default"
                    file.write(f" - Font: {font_name}, Size: {font_size}, Bold: {bold}, Italic: {italic}, Color: {text_color}\n")
                    run_details = {
                        'text': run.text,
                        'font_name': run.font.name if run.font.name else "Default",
                        'font_size': run.font.size.pt if run.font.size else "Default",
                        'bold': run.bold,
                        'italic': run.italic,
                        'text_color': run.font.color.rgb if run.font.color and run.font.color.rgb else "Default"
                    }
                    para_details['runs'].append(run_details)
                formatting_details.append(para_details)
                file.write("\n")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            para_details = {'text': paragraph.text, 'runs': []}
                            file.write(f"Text: {paragraph.text}\n")
                            for run in paragraph.runs:
                                font_name = run.font.name if run.font.name else "Default"
                                font_size = run.font.size.pt if run.font.size else "Default"
                                bold = "Yes" if run.bold else "No"
                                italic = "Yes" if run.italic else "No"
                                text_color = run.font.color.rgb if run.font.color and run.font.color.rgb else "Default"
                                file.write(f" - Font: {font_name}, Size: {font_size}, Bold: {bold}, Italic: {italic}, Color: {text_color}\n")
                                run_details = {
                                    'text': run.text,
                                    'font_name': run.font.name if run.font.name else "Default",
                                    'font_size': run.font.size.pt if run.font.size else "Default",
                                    'bold': run.bold,
                                    'italic': run.italic,
                                    'text_color': run.font.color.rgb if run.font.color and run.font.color.rgb else "Default"
                                }
                                para_details['runs'].append(run_details)
                            formatting_details.append(para_details)
                            file.write("\n")
                        file.write("\n")
    return formatting_details

def extract_formatting_from_pdf(pdf_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    with open(os.path.join(output_dir, 'extract_formatting_details_text.txt'), 'w', encoding='utf-8') as file:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            blocks = page.get_text("dict")["blocks"]
            for b in blocks:
                if b["type"] == 0:
                    for line in b["lines"]:
                        for span in line["spans"]:
                            file.write(f"Text: {span['text']}\n")
                            file.write(f" - Font: {span['font']}\n")
                            file.write(f" - Size: {span['size']}\n")
                            file.write(f" - Color: {span['color']}\n")
            file.write("\n")     

def compile_text_from_docx_to_uppercase(docx_path, output_dir):
    text = extract_text_from_docx(docx_path, output_dir)
    formatting = extract_formatting_from_docx(docx_path, output_dir)

    new_doc = Document()
    
    for para in formatting:
        new_para = new_doc.add_paragraph()
        for run in para['runs']:
            new_run = new_para.add_run(run['text'].upper())
            new_run.font.name = run['font_name']
            if run['font_size'] != "Default":
                new_run.font.size = Pt(run['font_size'])
            new_run.bold = run['bold']
            new_run.italic = run['italic']
            if run['text_color'] != "Default":
                new_run.font.color.rgb = RGBColor(run['text_color'][0], run['text_color'][1], run['text_color'][2])
    
    new_doc.save(os.path.join(output_dir, 'output_uppercased_docx_file.docx'))


pdf_path = 'pdf_mock_file.pdf'
pdf_output_dir = './pdf_extract'
doc_path = 'docx_mock_file.docx'
doc_output_dir = './docx_extract'


extract_text_from_docx(doc_path, doc_output_dir)
extract_text_images_from_pdf(pdf_path, pdf_output_dir)

extract_formatting_from_docx(doc_path, doc_output_dir)
extract_formatting_from_pdf(pdf_path, pdf_output_dir)

compile_text_from_docx_to_uppercase(doc_path, doc_output_dir)