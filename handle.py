import pymupdf
import os

def extract_text_images_from_pdf(pdf_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    for page in doc:
        text = page.get_text()

        # Save text
        with open(os.path.join(output_dir, f'page_{page.number + 1}.txt'), 'w', encoding='utf-8') as text_file:
            text_file.write(text)

        # # Save images
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            img_path = os.path.join(output_dir, f"page_{page.number + 1}_img_{img_index + 1}.png")
            with open(img_path, "wb") as img_file:
                img_file.write(image_bytes)

def extract_text_images_from_docx(docx_path, output_dir):
    print(docx_path)


pdf_path = 'pdf_mock_file.pdf'
pdf_output_dir = './pdf_extract'
doc_path = 'docx_mock_file.pdf'
doc_output_dir = './docx_extract'

extract_text_images_from_docx(doc_path, doc_output_dir)
# extract_text_images_from_pdf(pdf_path, pdf_output_dir)
