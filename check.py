import fitz  # PyMuPDF library

def extract_text_and_font(pdf_path):
    doc = fitz.open(pdf_path)

    for page_number in range(doc.page_count):
        page = doc[page_number]
        text_blocks = page.get_text("blocks")

        for block in text_blocks:
            for line in block:
                font_size = line["size"]
                text_content = line["text"]
                
                print(f"Page {page_number + 1}, Font Size: {font_size}, Text Content: {text_content}")

    doc.close()

# Example usage
pdf_path = 'done.pdf' 
extract_text_and_font(pdf_path)
