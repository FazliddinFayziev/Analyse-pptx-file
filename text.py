from pptx import Presentation

def get_font_size(pptx_path):
    presentation = Presentation(pptx_path)

    all_text_sizes = []

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text_content = run.text
                        font_size = run.font.size
                        all_text_sizes.append({'text_content': text_content, 'font_size': font_size})

    return all_text_sizes

# Example usage
pptx_path = 'education.pptx'
text_sizes = get_font_size(pptx_path)

# Print the results
for item in text_sizes:
    print(f"Text: {item['text_content']}, Font Size: {item['font_size']}")
