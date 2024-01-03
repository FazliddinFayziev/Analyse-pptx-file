from pptx import Presentation

pptx_file = 'education.pptx'
presentation = Presentation(pptx_file)

for slide_number, slide in enumerate(presentation.slides):
    # Iterate through each shape in the slide
    for shape in slide.shapes:
        if shape.has_text_frame:
            # Iterate through each paragraph in the text frame
            for paragraph in shape.text_frame.paragraphs:
                font_size = paragraph.font.size
                # font_size is in Pt, convert to a human-readable format if necessary
                font_size_pt = font_size.pt if font_size else 'Default size'
                print(f"Slide {slide_number + 1}, Text: {paragraph.text}, Font size: {font_size_pt}")