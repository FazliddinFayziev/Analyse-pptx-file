from pptx import Presentation
from pptx.util import Pt

pptx_file = 'sample.pptx'
presentation = Presentation(pptx_file)

for slide_number, slide in enumerate(presentation.slides):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    font_size = run.font.size
                    font_size_pt = font_size.pt if font_size else 'Default size'
                    print(f"Slide {slide_number + 1}, Text: {run.text}, Font size: {font_size_pt}")
                    if font_size and font_size < Pt(24):
                        pass
