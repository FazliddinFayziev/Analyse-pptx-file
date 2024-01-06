from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

# Function to get font size of a text run
def get_font_size(run):
    try:
        font_size = run.font.size.pt if run.font.size else None
        return font_size
    except Exception as e:
        print(f"Error accessing font size: {e}")
        return None

# Function to get font sizes from a slide
def get_text_boxes_font_sizes(slide):
    font_sizes = []

    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    font_size = get_font_size(run)
                    font_sizes.append(font_size)

    return font_sizes

# Example usage
pptx_path = 'sample.pptx'
presentation = Presentation(pptx_path)

for slide_index, slide in enumerate(presentation.slides):
    font_sizes = get_text_boxes_font_sizes(slide)
    print(f"Slide {slide_index + 1} Font Sizes: {font_sizes}")
