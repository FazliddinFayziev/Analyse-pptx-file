# Analyse pptx file and returns json data with these values:
# text_content, width (width of text frame), height (width of text frame), font_size (font size will be set to None in default)

from pptx import Presentation
import json

def get_text_boxes_info(pptx_path):
    presentation = Presentation(pptx_path)
    all_slides = []

    for slide_index, slide in enumerate(presentation.slides):
        slide_data = {
            'slide_index': slide_index + 1,
            'shapes': []
        }

        for shape in slide.shapes:
            if shape.has_text_frame:
                text_content = ""
                text_frame = shape.text_frame
                width = shape.width
                height = shape.height
                font_size = None

                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        text_content += run.text

                text_box_info = {
                    'text_content': text_content,
                    'width': width,
                    'height': height,
                    'font_size': font_size
                }

                slide_data['shapes'].append(text_box_info)

        all_slides.append(slide_data)

    return all_slides

# Example usage
slides_data = get_text_boxes_info('slide.pptx')

# Print the resulting JSON data
print(json.dumps(slides_data, indent=2))

# Correct
