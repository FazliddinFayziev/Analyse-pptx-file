from pptx import Presentation
import json

def get_shapes_info(pptx_path):
    presentation = Presentation(pptx_path)
    all_slides = []

    for slide_index, slide in enumerate(presentation.slides):
        slide_data = {
            'slide_index': slide_index + 1,
            'shapes': [],
            'image_path': []
        }

        index = 0
        for shape in slide.shapes:
            shape_info = {
                'id': index,
                'width': shape.width,
                'height': shape.height
            }

            if shape.has_text_frame:
                shape_info['text_content'] = ""
                shape_info['font_size'] = None

                text_frame = shape.text_frame

                for paragraph_index, paragraph in enumerate(text_frame.paragraphs):
                    for run_index, run in enumerate(paragraph.runs):
                        shape_info['text_content'] += run.text

                        # Add newline character if it's the end of a paragraph
                        if run_index == len(paragraph.runs) - 1 and paragraph_index != len(text_frame.paragraphs) - 1:
                            shape_info['text_content'] += "\n"

                        # Check for font size
                        if hasattr(run, 'font') and hasattr(run.font, 'size'):
                            shape_info['font_size'] = run.font.size if run.font.size else None

            elif shape.shape_type == 13:
                image_info = {
                    "path": None,
                    "width": shape.width,
                    "height": shape.height,
                    "top": shape.top,
                    "left": shape.left
                }
                slide_data['image_path'].append(image_info)

            if shape.shape_type != 13:  # Exclude image information from shapes list
                slide_data['shapes'].append(shape_info)
                index += 1

        all_slides.append(slide_data)

    return all_slides

# Example usage
slides_data = get_shapes_info('example.pptx')

# Print the resulting JSON data
print(json.dumps(slides_data, indent=2))
