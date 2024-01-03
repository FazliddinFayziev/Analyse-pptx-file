from pptx import Presentation
from pptx.util import Pt

def replace_text_in_text_frames(pptx_path, new_texts):
    prs = Presentation(pptx_path)
    
    for new_text in new_texts:
        index_slide = new_text["slide_index"] - 1  # Adjusting for 0-based index
        slide = prs.slides[index_slide]
        
        text_index = 0
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            
            text_frame = shape.text_frame
            text_frame.clear()

            p = text_frame.paragraphs[0]
            
            if text_index < len(new_text["shapes"]):
                text_content_data = new_text["shapes"][text_index]
                run = p.add_run()
                run.text = text_content_data["text_content"]

                # Check if font_size is specified, otherwise use default (12 points)
                font_size = text_content_data.get("font_size", 12)

                # Check if width is specified to handle text overflow
                width = text_content_data.get("width", None)

                if font_size is not None:
                    if width and font_size:
                        # Calculate the available width for the text based on the shape's width
                        available_width = shape.width - width
                        if available_width < 0:
                            # If the available width is negative, adjust the font size to fit the text
                            reduction_ratio = shape.width / (shape.width - available_width)
                            run.font.size = Pt(max(1, font_size / reduction_ratio))  # Ensure font size is at least 1
                        else:
                            run.font.size = Pt(font_size)
                    else:
                        run.font.size = Pt(font_size)

                text_index += 1

        prs.save('modified_education_presentation_fixed.pptx')

    return 'modified_education_presentation_fixed.pptx'

# Example usage
new_texts = [
    {
        "slide_index": 1, "shapes": [
            {"text_content": "Elegant Education Pack for Students in Malaysia and United States", "width": 6577800, "height": 2571750, "font_size": None}, 
            {"text_content": "Here is where your presentation begins", "width": 6577800, "height": 458100, "font_size": None}, 
            {"text_content": "Done by Fazliddin", "width": 6577800, "height": 458100, "font_size": None}
        ]
    }
]

replace_text_in_text_frames('education.pptx', new_texts)
print('done')
