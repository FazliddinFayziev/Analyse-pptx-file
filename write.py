from pptx import Presentation
from pptx.util import Pt

def emu_to_pt(value):
    return round(value / 12700)

def perimeter(width, height):
    w = emu_to_pt(width)
    h = emu_to_pt(height)
    return w * h

def text_covered(text, font_size):
    return text * (font_size ** 2)

def reduce_font_size(font_size):
    return font_size - 1

def replace_text_in_text_frames(pptx_path, new_texts):
    prs = Presentation(pptx_path)

    for new_text in new_texts:
        index_slide = new_text["slide_index"] - 1
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

                font_size_value = text_content_data["font_size"]
                text = text_covered(len(run.text), font_size_value)
                square = perimeter(text_content_data["width"], text_content_data["height"])

                while text > square:
                  font_size_value = reduce_font_size(font_size_value)
                  text = text_covered(len(run.text), font_size_value)
                  square = perimeter(text_content_data["width"], text_content_data["height"])

                run.font.size = Pt(font_size_value)
                text_index += 1

    modified_pptx_path = 'modified_education_presentation_fixed.pptx'
    prs.save(modified_pptx_path)

    return modified_pptx_path

# Example usage
new_texts = [
    {
      "slide_index": 1,
      "shapes": [
        {"id": 0, "text_content": "Sample PowerPoint File", "width": 7593330, "height": 3035808, "font_size": 64},
        {"id": 1, "text_content": "St. Cloud Technical College", "width": 5918454, "height": 1069848, "font_size": 18},
      ]
    },
    {
      "slide_index": 2,
      "shapes": [
        {"id": 0, "text_content": "This is a Sample Slide", "width": 7772400, "height": 1609344, "font_size": 42},
        {
          "id": 1,
          "text_content": "You can print out PPT files as handouts using the PRINT >   PRINT WHAT > HANDOUTS option\nHere is an outline of bulleted points\n",
          "width": 7924800,
          "height": 3510776,
          "font_size": 32,
        },
      ]
    }
]

replace_text_in_text_frames('example.pptx', new_texts)
print('Done')
