from pptx import Presentation
from pptx.util import Pt

def emu_to_pt(value):
    return round(value / 12700)

def perimeter(width, height):
    w = emu_to_pt(width)
    h = emu_to_pt(height)
    return (w + h) * 2

def average_char_width(font_size):
    return font_size * 0.8

def text_covered(text, font_size):
    avg_width = average_char_width(font_size)
    return len(text) * avg_width

def reduceFontSize(text, square, fontsize, run):
    additional_part = round((text - square) / len(run.text))
    if fontsize > additional_part:
      value = fontsize - additional_part
    else:
      value = additional_part - fontsize
    run.font.size = Pt(value)

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
                
                text = text_covered(run.text, text_content_data["font_size"])
                square = perimeter(text_content_data["width"], text_content_data["height"])

                if text > square:
                  reduceFontSize(text, square, text_content_data["font_size"], run)
                
                text_index += 1

        prs.save('modified_education_presentation_fixed.pptx')

    return 'modified_education_presentation_fixed.pptx'

# Example usage
new_texts = [
  {                                                                                                                                                               
    "slide_index": 1,
    "shapes": [
      {
        "id": 0,
        "text_content": "Sample PowerPoint File",
        "width": 7593330,
        "height": 3035808,
        "font_size": 64
      },
      {
        "id": 1,
        "text_content": "St. Cloud Technical College",
        "width": 5918454,
        "height": 1069848,
        "font_size": 18
      }
    ]
  },
  {
    "slide_index": 2,
    "shapes": [
      {
        "id": 0,
        "text_content": "This is a Sample Slide",
        "width": 7772400,
        "height": 1609344,
        "font_size": 42
      },
      {
        "id": 1,
        "text_content": "You can print out PPT files as handouts using the PRINT >   PRINT WHAT > HANDOUTS option\nHere is an outline of bulleted points\n",      
        "width": 7924800,
        "height": 3510776,
        "font_size": 32
      }
    ]
  }
]
replace_text_in_text_frames('example.pptx', new_texts)
print('done')
