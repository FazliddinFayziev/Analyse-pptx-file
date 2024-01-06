from pptx import Presentation
from pptx.util import Pt
import requests
from io import BytesIO

def emu_to_pt(value):
    return round(value / 12700)

def area(width, height):
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
        
        # Remove old images
        for shape in slide.shapes:
            if hasattr(shape, 'image'):
                slide.shapes._spTree.remove(shape._element)

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
        
                if "font_size" in text_content_data:
                    font_size_value = text_content_data["font_size"]
                    text = text_covered(len(run.text), font_size_value)
                    square = area(text_content_data["width"], text_content_data["height"])
        
                    while text > square:
                        font_size_value = reduce_font_size(font_size_value)
                        text = text_covered(len(run.text), font_size_value)
                        square = area(text_content_data["width"], text_content_data["height"])
                    
                    run.font.size = Pt(font_size_value)
                text_index += 1
        
            if new_text["image_path"]:
              for image_index, image in enumerate(new_text["image_path"]):
                image_url = image["path"]
                response = requests.get(image_url)

                if response.status_code == 200:
                  image_data = BytesIO(response.content)
                  left = image["left"]
                  top = image["top"]
                  width = image["width"]
                  height = image["height"]
                  slide.shapes.add_picture(image_data, left, top, width, height)

              new_text["image_path"] = []
              text_index += 1
              
    modified_pptx_path = 'modified_education_presentation_fixed.pptx'
    prs.save(modified_pptx_path)
    return modified_pptx_path

# Example usage
new_texts = [
    {
        "slide_index": 1,
        "shapes": [
            {
                "id": 0,
                "width": 7593330,
                "height": 3035808,
                "text_content": "Sample PowerPoint File",
                "font_size": 64
            },
            {
                "id": 1,
                "width": 5918454,
                "height": 1069848,
                "text_content": "St. Cloud Technical College",
                "font_size": 18
            }
        ],
        "image_path": []
    },
    {
        "slide_index": 2,
        "shapes": [
            {
                "id": 0,
                "width": 7772400,
                "height": 1609344,
                "text_content": "This is a Sample Slide",
                "font_size": 48
            },
            {
                "id": 1,
                "width": 7924800,
                "height": 3510776,
                "text_content": "You can print out PPT files as handouts using the PRINT >   PRINT WHAT > HANDOUTS option\nHere is an outline of bulleted points\n",
                "font_size": 32
            }
        ],
        "image_path": []
    },
    {
        "slide_index": 3,
        "shapes": [
            {
                "id": 0,
                "width": 7772400,
                "height": 1609344,
                "text_content": "My IMAGE with code",
                "font_size": 48
            }
        ],
        "image_path": [
        {
          "path": "https://upload.wikimedia.org/wikipedia/commons/5/5c/97-979947_writing-writer-essay-logo-act-writer-logo-png.png",
          "width": 3072857,
          "height": 4277600,
          "top": 1884407,
          "left": 5037798
        },
        {
          "path": "https://www.freeiconspng.com/thumbs/writing-png/3d-man-writing-png-11.png",
          "width": 3072857,
          "height": 4277600,
          "top": 1884407,
          "left": 1033345
        }
      ]
    }
]

replace_text_in_text_frames('example.pptx', new_texts)
print('Done')
