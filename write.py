from pptx import Presentation

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
            run = p.add_run()
            
            if text_index < len(new_text["shapes"]):
                run.text = new_text["shapes"][text_index]["text_content"]
                text_index += 1

        prs.save('modified_education_presentation_fixed.pptx')
    return 'modified_education_presentation_fixed.pptx'



new_texts = [
  {
    "slide_index": 1,
    "shapes": [
      {
        "text_content": "Addition",
        "width": 5748000,
        "height": 1886400,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "Pairs of numbers that add up to 10. Visual exercises about addition",
        "width": 4493400,
        "height": 833400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "1+1",
        "width": 663498,
        "height": 379749,
        "font_size": None
      },
      {
        "text_content": "2+4",
        "width": 858005,
        "height": 384909,
        "font_size": None
      },
      {
        "text_content": "3+3",
        "width": 873485,
        "height": 390585,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 2,
    "shapes": [
      {
        "text_content": "01",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Introduction",
        "width": 2704500,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "02",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "What is addition?",
        "width": 3181800,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "03",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Let\u2019s do it!",
        "width": 2223600,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "04",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Horizontal & vertical addition",
        "width": 5382900,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "05",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Pairs of numbers that add up to 10",
        "width": 5933100,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "06",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Mental calculation",
        "width": 3361200,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Sections",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "07",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Visual exercises about addition",
        "width": 5184600,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "08",
        "width": 771300,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "Numerical logic problems",
        "width": 5261700,
        "height": 612300,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "6+6",
        "width": 823956,
        "height": 384910,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 3,
    "shapes": [
      {
        "text_content": "Introduction",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "2+8",
        "width": 860585,
        "height": 393166,
        "font_size": None
      },
      {
        "text_content": "3+2",
        "width": 879159,
        "height": 390587,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "Hello! I am the addition. I love to unite, to join, to add.... To be more! Join my club :)",
        "width": 3994500,
        "height": 1328400,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 4,
    "shapes": [
      {
        "text_content": "",
        "width": 3090600,
        "height": 1269900,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 3994500,
        "height": 1269900,
        "font_size": None
      },
      {
        "text_content": "What is addition?",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "There are 3 red apples and 4 yellow apples. Let\u2019s put them together!",
        "width": 3994500,
        "height": 1023600,
        "font_size": None
      },
      {
        "text_content": "Look at these apples",
        "width": 3248700,
        "height": 648600,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "1+1",
        "width": 663498,
        "height": 379749,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 5,
    "shapes": [
      {
        "text_content": "That's right! There are 7 apples: 3 red and 4 yellow. 7 in total!",
        "width": 3025200,
        "height": 1328400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 3395400,
        "height": 3139500,
        "font_size": None
      },
      {
        "text_content": "What is addition?",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "3 + 4",
        "width": 3395400,
        "height": 1269900,
        "font_size": None
      },
      {
        "text_content": "We've put them all together in the fruit bowl. Can you help me count them?",
        "width": 3025200,
        "height": 1328400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 6,
    "shapes": [
      {
        "text_content": "",
        "width": 4177800,
        "height": 1449600,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 463500,
        "height": 463500,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 463500,
        "height": 463500,
        "font_size": None
      },
      {
        "text_content": "3 red and 4 yellow equals 7 apples",
        "width": 4177800,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "3 + 4 = 7",
        "width": 4177800,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "What is addition?",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "In math, to join, to add, to put together... is represented with me (+). My name is plus",
        "width": 4177800,
        "height": 1449600,
        "font_size": None
      },
      {
        "text_content": "I am the equal (=). After me, you will write the result of the addition",
        "width": 3242700,
        "height": 1449600,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 7,
    "shapes": [
      {
        "text_content": "What is addition?",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "My goal is to unite. Look at this other example",
        "width": 6858600,
        "height": 648600,
        "font_size": None
      },
      {
        "text_content": "=",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "and",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "6",
        "width": 3662100,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "3",
        "width": 1938900,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "+",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "=",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "9",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "9",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "=",
        "width": 963128,
        "height": 814715,
        "font_size": None
      },
      {
        "text_content": "There are 9 balls",
        "width": 3248700,
        "height": 648600,
        "font_size": None
      },
      {
        "text_content": "I\u2019m equally important",
        "width": 2475600,
        "height": 648600,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 8,
    "shapes": [
      {
        "text_content": "",
        "width": 4177800,
        "height": 2357100,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 4177800,
        "height": 2357100,
        "font_size": None
      },
      {
        "text_content": "Let\u2019s practice!",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "These are",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "fingers",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "These are",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "fingers",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "and",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "+",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "=",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "equals",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 9,
    "shapes": [
      {
        "text_content": "",
        "width": 4177800,
        "height": 2357100,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 4177800,
        "height": 2357100,
        "font_size": None
      },
      {
        "text_content": "Let\u2019s practice!",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "There are",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "apples",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "There are",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "apples",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "and",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "+",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "=",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "equals",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 10,
    "shapes": [
      {
        "text_content": "",
        "width": 4177800,
        "height": 2357100,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 4177800,
        "height": 2357100,
        "font_size": None
      },
      {
        "text_content": "Let\u2019s practice!",
        "width": 8660400,
        "height": 792000,
        "font_size": None
      },
      {
        "text_content": "This template has been created by Slidesgo",
        "width": 3681600,
        "height": 336000,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "and",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "+",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "=",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "equals",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 1600200,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "There are",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "cars",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 794400,
        "height": 794400,
        "font_size": None
      },
      {
        "text_content": "There are",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "cars",
        "width": 1287300,
        "height": 540300,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 217018,
        "height": 207799,
        "font_size": None
      },
      {
        "text_content": "",
        "width": 468555,
        "height": 473852,
        "font_size": None
      }
    ]
  },
  {
    "slide_index": 11,
    "shapes": [
      {
        "text_content": "Thank you",
        "width": 4331970,
        "height": 1015663,
        "font_size": None
      }
    ]
  }
]

replace_text_in_text_frames('slide.pptx', new_texts)
