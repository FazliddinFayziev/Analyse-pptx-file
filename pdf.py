from spire.presentation import *
from spire.presentation.common import *

presentation = Presentation()
presentation.LoadFromFile("info.pptx")
slide = presentation.Slides[0]

slide.SaveToFile("done.pdf", FileFormat.PDF)
presentation.Dispose()