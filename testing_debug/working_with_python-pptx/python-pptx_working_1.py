from pptx import Presentation

#
# Code explanation follows code here
#

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "python-pptx was here!"

prs.save('test.pptx')


#
# Code explanation
#

diagram = """
--------------------------------------------------------------
BUILD A NEW POWERPOINT PRESENTATION WITH:
   A TILE SLIDE AND TEXT (SHAPE)
   A SUBTITLE PLACEHOLER WITH TEXT
--------------------------------------------------------------

CODE:
from pptx import Presentation

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "python-pptx was here!"

prs.save('test.pptx')

CODE DIAGRAMMED
 ________________________________________
|    ________________________________    |
|   |                                |   |<--- prs = Presentation()
|   |                                |   |
|   |       ---------------          |<--|---- title_slide_layout = prs.slide_layouts[0]
|   |       | Hello World | <--------|---|---- title = slide.shapes.title
|   |       --------------- <--------|---|---- title.text = 'Hello World'
|   |                                |   |
|   |                                |   |
|   |   ------------------------- <--|---|---- subtitle = slide.placeholders[1]
|   |   | python-pptx was here! | <--|---|---- subtitle.text = 'python-pptx was here!'
|   |   -------------------------    |   |
|   |                                |   |
|   |                                |   |
|   |                                |   |
|   |                                |   |
|   |________________________________|   |
|________________________________________|

"""

print(diagram)
