from pypdf import PdfReader 
from pptx import Presentation
from pathlib import Path

reader = PdfReader("document.pdf")
page = reader.pages[0]
print(page.extract_text())

prs = Presentation('test.pptx')

SLD_LAYOUT_TITLE_AND_CONTENT = 1

slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_AND_CONTENT]
slide = prs.slides.add_slide(slide_layout)







prs.save('test.pptx')






