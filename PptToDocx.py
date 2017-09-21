from pptx import Presentation
from docx import Document
from docx.shared import Cm

from docx.enum.text import WD_ALIGN_PARAGRAPH

def pptToDocx(INPUT_FILE,OUTPUT_DIR,margin = 2):
    # Save docx with same filename
    OUTPUT_FILE = OUTPUT_DIR+(INPUT_FILE.split('/')[-1].split('.')[0]+'.docx')

    # Load presentation and create document file
    prs = Presentation(INPUT_FILE)
    document = Document()   

    # Set page margins
    for section in document.sections:
        section.top_margin = Cm(margin)
        section.bottom_margin = Cm(margin)
        section.left_margin = Cm(margin)
        section.right_margin = Cm(margin)

    # Image counter
    counter = 0

    isFirstSlide = True
    for slide in prs.slides:
        isFirst = True
        for shape in slide.shapes:
            # Autosave images
            if not shape.has_text_frame:
                filename = str(counter)+'.'+shape.image.ext
                with open(OUTPUT_DIR+filename,'wb') as f:
                    f.write(shape.image.blob)
                document.add_picture(OUTPUT_DIR+filename, width=Cm(14))
                document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                counter += 1
                continue
            
            for paragraph in shape.text_frame.paragraphs:
                # Make the first paragraph the header
                if isFirst:
                    # Make first paragraph in the first slide the document title
                    if isFirstSlide:
                        document.add_heading(paragraph.text, 0)
                        isFirstSlide = False
                    else:
                        document.add_heading(paragraph.text, 1)
                    isFirst = False
                else:
                    # Ignore empty paragraphs
                    if len(paragraph.text) > 0:
                        document.add_paragraph(paragraph.text)
                        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Save the output
    document.save(OUTPUT_FILE)

INPUT_FILE = './input/Review1.pptx'
OUTPUT_DIR = './output/'

pptToDocx(INPUT_FILE,OUTPUT_DIR)

