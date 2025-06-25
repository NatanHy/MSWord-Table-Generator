import docx.document
from docx.shared import Cm

# ===============================================
# Configure document by changing these variables
# ===============================================
# page margins, in centimeters
LEFT_MARGIN = 1.5
RIGHT_MARGIN = 1.5

def configure_document(doc : docx.document.Document): 
    sections = doc.sections

    for section in sections:
        section.left_margin = Cm(LEFT_MARGIN)
        section.right_margin = Cm(RIGHT_MARGIN)