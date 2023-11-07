from docx.text.paragraph import Paragraph
from docx.document import Document
from docx.table import _Cell, Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
import docx


def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            #yield Paragraph(child, parent)
            current_p=Paragraph(child,parent)
            yield current_p
        elif isinstance(child, CT_Tbl):
            table = Table(child, parent)
            for row in table.rows:
                for cell in row.cells:
                    yield from iter_block_items(cell)
                    
#doc = docx.Document('word.docx')
doc = docx.Document('Projectreports/22-1297_Tobias-Nimz.docx')

ignore_list=['Report Level','nach Z01D_Leitfaden','Dieses Dokument basiert auf der Report-Vorlage','Qualität und Bewertungskriterien', 'Hinweise zur Bearbeitung']

past_block=''
for block in iter_block_items(doc):
    current_block=block.text
    # ignore repeated cells, #ignore paragraphs starting with ignore_list
    if current_block != past_block: #and not current_block.startswith(ignore_list[0]) and not current_block.startswith(ignore_list[1]) and not current_block.startswith(ignore_list[2]) and not current_block.startswith(ignore_list[3]) and not current_block.startswith(ignore_list[4])
        print(block.text)
        print(block.style)
    past_block=current_block
