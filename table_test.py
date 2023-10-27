import os
from docx import Document

doc=Document("Projectreports/22-1297_Tobias-Nimz.docx")

#Initialise flag
found_heading=False

# Define namespaces
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

#Loop for every table in the document
for e in doc.element.body:
    #if found_heading:
    #check if the element is a paragraph
    if e.tag.endswith("p"):
        p=e
        p_style_element=p.find(".//w:pStyle", namespaces=ns)
        if p_style_element is not None:
            p_style=p_style_element.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
        # Check if the paragraph has the style "Heading 2"
            if p_style=='Heading 2':
                found_heading=True
    # Check if the element is a table and we've found a Heading
    elif e.tag.endswith('tbl') and found_heading==True:
        t=e
        print("Table after Heading 2 paragraph:")
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    print(p.text)
        
