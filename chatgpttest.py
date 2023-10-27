import os
from docx import Document

doc = Document("Projectreports/22-1297_Tobias-Nimz.docx")

# Initialize flag
found_heading = False

# Define namespaces
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# Loop for every element in the document
for element in doc.element.body:
    # Check if we've found a Heading 2 paragraph
    if found_heading:
        if element.tag.endswith('tbl'):
            # This element is a table
            t = element
            print("Table after Heading 2 paragraph:")
            #for row in t.iter("w:tr", namespaces=ns):
                #for cell in row.iter("w:tc", namespaces=ns):
                    #for p in cell.iter("w:p", namespaces=ns):
                        #print(p.text)
    elif element.tag.endswith("p"):
        #print('paragraph found')
        p = element
        p_style_element = p.find(".//w:pStyle", namespaces=ns)
        if p_style_element is not None:
            #print('paragraph with style:')
            p_style = p_style_element.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
            
            print(p_style,p.text)
            # Check if the paragraph has the style "Heading 2"
            if p_style == 'berschrift2' and p.text=='Steckbrief':
                found_heading = True
                print(found_heading)