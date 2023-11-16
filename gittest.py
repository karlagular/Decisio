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

#Dictionaries initialisation
#from the first table
D_Projektname={'Projektname':''}

#1. Projektdesign
D_Projektdesign={'Projektbeschreibung und -design':'', 'Beschreibung des Projekterfolges aus Sicht der Kunden/des Auftraggebers':'', 'Projektrelevanz und Einschätzung':''}

D_Steckbrief={'Projektbezeichnung':'', 
              'Projektnummer':'',
              'Auftraggeber':'',
              'Projektleitung':'',
              'Verantwortung':'',
              'Projektorganisationsform':'',
              'Oberziel':'',
              'Projektinhalt':'',
              'Projektinhalt':'',
              'Projektbeteiligte':'',
              'Projektumfeld':'',
              'Starttermin':'',
              'Endetermin': '',
              'Dauer': '',
              'Aufwand-gesamt':'',
              'Aufwand-intern':'',
              'Aufwand-extern':'',
              'Personalkosten-gesamt':'',
              'Personalkosten-intern':'',
              'Personalkosten-extern':'',
              'Investitionen-gesamt':'',
              'Investitionskosten-intern':'',
              'Investitionskosten-extern':'',
              'Budget-gesamt':'',
              'Budget-intern':'',
              'Budget-extern':'',
              'vrsl. Behinderungen/Risiken/Störungen':'',
              'Kunde':'',
              'Abnahmekriterien':''}

#
steckbrief_Suche=['Projektbezeichnung:',
                  'Projektnummer:',
                  'Auftraggeber (Name, Pos',
                  'Projektleitung, Verantw',
                  'Projektorgnisationsform',
                  'Oberziel:',
                  'Projektinhalt (was?):',
                  'Projekt']

#2. Anforderungen und Ziele
#Except for the Oberziel, it is possible for every other Sub-goal to be expanded later on
D_Ziele={'Oberziel-Zielbezeichnung':'',
         'Oberziel-Zielbeschreibung':'',
         'Oberziel-Messkriterium':'',
         'Finanzziel 1-Zielbezeicnung':'',
         'Finanzziel 1-Zielbeschreibung':'',
         'Finanzziel 1-Messkriterium':'',
         'Leistungsziel 1-Zielbezeicnung':'',
         'Leistungsziel 1-Zielbeschreibung':'',
         'Leistungsziel 1-Messkriterium':'',
         'Qualitätsziel 1-Zielbezeicnung':'',
         'Qualitätsziel 1-Zielbeschreibung':'',
         'Qualitätsziel 1-Messkriterium':'',
         'Sozialziel 1-Zielbezeicnung':'',
         'Sozialziel 1-Zielbeschreibung':'',
         'Sozialziel 1-Messkriterium':'',
         'Terminziel 1-Zielbezeicnung':'',
         'Terminziel 1-Zielbeschreibung':'',
         'Terminziel 1-Messkriterium':'',
         'Kostenziel 1-Zielbezeicnung':'',
         'Kostenziel 1-Zielbeschreibung':'',
         'Kostenziel 1-Messkriterium':'',
         'Aufwandsziel 1-Zielbezeicnung':'',
         'Aufwandsziel 1-Zielbeschreibung':'',
         'Aufwandsziel 1-Messkriterium':'',
         'Rahmenziel 1-Zielbezeicnung':'',
         'Rahmenziel 1-Zielbeschreibung':'',
         'Rahmenziel 1-Messkriterium':''}

#Can be expanded
D_Nichtziele={'Nichtziel 1':'',
              'Nichtziel 1-Beschreibung':''}

#Unlike D_Ziele, this Dictionary is not supposed to be expanded.
#If there are more Zielkonflikte, another Dictionary is to be created
D_Zielkonflikt_1={'konkurrierendes Ziel 1':'',
                  'konkurrierendes Ziel 2':'',
                  'Art des Zielkonfliktes':'',
                  'Priorität':'',
                  'Erklärung':'',
                  'Massnahmen':''}

#Anforderungen und Ziele: Projetrelevanz und Einschätzung
A_Z_PRE={'2 PRE':''}

#3. Qualität (Abnahmekriterien)
#All of the dictionaries in this chapter can be expanded but can also remain empty
D_Zielsystem={'Bezeichnung 1':'',
              'Beschreibung 1':'',
              'Abnahmekriterium 1':'',
              'Wer 1':'',
              'Zeitpunkt 1':''}

D_PMSystem={'Bezeichnung 1':'',
            'Beschreibung 1':'',
            'Abnahmekriterium 1':'',
            'Wer 1':'',
            'Zeitpunkt 1':''}

D_Compliance={'Bezeichnung 1':'',
              'Beschreibung 1':'',
              'Abnahmekriterium 1':'',
              'Wer 1':'',
              'Zeitpunkt 1':''}

D_Verträge={'Bezeichnung 1':'',
            'Beschreibung 1':'',
            'Abnahmekriterium 1':'',
            'Wer 1':'',
            'Zeitpunkt 1':''}

D_Projektträgerorganisation={'Bezeichnung 1':'',
                             'Beschreibung 1':'',
                             'Abnahmekriterium 1':'',
                             'Wer 1':'',
                             'Zeitpunkt 1':''}

D_Interessengruppen={'Bezeichnung 1':'',
                     'Beschreibung 1':'',
                     'Abnahmekriterium 1':'',
                     'Wer 1':'',
                     'Zeitpunkt 1':''}

D_sonstige={'Bezeichnung 1':'',
            'Beschreibung 1':'',
            'Abnahmekriterium 1':'',
            'Wer 1':'',
            'Zeitpunkt 1':''}

Q_PRE={'3 PRE':''}

#4. Stakeholder
D_Umfeldportfolio={'sozial intern':'',
                 'sozial extern': '',
                 'sachlich intern':'',
                 'sachlich extern':''}

#Stakeholder: Interessen, Erwartungen, Befrüchtungen, Massnahmen
#Is supposed to be expanded
D_Stakehoder_IEBM={'Stakeholder 1':'',
                   'Konfliktpotenzial 1':'',
                   'Einfluss 1':'',
                   'Interessen SH 1':'',
                   'Interessen Projekt 1':'',
                   'Strategie 1':'',
                   'Steuerung 1':''}

SH_PRE={'4 PRE':''}

#5. Chancen und Risiken
#At least 3 Risks, can be expanded
D_Risiken={'Risiko 1':'',
           'Beschreibung 1': '',
           'Art 1':'',
           'Ursache 1':'',
           'Risiko 2':'',
           'Beschreibung 2': '',
           'Art 2':'',
           'Ursache 2':'',
           'Risiko 3':'',
           'Beschreibung 3': '',
           'Art 3':'',
           'Ursache 3':''}

#Is dependant on the amount of Risks in D_Risiken
D_Massnahmen={'Eintrittsawhrscheinlichkeit 1':'',
              'Auswirkungen 1':'',
              'Risikowert 1':'',
              'Massnahmen p 1':'',
              'Massnahmen k 1':'',
              'Eintrittsawhrscheinlichkeit 2':'',
              'Auswirkungen 2':'',
              'Risikowert 2':'',
              'Massnahmen p 2':'',
              'Massnahmen k 2':'',
              'Eintrittsawhrscheinlichkeit 3':'',
              'Auswirkungen 3':'',
              'Risikowert 3':'',
              'Massnahmen p 3':'',
              'Massnahmen k 3':'',
              'Risikensumme':0}

#At least one, can be expanded. 
#Last entry to be used if the answer is in text instead of tables (2)
D_Chancen={'Chance 1':'',
           'Beschreibung 1':'',
           'Art 1':'',
           'Ursache 1':'',
           'Eintrittswahrscheilichkeit 1':'',
           'Auswirkungen 1':'',
           'Chancenwert 1':'',
           'Massnahmen 1':'',
           'Chancensumme':0,
           'Chancen Text':''}

C_R_PRE={'5 PRE':''}

past_block=''

#Prints all text blocks
for block in iter_block_items(doc):
    current_block=block.text
    # ignore repeated cells, #ignore paragraphs starting with ignore_list
    if current_block != past_block: #and not current_block.startswith(ignore_list[0]) and not current_block.startswith(ignore_list[1]) and not current_block.startswith(ignore_list[2]) and not current_block.startswith(ignore_list[3]) and not current_block.startswith(ignore_list[4])
        print(block.text)
        #print(block.style)
    past_block=current_block


#Saves the name of the project from the first table into a dictionary and prints it
past_block=''
for block in iter_block_items(doc):
    current_block=block
    # don't need to ignore repeated text from combined cells for this case
    if past_block!='' and past_block.text == 'Projekt:':
        D_Projektname['Projektname']=block.text
        #print(type(block))
    past_block=current_block
print(D_Projektname)


#Looks for paragraphs in Subchapter 1.1.1: 'Projektbeschreibung und -design'
# and saves them into a directory

# The text of the chapter will be saved into this variable
textvar=""
past_block=''
found_Chapter=False
for block in iter_block_items(doc):
    current_block=block
    # don't need to ignore repeated text from combined cells for this case
    if found_Chapter:
        if block.style.name=='Heading 3' or block.style.name=='Heading 1':
            break
        textvar+=block.text
    if past_block!='' and past_block.text.startswith('Projektbeschreibung und -design'):
        found_Chapter=True
    #print(block.style.name)
    past_block=current_block
D_Projektdesign['Projektbeschreibung und -design']=textvar


#Looks for paragraphs in Subchapter 1.1.2: 'Beschreibung des Projekterfolges aus Sicht der Kunden/des Auftraggebers'
# and saves them into a directory

# The text of the chapter will be saved into this variable
textvar=""
past_block=''
found_Chapter=False
for block in iter_block_items(doc):
    current_block=block
    # don't need to ignore repeated text from combined cells for this case
    if found_Chapter:
        if block.style.name=='Heading 3' or block.style.name=='Heading 1' or block.text=='Projektrelevanz und Einschätzung':
            break
        textvar+=block.text
    if past_block!='' and past_block.text.startswith('Beschreibung des Projekterfolges aus Sicht der Kunden/des Auftraggebers'):
        found_Chapter=True
    #print(block.style.name)
    past_block=current_block
D_Projektdesign['Beschreibung des Projekterfolges aus Sicht der Kunden/des Auftraggebers']=textvar


#Looks for paragraphs after 'Projektrelevanz und Einschätzung'
# and saves them into a directory

# The text of the chapter will be saved into this variable
textvar=""
past_block=''
found_Chapter=False
for block in iter_block_items(doc):
    current_block=block
    # don't need to ignore repeated text from combined cells for this case
    if found_Chapter:
        if block.style.name=='Heading 1':
            break
        textvar+=block.text
    if past_block!='' and past_block.text.startswith('Projektrelevanz und Einschätzung'):
        found_Chapter=True
    #print(block.style.name)
    past_block=current_block
D_Projektdesign['Projektrelevanz und Einschätzung']=textvar

print(D_Projektdesign)

#Looks for the contents of the Steckbrief table 
# and saves them into a dictionary

#textvar=""
#past_block=''
#found_Chapter=False
#i=0
#for item in D_Steckbrief:
#    D_Steckbrief[item]
#    i+=1
    