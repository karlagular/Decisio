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

#2. Anforderungen und Ziele
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

D_Nichtziele={'Nichtziel 1':'',
              'Nichtziel 1-Beschreibung':''}

D_Zielkonflikt_1={'konkurrierendes Ziel 1':'',
                  'konkurrierendes Ziel 2':'',
                  'Art des Zielkonfliktes':'',
                  'Priorität':'',
                  'Erklärung':'',
                  'Massnahmen':''}

#Anforderungen und Ziele: Projetrelevanz und Einschätzung
A_Z_PRE={'2 PRE':''}

#3. Qualität (Abnahmekriterien)
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

D_Stakehoder_IEBM={'Stakeholder 1':'',
                   'Konfliktpotenzial 1':'',
                   'Einfluss 1':'',
                   'Interessen SH 1':'',
                   'Interessen Projekt 1':'',
                   'Strategie 1':'',
                   'Steuerung 1':''}

SH_PRE={'4 PRE':''}

#5. Chancen und Risiken
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
              'Massnahmen k 3':'',}

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
for block in iter_block_items(doc):
    current_block=block.text
    # ignore repeated cells, #ignore paragraphs starting with ignore_list
    if current_block != past_block: #and not current_block.startswith(ignore_list[0]) and not current_block.startswith(ignore_list[1]) and not current_block.startswith(ignore_list[2]) and not current_block.startswith(ignore_list[3]) and not current_block.startswith(ignore_list[4])
        print(block.text)
        print(block.style)
    past_block=current_block

#print(D_Projektname)
