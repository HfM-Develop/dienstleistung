import os

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageTemplate, BaseDocTemplate, Image, \
    Preformatted, NextPageTemplate, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus.frames import Frame
from reportlab.pdfbase.ttfonts import TTFont
import datetime
from datetime import datetime

# Importiere die Schriftart und registriere sie
pdfmetrics.registerFont(TTFont('Aeonik-Air', 'Aeonik-Air.ttf'))
pdfmetrics.registerFont(TTFont('Aeonik-Bold', 'Aeonik-Bold.ttf'))
pdfmetrics.registerFont(TTFont('Aeonik-Light', 'Aeonik-Light.ttf'))


class MyDocTemplate(BaseDocTemplate):
    def __init__(self, filename="", **kwargs):
        BaseDocTemplate.__init__(self, filename, **kwargs)
        self.table_last = None
        self.table_footer = None
        self.table_content = None
        self.table_header = None
        self.table = None
        self.my_style = None
        self.addPageTemplates(self.create_page_template())

    def create_page_template(self):
        frame = Frame(
            30, 30, 530, 650,
            id='normal'
        )
        template = PageTemplate(
            id='test',
            frames=[frame],
            onPage=self.add_other_stuff
        )

        self.table_header = ParagraphStyle(
            name='TableHeader',
            fontName='Aeonik-Bold',
            fontSize=9,
            textColor='black',
            spaceBefore=10,
            spaceAfter=0,
            alignment=0)

        self.table_content = ParagraphStyle(
            name='TableContent',
            fontName='Aeonik-Light',
            fontSize=9,
            textColor='black',
            spaceBefore=0,
            spaceAfter=0,
            alignment=0)
        # leading=10)

        self.table_footer = ParagraphStyle(
            name='TableFooter',
            fontName='Aeonik-Bold',
            fontSize=9,
            textColor='black',
            spaceBefore=0,
            spaceAfter=0,
            alignment=0,
            leading=10,
            spacing=10)

        self.table_last = ParagraphStyle(
            name='TableLast',
            fontName='Aeonik-Bold',
            fontSize=9,
            textColor='black',
            spaceBefore=0,
            spaceAfter=0,
            alignment=0)

        return template

    def add_other_stuff(self, canvas, doc):
        # Übertragung der Werte aus MyAgrar
        customer = getattr(doc, 'customer', '')  # Abrufen des Werts der customer-Variable
        customer_adress = getattr(doc, 'customer_adress', '')  # Abrufen des Werts der customer-Variable
        customer_code = getattr(doc, 'customer_code', '')  # Abrufen des Werts der customer-Variable
        customer_domicile = getattr(doc, 'customer_domicile', '')  # Abrufen des Werts der customer-Variable
        servicetype = getattr(doc, 'servicetype', '')  # Abrufen des Werts der customer-Variable
        description = getattr(doc, 'description', '')  # Abrufen des Werts der customer-Variable
        start = getattr(doc, 'start', '')  # Abrufen des Werts der customer-Variable
        ende = getattr(doc, 'ende', '')  # Abrufen des Werts der customer-Variable
        number = getattr(doc, 'number', '')

        # Datum drehen
        try:
            start_old = datetime.strptime(str(start), '%Y-%m-%d')
            start_new = start_old.strftime('%d.%m.%Y')
            start = start_new
        except:
            pass

        try:
            ende_old = datetime.strptime(str(ende), '%Y-%m-%d')
            ende_new = ende_old.strftime('%d.%m.%Y')
            ende = ende_new
        except:
            pass

        # TopPdf Adress Zeile-----------------------------------------------------------
        header_text = "VR Agrarberatung AG, Rheiner Straße 127, 49809 Lingen"

        TopHeaderStyle = ParagraphStyle(
            name='HeaderStyle',
            parent=getSampleStyleSheet()['Heading1'],
            fontSize=9,  # Schriftgröße anpassen
            fontName='Aeonik-Light',  # Schriftart anpassen
            textColor='#717357',
        )
        header_paragraph = Preformatted(header_text, TopHeaderStyle)
        w, h = header_paragraph.wrap(doc.width, doc.topMargin)
        header_paragraph.drawOn(canvas, 30, 800)
        # Header Ende-----------------------------------------------------

        # Anschrift Kunde-----------------------------------------------------------------
        adress_text = "{}\n" \
                      "{}\n" \
                      "{} {}".format(customer, customer_adress, customer_code, customer_domicile)

        AdressStyle = ParagraphStyle(
            name='AdressStyle',
            parent=getSampleStyleSheet()['Heading1'],
            fontSize=11,  # Schriftgröße anpassen
            fontName='Aeonik-Light',  # Schriftart anpassen
            textColor='black',
            leading=13
        )
        adress_paragraph = Preformatted(adress_text, AdressStyle)
        w, h = adress_paragraph.wrap(doc.width, doc.topMargin)
        adress_paragraph.drawOn(canvas, 30, 760)
        # Anschrift Ende-----------------------------------------------------------------

        # Kontakt-----------------------------------------------------------
        contact_text = "Dienstleistungsnachweis {} Nr. {}\n" \
                       "für {},\n" \
                       "vom {} bis {}".format(servicetype, number, description, start, ende)

        contact_style = ParagraphStyle(
            name='ContactStyle',
            parent=getSampleStyleSheet()['Heading1'],
            fontSize=12,  # Schriftgröße anpassen
            fontName='Aeonik-Light',
            textColor='#717357',  # Schriftart anpassen
            leading=13
        )

        kontakt_paragraph = Preformatted(contact_text, contact_style)
        w, h = kontakt_paragraph.wrap(doc.width, doc.topMargin)
        kontakt_paragraph.drawOn(canvas, 30, 700)
        # Kontakt Ende-----------------------------------------------------

        # Logo
        logo_path = "LOGO_weißer_Hintergrund.png"  # Pfad zum Logo-Bild
        logo = Image(logo_path, width=100, height=15)
        logo.drawOn(canvas, 440, 810)

        # Footer
        footer_table = self.create_footer_table()
        footer_table.wrap(doc.width, doc.bottomMargin)
        footer_table.drawOn(canvas, 30, 30)

    def create_footer_table(self):
        data = [
            ["VR Agrarberatung", "Emsländische Volksbank eG", "Vorstand:        Sven Foppe"],
            ["Rheiner Strasse 127", "DE03 266600601100171200", "Aufsichtsrat:    Nobert Focks"],
            ["49809 Lingen", "BIC GENODEF1LIG", ""],
            ["Telefon        0591 804400", "", "UST.-ID-Nr. DE180925319"],
            ["Mail              info@vr-agrar.de", "Amtsgericht Osnabrück", "Steuer-Nr.   61/201/13138"],
            ["Web             www.vr-agrar.de", "HRB 101212", ""],
        ]

        table = Table(data, colWidths=[210, 210, 180])
        table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONT', (0, 0), (-1, -1), 'Aeonik-Light', 6),
        ]))

        return table

    def create_pdf(self, pdflist, customer, customer_adress, customer_code, customer_domicile, servicetype, description,
                   start, ende, number):
        # Daten für die Tabelle
        self.pdflist = pdflist

        # Definieren des Inhalts des PDF-Dokuments
        elements = []
        table_data = []
        mengensumme_alles = 0
        preissumme_alles = 0

        for service, positions in self.pdflist.items():
            positions = list(positions)
            # service enthält die Daten für eine Tabelle
            # Erstellen einer leeren Tabelle für die aktuellen Dienstleistungsdaten

            # Hinzufügen der Kopfzeile zur Tabelle mit eigenem Style (header_style)
            table_header = [service, "Bemerkung", "Datum", "Menge", "Einzelpreis [€]", "Gesamt [€]", "Berater"]
            table_header_paragraph = [Paragraph(cell, style=self.table_header) for cell in table_header]
            table_data.append(table_header_paragraph)

            mengensumme = 0
            preissumme = 0
            counter = 0
            pagebreake = 0

            for row in positions:
                new_list = []
                row = list(map(str, row))

                if service != "Reisekosten":
                    mengensumme = mengensumme + float(row[3])
                    preissumme = preissumme + float(row[5])
                else:
                    mengensumme = 0
                    preissumme = preissumme + float(row[5])

                mengensumme_formatted = "{:.2f}".format(mengensumme)
                preissumme_formatted = "{:.2f}".format(preissumme)

                for i in row:
                    new_list.append(str(i))
                table_row = []
                for cell in new_list:
                    paragraph = Paragraph(cell, style=self.table_content)
                    table_row.append(paragraph)
                table_data.append(table_row)
                counter += 1
                if counter >= 15:
                    pagebreake += 1

            if service != "Reisekosten":
                footer_row = ["Zwischenumme", "", "", str(mengensumme_formatted), "Stunden", str(preissumme_formatted),
                              ""]
                footer_row_paragraph = [Paragraph(cell, style=self.table_footer) for cell in footer_row]
                table_data.append(footer_row_paragraph)
                counter += 1
            else:
                footer_row = ["Zwischensumme", "", "", "", "", str(preissumme_formatted), ""]
                footer_row_paragraph = [Paragraph(cell, style=self.table_footer) for cell in footer_row]
                table_data.append(footer_row_paragraph)
                counter += 1

            mengensumme_alles = mengensumme_alles + mengensumme
            preissumme_alles = preissumme_alles + preissumme

            self.table = Table(table_data, colWidths=[132, 100, 65, 50, 75, 65, 65])
            self.table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ]))
            if counter >= 8:
                elements.append(self.table)
                elements.append(NextPageTemplate('test'))
                elements.append(PageBreak())
                counter = 0
            else:
                elements.append(self.table)
            self.table = ""
            table_data = []
            mengensumme = 0
            preissumme = 0

        mengensumme_alles_form = "{:.2f}".format(mengensumme_alles)
        preissumme_alles_form = "{:.2f}".format(preissumme_alles)

        table_data.append([""])
        last_row = ["Gesamt", "", "", str(mengensumme_alles_form), "Stunden", str(preissumme_alles_form), ""]
        last_row_paragraph = [Paragraph(cell, style=self.table_last) for cell in last_row]
        table_data.append(last_row_paragraph)

        self.table = Table(table_data, colWidths=[132, 100, 65, 50, 75, 65, 65])
        elements.append(self.table)

        # Setzen Sie den Pfad zum gewünschten Verzeichnis
        windows_username = os.environ.get('USERNAME')
        save_directory = "C:\\Users\\" + windows_username + "\\VR AgrarBeratung AG\\genoBIT Administrator - - VRAgrar Daten\\308 Projekte\\1_HfM\\Dienstleistung\\Archiv\\"

        # Erstellen des PDF-Dokuments
        self.filename = f"DLN für {customer} {datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.pdf"
        self.doc = MyDocTemplate(self.filename, pagesize=A4)

        # Werte übergeben
        self.doc.customer = customer  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.customer_adress = customer_adress  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.customer_code = customer_code  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.customer_domicile = customer_domicile  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.servicetype = servicetype  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.description = description  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.start = start  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.ende = ende  # Zuweisen der customer-Variable an das DocTemplate-Objekt
        self.doc.number = number

        # Generieren des PDFs
        self.doc.build(elements)

        # Öffnen des generierten PDFs
        os.startfile(self.filename)
