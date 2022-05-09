# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
"""
----------------------------------------------------------------------------------------------------------------------
Konvertiert CSV Dateien in PDF Dokumente

Autor: H. Füchtenhans

----------------------------------------------------------------------------------------------------------------------
"""

import configparser
import csv
import datetime
import email.mime.text
import email.mime.multipart
import email.mime.application
import email.header
import glob
import locale
import os
import smtplib
import ssl
import sys
import tempfile
from typing import Dict, List

import reportlab
import reportlab.pdfgen.canvas
from reportlab.lib.pagesizes import A4
import reportlab.platypus.paragraph


__version__ = "0.1.0"
STARTINFO = f"\nCSV_To_PDF       Version {__version__}"
HILFEMELDUNG = STARTINFO + ("\n\n"
                "Konvertiert CSV Dateien in PDF Dokumente\n"
                "Zum Aufbau der INI Datei, s. Example.ini\n"
                "Aufruf: csv_to_pdf IniDatei CSVDatei\n"
                "    IniDatei     Datei mit den Konvertierungsparametern\n"
                "    CSVDatei     Datei die in PDF Dokumente umgewandelt werden soll\n"
                "                 Es können hier auch * oder ? um Dateinamen verwendet\n"
                "                 werden. Damit werden dann alle entsprechenden Dateien\n"
                "                 mit der INI Datei verarbeitet."
                "Autor: Heribert Füchtenhans\n"
                "       Heribert.Fuechtenhans (at) yahoo.de")
DEBUG = False


# -----------------------------------------------------------------------------

class MyError(Exception):
    '''Exception Klasse für eigene Fehler'''


# -----------------------------------------------------------------------------

class IniDaten():
    """Klasse mit den aus der INI Datei gelesenen Werten und der Funktion um die INI Datei
    einzulesen"""

    def __init__(self) -> None:
        """Klasseninitialisierung"""
        self.gruppierungsspalte = ""
        self.spalten: list[str] = []
        self.ausgabeverzeichnis = ""
        self.mailgateway = ""
        self.mailgatewayport = 25
        self.mailsender = ""
        self.mailsenderpasswort = ""
        self.mailempfaenger: List[str] = []
        self.betreff = ""
        self.mailtext = ""
        self.sendmail = False


    def read_ini_datei(self, dateiname: str) -> None:
        """Liste die INI Datei ein.
        return: True wenn alles OK, false wenn eine Fehlermeldung ausgegeben wurde"""
        section = "Options"
        config = configparser.ConfigParser(interpolation=None)
        with open(dateiname, "rt", encoding="utf-8", errors='Replace') as infile:
            config.read_file(infile)
        self.gruppierungsspalte = config.get(section, "Gruppierungsspalte", fallback='')
        wert = config.get(section, "Spalten", fallback='')
        # Den Wert an den ; aufteilen, Leerzeichen entfernen und leer Spalten ignorieren
        self.spalten = [x.strip() for x in wert.split(",") if x.strip() != ""]
        self.ausgabeverzeichnis = config.get(section, "Ausgabeverzeichnis", fallback='')
        self.mailgateway = config.get(section, "Mailgateway", fallback='')
        self.mailgatewayport = config.getint(section, "MailgatewayPort", fallback=25)
        self.mailsender = config.get(section, "MailSender", fallback='')
        self.mailsenderpasswort = config.get(section, "MailSenderPasswort", fallback='')
        wert = config.get(section, "MailEmpfaenger", fallback='')
        # Den Wert an den ; aufteilen, Leerzeichen entfernen und leer Spalten ignorieren
        self.mailempfaenger = [x.strip().lower() for x in wert.split(",") if x.strip() != ""]
        self.betreff = config.get(section, "Betreff", fallback='CSV to PDF')
        self.mailtext = config.get(section, "Mailtext", fallback='FYI\n\nDie Mail wurde automatisch erzeugt, bitte nicht darauf antworten.\n')
        if DEBUG:
            print(f"Debug IniDaten:\n{vars(self)}")
        # Check die Daten
        if self.gruppierungsspalte == "":
            raise MyError("ERROR:\tDer Eintrag für die Gruppierungsspalte darf nicht leer sein.")
        if len(self.spalten) == 0:
            raise MyError("ERROR:\tEs wurden keine Spalten für die PDF Datei angegeben, Eintrag Spalten ist leer.")
        if self.ausgabeverzeichnis == "" and (self.mailgateway == "" or len(self.mailempfaenger) == 0 or self.mailsender == ""):
            raise MyError("ERROR:\tEs wurden weder ein Ausgabeverzeichnis noch ein Mailgateway / Mailempfaenger / Mailsender angegeben.")
        if self.mailgateway != "" and len(self.mailempfaenger) != 0 and self.mailsender != "":
            self.sendmail = True


# -----------------------------------------------------------------------------

def per_mail_versenden(inidaten: IniDaten, filename: str) -> None:
    """Versendet die Datei per Mail"""
    msg = email.mime.multipart.MIMEMultipart()
    msg['From'] = email.header.Header(inidaten.mailsender)
    msg['To'] = email.header.Header(";".join(inidaten.mailempfaenger))
    msg['Subject'] = email.header.Header(inidaten.betreff)
    msg.attach(email.mime.text.MIMEText(inidaten.mailtext, 'plain', 'utf-8'))
    # PDF anhängen
    att_name = os.path.basename(filename)
    _f = open(filename, 'rb')
    att = email.mime.application.MIMEApplication(_f.read(), _subtype="pdf")
    _f.close()
    att.add_header('Content-Disposition', 'attachment', filename=att_name)
    if DEBUG:
        print(f"Debug Maildaten:\n{msg}")
    msg.attach(att)
    # Mail versenden
    server = smtplib.SMTP(inidaten.mailgateway, inidaten.mailgatewayport)
    if inidaten.mailsenderpasswort != "":
        server.starttls(context=ssl.create_default_context())
        server.login(inidaten.mailsender, inidaten.mailsenderpasswort)
    # send email and quit server
    server.sendmail(inidaten.mailsender, ";".join(inidaten.mailempfaenger), msg.as_string())
    server.quit()


def erstelle_pdf(inidaten: IniDaten, datenliste: List[Dict[str, str]], csv_datei: str, gruppierungswert: str) -> None:
    """Erstellt die PDF Datei"""
    # initializing variables with values
    filename = os.path.split(csv_datei)[1]
    filename = os.path.splitext(filename)[0]
    filename = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{gruppierungswert}_{filename}.pdf"
    if inidaten.ausgabeverzeichnis != "":
        filename = os.path.join(inidaten.ausgabeverzeichnis, filename)
    else:
        filename = os.path.join(tempfile.gettempdir(), filename)
    # creating a pdf object
    pdf = reportlab.pdfgen.canvas.Canvas(filename, pagesize=A4)
    width, height = A4 #keep for later
    # setting the title of the document
    pdf.setTitle(gruppierungswert)
    # creating the title by setting it's font
    # and putting it on the canvas
    pdf.setFont("Courier-Bold", 16)
    pdf.drawString(30, height-30, f"{inidaten.gruppierungsspalte}: {gruppierungswert}")
    # drawing a line
    pdf.line(30, height-60, width-30, height-60)
    # creating a multiline text using
    pdf.setFont("Helvetica", 12)
    # Bestimme den breitest Eintrag einer Spaltenüberschrift
    max_spalten_laenge = max([reportlab.platypus.paragraph.stringWidth(key, "Helvetica", 12) for key in inidaten.spalten])
    ypos = height - 75
    for entry in datenliste:
        for key in inidaten.spalten:
            if key in entry:
                pdf.drawString(30, ypos, key)
                pdf.drawString(30 + max_spalten_laenge + 10, ypos, f": {entry[key]}")
                ypos -= 16
                if ypos < 20:
                    pdf.showPage()
                    ypos = height -60
                    pdf.setFont("Helvetica", 12)
        ypos -= 32
        if ypos < 20:
            pdf.showPage()
            ypos = height -60
            pdf.setFont("Helvetica", 12)
    pdf.save()
    if inidaten.sendmail:
        per_mail_versenden(inidaten, filename)
    if inidaten.ausgabeverzeichnis == "":
        os.remove(filename)


def daten_bearbeiten(inidaten: IniDaten, csv_datei: str) -> None:
    """Liest die CSV Datei ein und erstellt die PDF Dateien"""
    with open(csv_datei, mode ='rt', encoding="Ansi") as file:
        csv_file = csv.DictReader(file, delimiter=';')
        zeilennr = 0
        gruppierungswert = ""
        datenliste: List[Dict[str, str]] = []
        # Werte anzeigen
        for zeile in csv_file:
            zeilennr += 1
            if zeilennr == 1:
                # Test ob die Gruppierungsspalte vorhanden ist, wenn ja Inhalt merken
                if not inidaten.gruppierungsspalte in zeile:
                    raise MyError(f"ERROR:\tDie Gruppierungsspalte '{inidaten.gruppierungsspalte}' wurde nicht in der CSV Datei gefunden.")
                gruppierungswert = zeile[inidaten.gruppierungsspalte]
            if gruppierungswert != zeile[inidaten.gruppierungsspalte] and len(datenliste) != 0:
                erstelle_pdf(inidaten, datenliste, csv_datei, gruppierungswert)
                gruppierungswert = zeile[inidaten.gruppierungsspalte]
                datenliste = []
            datenzeile: Dict[str, str] = {}
            for key in inidaten.spalten:
                if key in zeile:
                    datenzeile[key] = zeile[key]
            datenliste.append(datenzeile)
        if len(datenliste) != 0:
            erstelle_pdf(inidaten, datenliste, csv_datei, gruppierungswert)


# -----------------------------------------------------------------------------

def main(argv: List[str]) -> int:    # pylint: disable=too-many-branches, too-many-statements
    """Hauptprogramm"""
    locale.setlocale(locale.LC_ALL, '')
    # Startparameter auswerten
    if len(argv) < 3:
        print(HILFEMELDUNG)
        return 1
    ini_datei = argv[1]
    csv_datei = argv[2]
    if not os.path.exists(ini_datei):
        print(f"ERROR:\tIniDatei '{ini_datei}' nicht gefunden.")
        return 2
    print(f"INFO:\tVerwende INI Datei {ini_datei}")
    try:
        # Daten der INI Datei einlesen
        inidaten = IniDaten()
        inidaten.read_ini_datei(ini_datei)
        for dateiname in glob.glob(csv_datei):
            print(f"INFO:\tBearbeite: {dateiname}")
            daten_bearbeiten(inidaten, dateiname)
    except MyError as err:
        print(err)
    return 0


# ----------------------------------------------------------------------------
if __name__ == '__main__':
    sys.exit(main(sys.argv))
