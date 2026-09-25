"""Builds "Wohnungsbewerbung Schweiz" and "Bewerbungsvorlagen Schweiz" (Word).
Run: python products/dossier-schweiz/build.py
"""

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt, RGBColor

HERE = Path(__file__).parent
DIST = HERE / "dist"
sys.path.insert(0, str(HERE.parent.parent / "tools"))
from cover import make_cover  # noqa: E402

RED = RGBColor(0xD5, 0x2B, 0x1E)


def new_doc(title, intro):
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading(title, 0)
    doc.add_paragraph(intro)
    return doc


def page_break(doc):
    doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)


def tip(doc, text):
    r = doc.add_paragraph().add_run("Tipp: " + text)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RED


def checklist(doc, items):
    for item in items:
        doc.add_paragraph("☐  " + item)


def table(doc, rows, widths=None):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Table Grid"
    for i, (a, b) in enumerate(rows):
        t.cell(i, 0).text = a
        t.cell(i, 1).text = b
        t.cell(i, 0).paragraphs[0].runs[0].bold = True


def build_wohnung(path):
    doc = new_doc("Wohnungsbewerbung Schweiz",
                  "Vorlagen und Checkliste für ein vollständiges Bewerbungsdossier – damit deine Bewerbung "
                  "bei der Verwaltung oben auf dem Stapel landet. Für Word, Pages und Google Docs.")
    doc.add_heading("1. Checkliste: Das gehört ins Dossier", 1)
    checklist(doc, [
        "Anmeldeformular der Verwaltung – vollständig ausgefüllt und unterschrieben",
        "Betreibungsregisterauszug (aktuell, meist nicht älter als 3 Monate) – bei der Wohngemeinde oder online",
        "Kopie Ausweis bzw. Ausländerausweis (Aufenthaltsbewilligung) aller Mietenden",
        "Lohnnachweis: die letzten 3 Lohnabrechnungen oder Arbeitsvertrag (bei Stellenantritt)",
        "Referenz der aktuellen Verwaltung / des aktuellen Vermieters (Name, Telefon)",
        "Kurzes Anschreiben (Vorlage unten) und Mieter-Selbstauskunft (Vorlage unten)",
        "Optional: Nachweis Privathaftpflichtversicherung, Foto von dir / euch",
    ])
    tip(doc, "Faustregel vieler Verwaltungen: Die Bruttomiete sollte höchstens rund ein Drittel des "
             "Bruttoeinkommens betragen. Alles als EIN PDF in der Reihenfolge der Checkliste senden – "
             "Dateiname z. B. «Bewerbung_Muster_Seestrasse12.pdf».")

    page_break(doc)
    doc.add_heading("2. Anschreiben", 1)
    doc.add_paragraph("[Vorname Name]\n[Strasse Nr.]\n[PLZ Ort]\n[Telefon] · [E-Mail]")
    doc.add_paragraph("[Verwaltung]\n[z. Hd. Name]\n[Strasse Nr.]\n[PLZ Ort]")
    doc.add_paragraph("[Ort], [Datum]")
    doc.add_paragraph("Bewerbung für die [Anzahl]-Zimmer-Wohnung an der [Adresse], Objekt-Nr. [Nummer]").runs[0].bold = True
    for para in (
        "Sehr geehrte/r [Frau/Herr Name]",
        "Vielen Dank für die Besichtigung vom [Datum]. Die Wohnung gefällt mir / uns sehr, und ich bewerbe mich "
        "hiermit gerne dafür. Gewünschter Mietbeginn: [Datum].",
        "Zu mir / uns: Ich arbeite seit [Jahr] unbefristet als [Beruf] bei [Arbeitgeber]. [Optional: Partner/in, "
        "Kinder, Haustiere.] Wir sind Nichtraucher, ruhig und pflegen unsere Wohnung sorgfältig – gerne bestätigt "
        "Ihnen dies unsere aktuelle Verwaltung [Name, Telefon].",
        "Warum diese Wohnung: [z. B. Nähe zum Arbeitsplatz, Schule der Kinder, langfristige Perspektive].",
        "Alle Unterlagen finden Sie im beiliegenden Dossier. Für Fragen bin ich jederzeit unter [Telefon] erreichbar.",
        "Freundliche Grüsse\n\n\n[Vorname Name]",
    ):
        doc.add_paragraph(para)

    page_break(doc)
    doc.add_heading("3. Mieter-Selbstauskunft", 1)
    doc.add_paragraph("Für jede mietende Person ausfüllen.")
    table(doc, [
        ("Name, Vorname", ""), ("Geburtsdatum", ""), ("Nationalität / Bewilligung", ""),
        ("Zivilstand", ""), ("Aktuelle Adresse", ""), ("Telefon / E-Mail", ""),
        ("Beruf / Arbeitgeber", ""), ("Angestellt seit / unbefristet?", ""),
        ("Bruttoeinkommen pro Jahr (CHF)", ""), ("Aktuelle Verwaltung (Referenz)", ""),
        ("Grund für den Wohnungswechsel", ""), ("Anzahl Personen / Kinder", ""),
        ("Haustiere", ""), ("Raucher", ""), ("Musikinstrumente", ""),
        ("Fahrzeug / Parkplatz benötigt", ""), ("Privathaftpflicht vorhanden", ""),
        ("Betreibungen in den letzten Jahren", ""),
    ])
    doc.add_paragraph("\nIch bestätige, dass die Angaben vollständig und wahr sind.\n\n[Ort, Datum]            "
                      "[Unterschrift]")
    doc.save(path)


def build_bewerbung(path):
    doc = new_doc("Bewerbungsvorlagen Schweiz",
                  "Lebenslauf im Schweizer Stil, Motivationsschreiben und Dossier-Checkliste. "
                  "Für Word, Pages und Google Docs – alle [Platzhalter] ersetzen und Tipps löschen.")
    doc.add_heading("1. Checkliste: Vollständiges Bewerbungsdossier", 1)
    checklist(doc, [
        "Motivationsschreiben (max. 1 Seite)",
        "Lebenslauf (1–2 Seiten), in der Schweiz meist mit Foto",
        "Arbeitszeugnisse aller bisherigen Stellen (bei Stellenwechsel ggf. Zwischenzeugnis)",
        "Diplome und Abschlusszeugnisse, Weiterbildungszertifikate",
        "Referenzen (Name, Funktion, Telefon – vorher um Erlaubnis fragen)",
        "Bei ausländischen Diplomen: Anerkennung oder Einstufung, falls vorhanden",
        "Alles als EIN PDF: Motivationsschreiben, Lebenslauf, Zeugnisse, Diplome",
    ])

    page_break(doc)
    doc.add_heading("2. Lebenslauf", 1)
    tip(doc, "Schweizer Arbeitgeber erwarten meist ein professionelles Foto oben rechts. Umgekehrt "
             "chronologisch: die neueste Stelle zuerst. Lücken kurz erklären (Reise, Weiterbildung, Familie).")
    doc.add_heading("Persönliche Angaben", 2)
    table(doc, [
        ("Name", "[Vorname Name]"), ("Adresse", "[Strasse Nr., PLZ Ort]"),
        ("Telefon / E-Mail", "[+41 …] · [E-Mail]"), ("Geburtsdatum", "[TT.MM.JJJJ]"),
        ("Nationalität", "[Land] – [Bewilligung B/C/G, falls nicht Schweizer/in]"),
        ("LinkedIn", "[optional]"),
    ])
    doc.add_heading("Berufserfahrung", 2)
    for _ in range(3):
        doc.add_paragraph("[MM.JJJJ] – [MM.JJJJ / heute]    [Funktion], [Firma], [Ort]").runs[0].bold = True
        doc.add_paragraph("[Wichtigste Aufgabe oder messbarer Erfolg, z. B. «Kosten um 15 % gesenkt»]", style="List Bullet")
        doc.add_paragraph("[Aufgabe / Erfolg]", style="List Bullet")
    doc.add_heading("Aus- und Weiterbildung", 2)
    for _ in range(2):
        doc.add_paragraph("[JJJJ] – [JJJJ]    [Abschluss, z. B. EFZ / Bachelor / CAS], [Schule], [Ort]")
    doc.add_heading("Sprachen", 2)
    doc.add_paragraph("Deutsch: [Muttersprache / C1]   ·   Französisch: [B2]   ·   Englisch: [C1]")
    doc.add_heading("IT-Kenntnisse", 2)
    doc.add_paragraph("[z. B. MS Office (sehr gut), SAP (gut), …]")
    doc.add_heading("Interessen", 2)
    doc.add_paragraph("[2–3 Punkte, z. B. Vereinsarbeit, Sport, Ehrenamt]")
    doc.add_heading("Referenzen", 2)
    doc.add_paragraph("[Name, Funktion, Firma, Telefon] – oder: «Referenzen auf Anfrage»")

    page_break(doc)
    doc.add_heading("3. Motivationsschreiben", 1)
    tip(doc, "Max. eine Seite. Nicht den Lebenslauf wiederholen, sondern zeigen, was du der Firma bringst. "
             "Stelleninserat nennen und 2–3 Anforderungen konkret mit Beispielen belegen.")
    doc.add_paragraph("[Vorname Name]\n[Strasse Nr.]\n[PLZ Ort]\n[Telefon] · [E-Mail]")
    doc.add_paragraph("[Firma]\n[z. Hd. Name]\n[Strasse Nr.]\n[PLZ Ort]")
    doc.add_paragraph("[Ort], [Datum]")
    doc.add_paragraph("Bewerbung als [Funktion] – Ihr Inserat auf [Plattform] vom [Datum]").runs[0].bold = True
    for para in (
        "Sehr geehrte/r [Frau/Herr Name]",
        "[Einstieg: Warum genau diese Stelle und diese Firma? Ein konkreter Bezug, z. B. ein Projekt, ein "
        "Produkt oder ein Wert der Firma.]",
        "[Hauptteil: 2–3 Anforderungen aus dem Inserat, jeweils mit einem konkreten Beispiel aus deiner "
        "Erfahrung belegen – idealerweise mit Zahlen.]",
        "[Was du mitbringst: Arbeitsweise, Sprachen, Verfügbarkeit / Kündigungsfrist, ggf. Bewilligung.]",
        "Gerne überzeuge ich Sie in einem persönlichen Gespräch. Ich freue mich auf Ihre Rückmeldung.",
        "Freundliche Grüsse\n\n\n[Vorname Name]\n\nBeilagen: Lebenslauf, Arbeitszeugnisse, Diplome",
    ):
        doc.add_paragraph(para)
    doc.save(path)


if __name__ == "__main__":
    DIST.mkdir(exist_ok=True)
    build_wohnung(DIST / "Wohnungsbewerbung-Schweiz.docx")
    build_bewerbung(DIST / "Bewerbungsvorlagen-Schweiz.docx")
    make_cover(["Wohnungs-", "bewerbung Schweiz"], "Vorlagen & Dossier-Checkliste · Word",
               ["Checkliste: alles, was Verwaltungen verlangen",
                "Anschreiben-Vorlage, die überzeugt",
                "Mieter-Selbstauskunft zum Ausfüllen",
                "Tipps für Zürich, Genf, Basel & Co.",
                "Sofort-Download · Word & Google Docs"], DIST / "cover-wohnung.png")
    make_cover(["Bewerbungs-", "vorlagen Schweiz"], "Lebenslauf & Motivationsschreiben · Word",
               ["Lebenslauf im Schweizer Stil (mit Fotoplatz)",
                "Motivationsschreiben mit Aufbau-Anleitung",
                "Checkliste für das komplette Dossier",
                "Tipps zu Zeugnissen, Referenzen, Bewilligung",
                "Sofort-Download · Word & Google Docs"], DIST / "cover-bewerbung.png")
    print("OK")
