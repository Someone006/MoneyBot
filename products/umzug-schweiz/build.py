"""Builds "Umzugs-Planer Schweiz" / "Moving in Switzerland Planner" (DE + EN).

Enter the moving date once; every task gets its due date, overdue open tasks
turn red, and a progress bar tracks what is done.
Run: python products/umzug-schweiz/build.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

DIST = Path(__file__).parent / "dist"
RED, DARK, INPUT, PALE_RED = "D52B1E", "1F2937", "FFF7D6", "FDECEA"
thin = Side(style="thin", color="D1D5DB")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

# (days relative to moving day, German task, English task)
TASKS = [
    (-90, "Alte Wohnung kündigen – per Einschreiben, Kündigungsfrist und -termine im Mietvertrag prüfen (oft 3 Monate). Wohnt ihr als Ehepaar/eingetragene Partner dort, müssen beide unterschreiben.",
     "Give notice on your old flat – by registered letter; check notice period and termination dates in your lease (often 3 months). Married couples/registered partners living there must both sign."),
    (-75, "Neue Wohnung: Mietvertrag unterschreiben, Mietkaution organisieren (max. 3 Monatsmieten, Sperrkonto oder Kautionsversicherung).",
     "New flat: sign the lease and arrange the rental deposit (max. 3 months' rent, blocked deposit account or deposit insurance)."),
    (-60, "Umzugsfirma oder Transporter reservieren, mehrere Offerten vergleichen.",
     "Book a moving company or van – compare several quotes."),
    (-60, "Arbeitgeber informieren, Umzugstag frei nehmen.",
     "Tell your employer and take the moving day off."),
    (-45, "Internet/TV für die neue Adresse bestellen oder umziehen (Aufschaltung kann dauern).",
     "Order or transfer internet/TV for the new address (activation can take time)."),
    (-45, "Endreinigung mit Abnahmegarantie buchen – oder selbst planen.",
     "Book end-of-tenancy cleaning with handover guarantee – or plan to do it yourself."),
    (-30, "Nachsendeauftrag bei der Post erfassen.",
     "Set up mail forwarding with Swiss Post."),
    (-30, "Adresse melden: Bank, Krankenkasse (Prämienregion kann sich ändern!), Versicherungen, Arbeitgeber, Abos, Online-Shops.",
     "Update your address: bank, health insurer (premium region may change!), insurances, employer, subscriptions, online shops."),
    (-30, "Hausratversicherung: neue Adresse und Versicherungssumme anpassen.",
     "Household contents insurance: update address and sum insured."),
    (-21, "Aussortieren, verkaufen, verschenken; Sperrgut-Abfuhr organisieren.",
     "Declutter, sell or give away; arrange bulky waste pick-up."),
    (-14, "Parkbewilligung für den Umzugswagen bei der Gemeinde/Stadt beantragen (falls nötig).",
     "Apply for a parking permit for the moving van at the commune/city (if needed)."),
    (-14, "Wohnungsabgabe-Termin mit der Verwaltung vereinbaren; Mängel und Reparaturen erledigen.",
     "Agree the flat handover date with the landlord/agency; fix damages and repairs."),
    (-7, "Kartons beschriften, Umzugskiste mit Wichtigem (Dokumente, Ladekabel, Medikamente) packen.",
     "Label boxes; pack an essentials box (documents, chargers, medication)."),
    (-3, "Schlüsselübergabe der neuen Wohnung bestätigen.",
     "Confirm key handover for the new flat."),
    (0, "Zählerstände (Strom, ggf. Gas/Wasser) in alter und neuer Wohnung fotografieren.",
     "Photograph meter readings (electricity, gas/water if any) in old and new flat."),
    (0, "Neue Wohnung: Übergabeprotokoll prüfen, vorhandene Mängel sofort schriftlich melden.",
     "New flat: check the handover report and report existing defects in writing straight away."),
    (1, "Alte Wohnung: Abgabe mit Protokoll, alle Schlüssel zurückgeben.",
     "Old flat: handover with report, return all keys."),
    (7, "Abmeldung bei der alten Gemeinde (in vielen Kantonen online via eUmzugCH).",
     "Deregister at your old commune (online via eUmzugCH in many cantons)."),
    (14, "Anmeldung bei der neuen Gemeinde (Einwohnerkontrolle) – innerhalb von 14 Tagen! Ausweis, Mietvertrag, ggf. Familienausweis mitnehmen.",
     "Register at your new commune (residents' office) – within 14 days! Bring ID, lease and family documents if applicable."),
    (14, "Auto/Motorrad: Adressänderung beim Strassenverkehrsamt melden (Frist 14 Tage).",
     "Car/motorbike: report the address change to the road traffic office (14-day deadline)."),
    (14, "Hund: neue Adresse in der Heimtierdatenbank (AMICUS) erfassen.",
     "Dog: update the address in the pet database (AMICUS)."),
    (30, "Namensschild an Briefkasten und Tür anbringen.",
     "Put your name on the mailbox and door."),
    (30, "Mietkaution der alten Wohnung zurückfordern (nach Abgabe ohne offene Forderungen).",
     "Claim back the deposit for the old flat (after handover with no open claims)."),
    (60, "Kontrolle: Kommt die Post richtig an? Rechnungen mit neuer Adresse?",
     "Check: is mail arriving correctly? Bills showing the new address?"),
]
# Only when moving to Switzerland from abroad
EXPAT_TASKS = [
    (14, "Anmeldung bei der Gemeinde vor Arbeitsbeginn; Aufenthaltsbewilligung beantragen.",
     "Register at the commune before starting work; apply for your residence permit."),
    (30, "Schweizer Bankkonto eröffnen.",
     "Open a Swiss bank account."),
    (30, "Handy-Abo und Internet abschliessen.",
     "Get a Swiss mobile plan and internet."),
    (60, "Halbtax oder GA prüfen.",
     "Consider a Half Fare Travelcard or GA."),
    (90, "Krankenversicherung (Grundversicherung) abschliessen – Pflicht innerhalb von 3 Monaten nach Anmeldung.",
     "Take out basic health insurance – mandatory within 3 months of registering."),
    (365, "Ausländischen Führerausweis umtauschen – innerhalb von 12 Monaten.",
     "Exchange your foreign driving licence – within 12 months."),
]

TEXT = {
    "de": {
        "file": "Umzugs-Planer-Schweiz.xlsx",
        "sheet": "Umzugs-Planer",
        "title": "Umzugs-Planer Schweiz",
        "sub": "Umzugsdatum eintragen – alle Fristen berechnen sich automatisch.",
        "date": "Umzugsdatum:",
        "progress": "Fortschritt:",
        "headers": ["Fällig am", "Tage vor/nach Umzug", "Aufgabe", "Status", "Notizen"],
        "status": ["Offen", "Erledigt", "Nicht relevant"],
        "expat": "NUR BEI ZUZUG AUS DEM AUSLAND",
        "legend": "Rot = überfällig und noch offen. Fristen können je nach Kanton und Mietvertrag abweichen – im Zweifel bei Gemeinde/Verwaltung nachfragen.",
    },
    "en": {
        "file": "Moving-in-Switzerland-Planner.xlsx",
        "sheet": "Moving planner",
        "title": "Moving in Switzerland Planner",
        "sub": "Enter your moving date – every deadline is calculated automatically.",
        "date": "Moving date:",
        "progress": "Progress:",
        "headers": ["Due", "Days before/after move", "Task", "Status", "Notes"],
        "status": ["Open", "Done", "Not relevant"],
        "expat": "ONLY WHEN MOVING FROM ABROAD",
        "legend": "Red = overdue and still open. Deadlines can differ by canton and lease – if in doubt, ask your commune or landlord.",
    },
}


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def build(lang):
    t = TEXT[lang]
    wb = Workbook()
    ws = wb.active
    ws.title = t["sheet"]
    ws["A1"] = t["title"]
    ws["A1"].font = Font(size=18, bold=True, color=RED)
    ws["A2"] = t["sub"]
    ws["A2"].font = Font(italic=True, color="6B7280")
    ws["A4"] = t["date"]
    ws["A4"].font = Font(bold=True)
    ws["B4"].fill = fill(INPUT)
    ws["B4"].number_format = "DD.MM.YYYY"
    ws["B4"].border = BOX
    ws["D4"] = t["progress"]
    ws["D4"].font = Font(bold=True)

    for col, text in enumerate(t["headers"], 1):
        c = ws.cell(row=6, column=col, value=text)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill(DARK)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A7"

    open_, done, na = t["status"]
    dv = DataValidation(type="list", formula1='"' + ",".join(t["status"]) + '"', allow_blank=False)
    ws.add_data_validation(dv)

    def task_row(r, offset, de, en):
        ws.cell(row=r, column=1, value=f'=IF($B$4="","",$B$4+B{r})').number_format = "DD.MM.YYYY"
        ws.cell(row=r, column=2, value=offset)
        ws.cell(row=r, column=3, value=de if lang == "de" else en).alignment = Alignment(wrap_text=True, vertical="top")
        status = ws.cell(row=r, column=4, value=open_)
        status.fill = fill(INPUT)
        dv.add(status)
        ws.cell(row=r, column=5).fill = fill(INPUT)
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = BOX

    r = 7
    for task in TASKS:
        task_row(r, *task)
        r += 1
    first, last_main = 7, r - 1
    ws.cell(row=r, column=1, value=t["expat"]).font = Font(bold=True, color=RED)
    for col in range(1, 6):
        ws.cell(row=r, column=col).fill = fill(PALE_RED)
    r += 1
    for task in EXPAT_TASKS:
        task_row(r, *task)
        r += 1
    last = r - 1

    rng = f"D{first}:D{last}"
    ws["E4"] = f'=IFERROR(COUNTIF({rng},"{done}")/(COUNTA({rng})-COUNTIF({rng},"{na}")),0)'
    ws["E4"].number_format = "0%"
    ws["E4"].font = Font(bold=True, size=14, color="15803D")
    ws.cell(row=last + 2, column=1, value=t["legend"]).font = Font(italic=True, color="6B7280")

    red = Font(color="B91C1C", bold=True)
    ws.conditional_formatting.add(
        f"A{first}:E{last}",
        FormulaRule(formula=[f'AND($A{first}<>"",$A{first}<TODAY(),$D{first}="{open_}")'], font=red, fill=fill(PALE_RED)))
    ws.conditional_formatting.add(
        f"A{first}:E{last}",
        FormulaRule(formula=[f'$D{first}="{done}"'], font=Font(color="9CA3AF", strike=True)))

    for col, width in zip("ABCDE", (13, 12, 90, 15, 30)):
        ws.column_dimensions[col].width = width
    wb.save(DIST / t["file"])
    return t["file"]


if __name__ == "__main__":
    DIST.mkdir(exist_ok=True)
    for lang in TEXT:
        print("OK:", build(lang))
