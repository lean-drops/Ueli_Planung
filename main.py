from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

# Neues Arbeitsbuch und Arbeitsblatt erstellen
wb = Workbook()
ws = wb.active

# Dropdown-Optionen für die Tage
tage_optionen = 'Di, Mi, Do, Fr, Sa, So'

# Erstellen der DataValidation-Objekte
dv = DataValidation(type="list", formula1=f'"{tage_optionen}"', showDropDown=True)

# Kommentar hinzufügen, um die Nutzung zu erklären
kommentar_text = "Bitte wählen Sie die Tage aus, an denen Sie arbeiten können."
kommentar = Comment(kommentar_text, "System")

# Anwenden der DataValidation und des Kommentars auf Zellen
for row in ws.iter_rows(min_row=2, max_row=7, min_col=1, max_col=1):
    for cell in row:
        cell.comment = kommentar
        ws.add_data_validation(dv)
        dv.add(cell)

# Speichern der Excel-Datei
excel_dateiname = "Mitarbeiter_Verfügbarkeiten.xlsx"
wb.save(filename=excel_dateiname)

print(f"Excel-Datei '{excel_dateiname}' wurde erfolgreich erstellt.")
