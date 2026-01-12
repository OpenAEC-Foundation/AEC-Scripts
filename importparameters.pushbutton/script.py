# -*- coding: utf-8 -*-
"""CSV naar Sheet Tabel
Importeer CSV data en plaats het als een tabel op een sheet.
Gebruikt text notes in een drafting view voor maximale flexibiliteit.
"""

__title__ = "CSV naar\nSheet Tabel"
__author__ = "PyRevit Tool"

from pyrevit import revit, DB, forms, script
import codecs

# Haal de huidige document op
doc = revit.doc
uidoc = revit.uidoc
active_view = uidoc.ActiveView

# Check of we in een sheet zitten
if isinstance(active_view, DB.ViewSheet):
    target_sheet = active_view
    use_current_sheet = True
else:
    use_current_sheet = False
    target_sheet = None

# Vraag gebruiker om CSV bestand te selecteren
from System.Windows.Forms import OpenFileDialog, DialogResult
open_dialog = OpenFileDialog()
open_dialog.Filter = "CSV bestanden (*.csv)|*.csv"
open_dialog.Title = "Selecteer CSV bestand"

if open_dialog.ShowDialog() != DialogResult.OK:
    script.exit()

file_path = open_dialog.FileName

# Lees CSV bestand
try:
    with codecs.open(file_path, 'r', encoding='utf-8-sig') as csvfile:
        lines = csvfile.readlines()
        
    if not lines:
        forms.alert("CSV bestand is leeg.", exitscript=True)
    
    # Parse header en data
    delimiter = ';' if ';' in lines[0] else ','
    
    header = [h.strip() for h in lines[0].strip().split(delimiter)]
    data_rows = []
    
    for line in lines[1:]:
        if line.strip():
            row = [cell.strip().strip('"') for cell in line.strip().split(delimiter)]
            while len(row) < len(header):
                row.append("")
            data_rows.append(row)
    
except Exception as e:
    forms.alert("Kon CSV bestand niet lezen:\n\n{}".format(str(e)), title="Fout")
    script.exit()

if not data_rows:
    forms.alert("Geen data gevonden in CSV bestand.", exitscript=True)

# Bepaal target sheet
if not use_current_sheet:
    # Kies een sheet om de tabel op te plaatsen
    all_sheets = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSheet)\
        .WhereElementIsNotElementType()\
        .ToElements()

    if not all_sheets:
        forms.alert("Geen sheets gevonden in project.", exitscript=True)

    sheet_options = {"{} - {}".format(s.SheetNumber, s.Name): s for s in all_sheets}

    selected_sheet_name = forms.SelectFromList.show(
        sorted(sheet_options.keys()),
        title="Selecteer Sheet voor Tabel",
        button_name="OK"
    )

    if not selected_sheet_name:
        script.exit()

    target_sheet = sheet_options[selected_sheet_name]

# Vraag tabel naam
table_name = forms.ask_for_string(
    prompt="Geef een naam voor de drafting view:",
    title="View Naam",
    default="CSV Tabel - {}".format(file_path.split('\\')[-1].replace('.csv', ''))
)

if not table_name:
    script.exit()

# Start transactie
t = DB.Transaction(doc, "Maak CSV Tabel op Sheet")
t.Start()

try:
    # Maak een Drafting View
    view_family_types = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewFamilyType)\
        .ToElements()
    
    drafting_view_type = None
    for vft in view_family_types:
        if vft.ViewFamily == DB.ViewFamily.Drafting:
            drafting_view_type = vft
            break
    
    if not drafting_view_type:
        forms.alert("Geen Drafting View type gevonden.", exitscript=True)
    
    drafting_view = DB.ViewDrafting.Create(doc, drafting_view_type.Id)
    drafting_view.Name = table_name
    
    # Zoek text note type
    text_note_types = DB.FilteredElementCollector(doc)\
        .OfClass(DB.TextNoteType)\
        .ToElements()
    
    if not text_note_types:
        forms.alert("Geen Text Note types gevonden.", exitscript=True)
    
    # Zoek kleinste text type (voor tabel data)
    text_note_type = min(text_note_types, key=lambda t: t.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).AsDouble())
    
    # Bereken kolom breedtes
    col_widths = []
    for col_idx in range(len(header)):
        max_length = len(header[col_idx])
        for row_data in data_rows:
            if col_idx < len(row_data):
                max_length = max(max_length, len(str(row_data[col_idx])))
        # Breedte in feet (ongeveer 0.04 feet per karakter)
        width = max(max_length * 0.04, 0.3)
        col_widths.append(width)
    
    # Start positie voor tabel
    start_x = 0.5  # feet
    start_y = 2.0  # feet
    row_height = 0.15  # feet
    
    # Maak detail lines voor tabel grid
    detail_line_style = None
    line_styles = drafting_view.GetValidTypes()
    
    # Gebruik gewoon de eerste beschikbare line style
    if line_styles:
        detail_line_style = line_styles[0]
    
    # Teken horizontale lijnen
    table_width = sum(col_widths)
    
    # Bovenkant tabel
    line = DB.Line.CreateBound(
        DB.XYZ(start_x, start_y, 0),
        DB.XYZ(start_x + table_width, start_y, 0)
    )
    doc.Create.NewDetailCurve(drafting_view, line)
    
    # Lijn onder header
    line = DB.Line.CreateBound(
        DB.XYZ(start_x, start_y - row_height, 0),
        DB.XYZ(start_x + table_width, start_y - row_height, 0)
    )
    doc.Create.NewDetailCurve(drafting_view, line)
    
    # Lijnen voor elke data rij
    for row_idx in range(len(data_rows) + 1):
        y = start_y - (row_idx + 1) * row_height
        line = DB.Line.CreateBound(
            DB.XYZ(start_x, y, 0),
            DB.XYZ(start_x + table_width, y, 0)
        )
        doc.Create.NewDetailCurve(drafting_view, line)
    
    # Teken verticale lijnen
    x_pos = start_x
    for col_idx in range(len(header) + 1):
        y_end = start_y - (len(data_rows) + 1) * row_height
        line = DB.Line.CreateBound(
            DB.XYZ(x_pos, start_y, 0),
            DB.XYZ(x_pos, y_end, 0)
        )
        doc.Create.NewDetailCurve(drafting_view, line)
        
        if col_idx < len(col_widths):
            x_pos += col_widths[col_idx]
    
    # Plaats tekst voor header
    x_pos = start_x + 0.02  # Kleine padding
    for col_idx, col_name in enumerate(header):
        location = DB.XYZ(x_pos, start_y - row_height * 0.5, 0)
        text_note = DB.TextNote.Create(
            doc,
            drafting_view.Id,
            location,
            str(col_name),
            text_note_type.Id
        )
        x_pos += col_widths[col_idx]
    
    # Plaats tekst voor data rijen
    for row_idx, row_data in enumerate(data_rows):
        x_pos = start_x + 0.02
        y_pos = start_y - (row_idx + 1.5) * row_height
        
        for col_idx, cell_value in enumerate(row_data):
            if col_idx < len(col_widths):
                location = DB.XYZ(x_pos, y_pos, 0)
                text_note = DB.TextNote.Create(
                    doc,
                    drafting_view.Id,
                    location,
                    str(cell_value),
                    text_note_type.Id
                )
                x_pos += col_widths[col_idx]
    
    # Plaats drafting view op sheet
    # Zoek vrije locatie op sheet (centrum)
    location = DB.XYZ(1.0, 1.5, 0)
    
    viewport = DB.Viewport.Create(
        doc,
        target_sheet.Id,
        drafting_view.Id,
        location
    )
    
    t.Commit()
    
    # Open de sheet
    uidoc.ActiveView = target_sheet
    
    forms.alert(
        "CSV tabel geplaatst op sheet!\n\n"
        "Sheet: {} - {}\n"
        "Drafting View: {}\n"
        "Kolommen: {}\n"
        "Rijen: {}\n\n"
        "De tabel is gemaakt met text notes en detail lines.\n"
        "Je kunt de viewport verplaatsen op de sheet.".format(
            target_sheet.SheetNumber,
            target_sheet.Name,
            table_name,
            len(header),
            len(data_rows)
        ),
        title="Tabel Geplaatst"
    )
    
except Exception as e:
    import traceback
    t.RollBack()
    error_msg = "Er is een fout opgetreden:\n\n{}\n\n{}".format(
        str(e),
        traceback.format_exc()
    )
    forms.alert(error_msg, title="Fout")
