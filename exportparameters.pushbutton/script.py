# -*- coding: utf-8 -*-
"""Element Parameters naar Excel
Selecteer parameters van elementen op de huidige sheet en exporteer ze naar Excel.
"""

__title__ = "Parameters\nNaar Excel"
__author__ = "PyRevit Tool"

from pyrevit import revit, DB, forms, script
from collections import defaultdict
import codecs
import os

# Haal de huidige document en view op
doc = revit.doc
uidoc = revit.uidoc
active_view = uidoc.ActiveView

# Controleer of we in een sheet zitten
if not isinstance(active_view, DB.ViewSheet):
    forms.alert("Je moet een sheet geopend hebben om deze tool te gebruiken.", exitscript=True)

current_sheet = active_view

# Verzamel alle viewports op de sheet
viewport_ids = current_sheet.GetAllViewports()

if not viewport_ids:
    forms.alert("Er zijn geen views geplaatst op deze sheet.", exitscript=True)

# Verzamel alle elementen van alle views op de sheet
all_elements = []
all_categories = set()

for vp_id in viewport_ids:
    viewport = doc.GetElement(vp_id)
    view_id = viewport.ViewId
    view = doc.GetElement(view_id)
    
    # Haal alle zichtbare elementen in deze view op
    collector = DB.FilteredElementCollector(doc, view_id)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    for elem in collector:
        # Sla alleen elementen op met een categorie
        if elem.Category:
            all_elements.append(elem)
            all_categories.add(elem.Category.Name)

if not all_elements:
    forms.alert("Er zijn geen elementen gevonden op de views van deze sheet.", exitscript=True)

# Laat gebruiker een categorie kiezen
selected_category = forms.SelectFromList.show(
    sorted(list(all_categories)),
    title="Kies een Categorie",
    button_name="Volgende"
)

if not selected_category:
    script.exit()

# Filter elementen op gekozen categorie
filtered_elements = [e for e in all_elements if e.Category.Name == selected_category]

# Verzamel alle unieke parameters van deze elementen
param_dict = defaultdict(int)
for elem in filtered_elements:
    for param in elem.Parameters:
        param_name = param.Definition.Name
        param_dict[param_name] += 1

# Sorteer parameters op naam
sorted_params = sorted(param_dict.keys())

# Laat gebruiker parameters selecteren
selected_params = forms.SelectFromList.show(
    sorted_params,
    title="Selecteer Parameters voor {} ({} elementen)".format(selected_category, len(filtered_elements)),
    multiselect=True,
    button_name="Exporteer naar CSV"
)

if not selected_params:
    script.exit()

# Vraag waar het bestand opgeslagen moet worden
from System.Windows.Forms import SaveFileDialog, DialogResult
save_dialog = SaveFileDialog()
save_dialog.Filter = "CSV bestanden (*.csv)|*.csv"
save_dialog.FileName = "{}_Sheet_{}.csv".format(
    selected_category.replace(" ", "_"),
    current_sheet.SheetNumber
)

if save_dialog.ShowDialog() == DialogResult.OK:
    file_path = save_dialog.FileName
    
    try:
        # Maak CSV bestand met UTF-8 encoding
        with codecs.open(file_path, 'w', encoding='utf-8-sig') as csvfile:
            # Schrijf header: Element ID + geselecteerde parameters
            header = ['Element ID', 'Category'] + list(selected_params)
            csvfile.write(';'.join(header) + '\n')
            
            # Schrijf data voor elk element
            for elem in filtered_elements:
                row_data = [
                    str(elem.Id.IntegerValue),
                    selected_category
                ]
                
                # Haal parameterwaarden op
                elem_params = {p.Definition.Name: p for p in elem.Parameters}
                
                for param_name in selected_params:
                    value = ""
                    
                    if param_name in elem_params:
                        param = elem_params[param_name]
                        
                        # Haal waarde op afhankelijk van type
                        try:
                            if param.HasValue:
                                if param.StorageType == DB.StorageType.String:
                                    value = param.AsString() or ""
                                elif param.StorageType == DB.StorageType.Integer:
                                    value = str(param.AsInteger())
                                elif param.StorageType == DB.StorageType.Double:
                                    # Probeer eerst AsValueString (toont waarde zoals in UI)
                                    value = param.AsValueString()
                                    if not value:
                                        # Fallback: converteer zelf
                                        try:
                                            display_unit = param.DisplayUnitType
                                            if display_unit == DB.DisplayUnitType.DUT_MILLIMETERS:
                                                value_mm = DB.UnitUtils.ConvertFromInternalUnits(
                                                    param.AsDouble(), 
                                                    DB.DisplayUnitType.DUT_MILLIMETERS
                                                )
                                                value = str(round(value_mm, 2))
                                            elif display_unit == DB.DisplayUnitType.DUT_METERS:
                                                value_m = DB.UnitUtils.ConvertFromInternalUnits(
                                                    param.AsDouble(), 
                                                    DB.DisplayUnitType.DUT_METERS
                                                )
                                                value = str(round(value_m, 3))
                                            else:
                                                value = str(round(param.AsDouble(), 4))
                                        except:
                                            value = str(param.AsDouble())
                                elif param.StorageType == DB.StorageType.ElementId:
                                    elem_id = param.AsElementId()
                                    if elem_id.IntegerValue > 0:
                                        ref_elem = doc.GetElement(elem_id)
                                        if ref_elem and hasattr(ref_elem, 'Name'):
                                            value = ref_elem.Name or ""
                                        else:
                                            value = str(elem_id.IntegerValue)
                                    else:
                                        value = ""
                                else:
                                    # Fallback naar AsValueString
                                    value = param.AsValueString() or ""
                        except:
                            value = ""
                    
                    # Escape puntkomma's en nieuwe regels in waarden
                    if value:
                        value = str(value).replace('\n', ' ').replace('\r', '')
                        if ';' in value or '"' in value:
                            value = '"{}"'.format(value.replace('"', '""'))
                    
                    row_data.append(value)
                
                # Schrijf rij
                csvfile.write(';'.join(row_data) + '\n')
        
        forms.alert(
            "Export succesvol!\n\n"
            "Bestand: {}\n\n"
            "{} elementen geëxporteerd\n"
            "{} parameters per element".format(
                file_path,
                len(filtered_elements),
                len(selected_params)
            ),
            title="Gereed"
        )
        
        # Open de map waar het bestand is opgeslagen
        os.startfile(os.path.dirname(file_path))
        
    except Exception as e:
        import traceback
        error_msg = "Er is een fout opgetreden bij het exporteren:\n\n{}\n\n{}".format(
            str(e),
            traceback.format_exc()
        )
        forms.alert(error_msg, title="Fout")
