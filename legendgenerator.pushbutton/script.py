# -*- coding: utf-8 -*-
"""
Sheet Legenda Generator
Genereert een legenda met filled regions van bouwkundige onderdelen op een sheet.
"""

__title__ = "Sheet\nLegenda"
__author__ = "Your Name"

from pyrevit import revit, DB, forms, script
from collections import defaultdict
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (
    Form, Label, Button, CheckBox, TextBox, DataGridView, DataGridViewCheckBoxColumn,
    DataGridViewTextBoxColumn, DialogResult, FormStartPosition, DockStyle,
    Panel, AnchorStyles, ScrollBars, BorderStyle, DataGridViewAutoSizeColumnsMode,
    DataGridViewSelectionMode
)
from System.Drawing import Point, Size, Font, FontStyle, Color

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


class LegendItem:
    """Data class voor een legenda item"""
    def __init__(self, category, type_name, elements, fill_pattern_id=None):
        self.category = category
        self.type_name = type_name
        self.display_name = type_name  # Kan worden aangepast door gebruiker
        self.elements = elements
        self.count = len(elements)
        self.selected = True
        self.fill_pattern_id = fill_pattern_id
        self.material = None
        self.color = None


class LegendSelectorForm(Form):
    """WPF-style dialoog voor selectie van legenda items"""
    
    def __init__(self, legend_items):
        self.legend_items = legend_items
        self.result_items = []
        
        self.Text = 'Sheet Legenda Generator'
        self.Size = Size(800, 600)
        self.MinimumSize = Size(600, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        
        self._build_ui()
        self._populate_grid()
    
    def _build_ui(self):
        # Instructie label
        self.label = Label()
        self.label.Text = 'Selecteer de onderdelen voor de legenda en pas eventueel de weergavenaam aan:'
        self.label.Location = Point(10, 10)
        self.label.Size = Size(760, 20)
        self.label.Font = Font(self.label.Font, FontStyle.Bold)
        self.Controls.Add(self.label)
        
        # DataGridView voor items
        self.grid = DataGridView()
        self.grid.Location = Point(10, 40)
        self.grid.Size = Size(760, 450)
        self.grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.grid.AllowUserToAddRows = False
        self.grid.AllowUserToDeleteRows = False
        self.grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid.RowHeadersVisible = False
        
        # Kolommen
        col_check = DataGridViewCheckBoxColumn()
        col_check.Name = 'Selected'
        col_check.HeaderText = 'Selecteer'
        col_check.Width = 60
        col_check.FillWeight = 10
        self.grid.Columns.Add(col_check)
        
        col_cat = DataGridViewTextBoxColumn()
        col_cat.Name = 'Category'
        col_cat.HeaderText = 'Categorie'
        col_cat.ReadOnly = True
        col_cat.FillWeight = 25
        self.grid.Columns.Add(col_cat)
        
        col_type = DataGridViewTextBoxColumn()
        col_type.Name = 'TypeName'
        col_type.HeaderText = 'Type Naam'
        col_type.ReadOnly = True
        col_type.FillWeight = 30
        self.grid.Columns.Add(col_type)
        
        col_display = DataGridViewTextBoxColumn()
        col_display.Name = 'DisplayName'
        col_display.HeaderText = 'Weergavenaam (aanpasbaar)'
        col_display.ReadOnly = False
        col_display.FillWeight = 30
        self.grid.Columns.Add(col_display)
        
        col_count = DataGridViewTextBoxColumn()
        col_count.Name = 'Count'
        col_count.HeaderText = 'Aantal'
        col_count.ReadOnly = True
        col_count.Width = 60
        col_count.FillWeight = 10
        self.grid.Columns.Add(col_count)
        
        self.Controls.Add(self.grid)
        
        # Button panel
        btn_panel = Panel()
        btn_panel.Location = Point(10, 500)
        btn_panel.Size = Size(760, 50)
        btn_panel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        
        # Alles selecteren button
        self.btn_select_all = Button()
        self.btn_select_all.Text = 'Alles selecteren'
        self.btn_select_all.Location = Point(0, 10)
        self.btn_select_all.Size = Size(120, 30)
        self.btn_select_all.Click += self._select_all_click
        btn_panel.Controls.Add(self.btn_select_all)
        
        # Niets selecteren button
        self.btn_select_none = Button()
        self.btn_select_none.Text = 'Niets selecteren'
        self.btn_select_none.Location = Point(130, 10)
        self.btn_select_none.Size = Size(120, 30)
        self.btn_select_none.Click += self._select_none_click
        btn_panel.Controls.Add(self.btn_select_none)
        
        # OK button
        self.btn_ok = Button()
        self.btn_ok.Text = 'Legenda Maken'
        self.btn_ok.Location = Point(500, 10)
        self.btn_ok.Size = Size(120, 30)
        self.btn_ok.Click += self._ok_click
        btn_panel.Controls.Add(self.btn_ok)
        
        # Cancel button
        self.btn_cancel = Button()
        self.btn_cancel.Text = 'Annuleren'
        self.btn_cancel.Location = Point(630, 10)
        self.btn_cancel.Size = Size(120, 30)
        self.btn_cancel.Click += self._cancel_click
        btn_panel.Controls.Add(self.btn_cancel)
        
        self.Controls.Add(btn_panel)
    
    def _populate_grid(self):
        """Vul de grid met legend items"""
        for item in self.legend_items:
            row_index = self.grid.Rows.Add()
            row = self.grid.Rows[row_index]
            row.Cells['Selected'].Value = item.selected
            row.Cells['Category'].Value = item.category
            row.Cells['TypeName'].Value = item.type_name
            row.Cells['DisplayName'].Value = item.display_name
            row.Cells['Count'].Value = str(item.count)
            row.Tag = item
    
    def _select_all_click(self, sender, e):
        for row in self.grid.Rows:
            row.Cells['Selected'].Value = True
    
    def _select_none_click(self, sender, e):
        for row in self.grid.Rows:
            row.Cells['Selected'].Value = False
    
    def _ok_click(self, sender, e):
        self.result_items = []
        for row in self.grid.Rows:
            if row.Cells['Selected'].Value:
                item = row.Tag
                item.display_name = row.Cells['DisplayName'].Value or item.type_name
                self.result_items.append(item)
        
        if not self.result_items:
            forms.alert('Selecteer minstens een onderdeel.')
            return
        
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def _cancel_click(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def get_active_sheet():
    """Haal de actieve sheet op."""
    active_view = uidoc.ActiveView
    if active_view and active_view.ViewType == DB.ViewType.DrawingSheet:
        return active_view
    return None


def get_fill_pattern_from_element(elem):
    """Probeer het fill pattern te krijgen van een element"""
    try:
        # Probeer via material
        mat_id = None
        
        # Zoek material parameter
        mat_param = elem.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
        if mat_param and mat_param.HasValue:
            mat_id = mat_param.AsElementId()
        
        if not mat_id or mat_id == DB.ElementId.InvalidElementId:
            # Probeer type material
            type_id = elem.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type:
                    mat_param = elem_type.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
                    if mat_param and mat_param.HasValue:
                        mat_id = mat_param.AsElementId()
        
        if mat_id and mat_id != DB.ElementId.InvalidElementId:
            material = doc.GetElement(mat_id)
            if material:
                # Haal surface pattern
                fg_pattern_id = material.SurfaceForegroundPatternId
                if fg_pattern_id and fg_pattern_id != DB.ElementId.InvalidElementId:
                    return fg_pattern_id, material
        
        return None, None
    except:
        return None, None


def get_elements_on_sheet(sheet):
    """Verzamel alle elementen die zichtbaar zijn op de sheet."""
    placed_views = sheet.GetAllPlacedViews()
    elements_dict = defaultdict(list)
    pattern_dict = {}
    
    # Categorien die we willen opnemen
    building_categories = [
        'Walls', 'Wanden',
        'Floors', 'Vloeren',
        'Roofs', 'Daken',
        'Ceilings', 'Plafonds',
        'Doors', 'Deuren',
        'Windows', 'Ramen',
        'Stairs', 'Trappen',
        'Railings', 'Leuningen',
        'Columns', 'Kolommen',
        'Structural Columns', 'Constructiekolommen',
        'Structural Framing', 'Constructiebalken',
        'Structural Foundations', 'Funderingen',
        'Curtain Walls', 'Vliesgevels',
        'Curtain Panels', 'Vliesgevel panelen',
        'Generic Models', 'Generieke modellen',
        'Furniture', 'Meubilair',
        'Casework', 'Kasten',
        'Plumbing Fixtures', 'Sanitair',
        'Mechanical Equipment', 'Mechanische apparatuur',
        'Electrical Equipment', 'Elektrische apparatuur',
        'Lighting Fixtures', 'Verlichtingsarmaturen',
        'Specialty Equipment', 'Speciale apparatuur'
    ]
    
    for view_id in placed_views:
        view = doc.GetElement(view_id)
        if view is None:
            continue
        
        try:
            collector = DB.FilteredElementCollector(doc, view_id)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            for elem in collector:
                if elem is None:
                    continue
                try:
                    if elem.Category and elem.Category.Name:
                        category_name = elem.Category.Name
                        
                        # Filter op bouwkundige categorien
                        if category_name not in building_categories:
                            continue
                        
                        type_id = elem.GetTypeId()
                        if type_id and type_id != DB.ElementId.InvalidElementId:
                            elem_type = doc.GetElement(type_id)
                            if elem_type:
                                type_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                                if type_param and type_param.AsString():
                                    type_name = type_param.AsString()
                                    key = (category_name, type_name)
                                    elements_dict[key].append(elem)
                                    
                                    # Haal fill pattern op (alleen eerste keer)
                                    if key not in pattern_dict:
                                        pattern_id, material = get_fill_pattern_from_element(elem)
                                        pattern_dict[key] = (pattern_id, material)
                except:
                    continue
        except:
            continue
    
    return elements_dict, pattern_dict


def get_drafting_view_type_id():
    """Vind de Drafting ViewFamilyType ID."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    for vft in collector:
        try:
            if vft.ViewFamily == DB.ViewFamily.Drafting:
                return vft.Id
        except:
            continue
    return None


def get_text_note_type_id(target_height_mm=2.5):
    """Haal text note type ID op met gewenste teksthoogte."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).ToElements()
    
    # Converteer mm naar feet
    target_height_feet = target_height_mm / 304.8
    
    best_match = None
    best_diff = float('inf')
    
    for tnt in collector:
        try:
            # Haal teksthoogte parameter op
            height_param = tnt.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
            if height_param:
                height = height_param.AsDouble()
                diff = abs(height - target_height_feet)
                if diff < best_diff:
                    best_diff = diff
                    best_match = tnt
        except:
            continue
    
    # Als geen match gevonden, pak eerste beschikbare
    if best_match is None and collector:
        best_match = collector[0]
    
    return best_match.Id if best_match else None


def get_filled_region_type():
    """Haal een filled region type op."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType)
    for frt in collector:
        return frt  # Eerste beschikbare
    return None


def create_filled_region(view, location, width, height, filled_region_type_id):
    """Maak een filled region rechthoek"""
    try:
        x, y = location.X, location.Y
        
        # Maak rechthoek curve loop
        p1 = DB.XYZ(x, y, 0)
        p2 = DB.XYZ(x + width, y, 0)
        p3 = DB.XYZ(x + width, y - height, 0)
        p4 = DB.XYZ(x, y - height, 0)
        
        curves = [
            DB.Line.CreateBound(p1, p2),
            DB.Line.CreateBound(p2, p3),
            DB.Line.CreateBound(p3, p4),
            DB.Line.CreateBound(p4, p1)
        ]
        
        curve_loop = DB.CurveLoop()
        for curve in curves:
            curve_loop.Append(curve)
        
        curve_loops = [curve_loop]
        
        # Maak filled region
        filled_region = DB.FilledRegion.Create(doc, filled_region_type_id, view.Id, curve_loops)
        return filled_region
    except Exception as e:
        print("Filled region error: {}".format(str(e)))
        return None


def create_text_note(doc, view_id, location, text, text_type_id):
    """Maak een text note met fallback methodes"""
    try:
        options = DB.TextNoteOptions(text_type_id)
        return DB.TextNote.Create(doc, view_id, location, text, options)
    except:
        pass
    
    try:
        return DB.TextNote.Create(doc, view_id, location, text, text_type_id)
    except:
        pass
    
    return None


def create_legend(sheet, selected_items):
    """Maak de legenda view met content in 2 kolommen"""
    
    # Haal benodigde types op (2.5mm teksthoogte)
    text_type_id = get_text_note_type_id(2.5)
    if not text_type_id:
        forms.alert("Geen TextNoteType gevonden.", exitscript=True)
        return None
    
    drafting_type_id = get_drafting_view_type_id()
    if not drafting_type_id:
        forms.alert("Geen Drafting ViewFamilyType gevonden.", exitscript=True)
        return None
    
    filled_region_type = get_filled_region_type()
    filled_region_type_id = filled_region_type.Id if filled_region_type else None
    
    # Sheet naam voor view name
    try:
        sheet_name = sheet.Name
    except:
        sheet_name = "Sheet"
    
    # Transactie 1: Maak view
    t1 = DB.Transaction(doc, "Maak Legenda View")
    t1.Start()
    
    try:
        legend = DB.ViewDrafting.Create(doc, drafting_type_id)
        
        # Zet view naam op Legend_(sheetname)
        base_name = "Legend_{}".format(sheet_name)
        counter = 0
        while counter < 100:
            try:
                name = base_name if counter == 0 else "{}_{}".format(base_name, counter)
                legend.Name = name
                break
            except:
                counter += 1
        
        legend_id = legend.Id
        t1.Commit()
        
    except Exception as e:
        if t1.HasStarted():
            t1.RollBack()
        forms.alert("Fout bij aanmaken view: {}".format(str(e)))
        return None
    
    # Haal view opnieuw op
    legend = doc.GetElement(legend_id)
    if legend is None:
        forms.alert("Kon view niet ophalen na aanmaken.")
        return None
    
    # Transactie 2: Voeg content toe
    t2 = DB.Transaction(doc, "Voeg Legenda Content Toe")
    t2.Start()
    
    try:
        # === LAYOUT INSTELLINGEN (in feet) ===
        mm_to_feet = 1.0 / 304.8
        
        # Kolom 1: Filled Region
        col1_x = 0.0                        # Start positie kolom 1
        box_width = 10.0 * mm_to_feet       # 10mm breed
        box_height = 5.0 * mm_to_feet       # 5mm hoog
        
        # Kolom 2: Tekst
        col2_x = 15.0 * mm_to_feet          # 15mm vanaf links
        
        # Rij instellingen
        row_height = 8.0 * mm_to_feet       # 8mm tussen rijen
        text_height = 2.5 * mm_to_feet      # 2.5mm teksthoogte
        
        y_start = 0.0
        y_offset = y_start
        
        # Sorteer items op categorie en naam
        sorted_items = sorted(selected_items, key=lambda x: (x.category, x.display_name))
        
        for item in sorted_items:
            # Bereken verticaal midden van de rij
            row_center_y = y_offset - (box_height / 2)
            
            # === KOLOM 1: Filled Region (verticaal gecentreerd) ===
            if filled_region_type_id:
                box_location = DB.XYZ(col1_x, y_offset, 0.0)
                create_filled_region(legend, box_location, box_width, box_height, filled_region_type_id)
            
            # === KOLOM 2: Weergavenaam (verticaal gecentreerd met box) ===
            text_y = row_center_y + (text_height / 2)
            text_location = DB.XYZ(col2_x, text_y, 0.0)
            create_text_note(doc, legend_id, text_location, item.display_name, text_type_id)
            
            y_offset -= row_height
        
        t2.Commit()
        return legend
        
    except Exception as e:
        if t2.HasStarted():
            t2.RollBack()
        forms.alert("Fout bij content: {}".format(str(e)))
        return None


def main():
    """Hoofdfunctie"""
    output.print_md("## Sheet Legenda Generator")
    output.print_md("---")
    
    # Check actieve sheet
    sheet = get_active_sheet()
    if not sheet:
        forms.alert("Open een sheet view om deze tool te gebruiken.", exitscript=True)
        return
    
    try:
        sheet_name = "{} - {}".format(sheet.SheetNumber, sheet.Name)
    except:
        sheet_name = "Sheet"
    
    output.print_md("**Sheet:** {}".format(sheet_name))
    output.print_md("Verzamel elementen...")
    
    # Verzamel elementen
    elements_dict, pattern_dict = get_elements_on_sheet(sheet)
    
    if not elements_dict:
        forms.alert("Geen bouwkundige componenten gevonden op deze sheet.", exitscript=True)
        return
    
    output.print_md("**Gevonden:** {} unieke types".format(len(elements_dict)))
    
    # Maak legend items
    legend_items = []
    for (category, type_name), elements in elements_dict.items():
        pattern_id, material = pattern_dict.get((category, type_name), (None, None))
        item = LegendItem(category, type_name, elements, pattern_id)
        item.material = material
        legend_items.append(item)
    
    # Sorteer op categorie en naam
    legend_items.sort(key=lambda x: (x.category, x.type_name))
    
    # Toon selectie dialoog
    dialog = LegendSelectorForm(legend_items)
    result = dialog.ShowDialog()
    
    if result != DialogResult.OK:
        output.print_md("**Geannuleerd**")
        return
    
    selected_items = dialog.result_items
    output.print_md("**Geselecteerd:** {} items".format(len(selected_items)))
    
    # Maak legenda
    output.print_md("\n### Legenda aanmaken...")
    legend = create_legend(sheet, selected_items)
    
    if legend:
        output.print_md("\n## Klaar!")
        output.print_md("Legenda view '{}' is aangemaakt.".format(legend.Name))
        
        # Open de legenda view
        try:
            uidoc.ActiveView = legend
        except:
            pass
        
        # Vraag of de legenda op de sheet geplaatst moet worden
        place_result = forms.alert(
            "Wil je de legenda op de sheet plaatsen?",
            yes=True, no=True
        )
        
        if place_result:
            try:
                t3 = DB.Transaction(doc, "Plaats Legenda op Sheet")
                t3.Start()
                
                # Plaats viewport
                viewport_location = DB.XYZ(0.5, 0.5, 0)
                viewport = DB.Viewport.Create(doc, sheet.Id, legend.Id, viewport_location)
                
                # Zet viewport title op "No Title" / verberg titel
                if viewport:
                    # Zoek een viewport type zonder titel
                    vp_types = DB.FilteredElementCollector(doc).OfClass(DB.ElementType).ToElements()
                    no_title_type = None
                    
                    for vp_type in vp_types:
                        try:
                            if vp_type.FamilyName == "Viewport":
                                type_name = vp_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                                if type_name and ("No Title" in type_name or "Geen titel" in type_name or "Empty" in type_name or "Leeg" in type_name):
                                    no_title_type = vp_type.Id
                                    break
                        except:
                            continue
                    
                    # Als geen "No Title" type gevonden, probeer titel te verbergen via parameter
                    if no_title_type:
                        viewport.ChangeTypeId(no_title_type)
                    else:
                        # Probeer Show Title parameter uit te zetten
                        try:
                            title_param = viewport.get_Parameter(DB.BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL)
                            if title_param and not title_param.IsReadOnly:
                                title_param.Set(0)
                        except:
                            pass
                
                t3.Commit()
                
                # Ga terug naar sheet
                uidoc.ActiveView = sheet
                
                forms.alert("Legenda is op de sheet geplaatst.\nVersleep naar gewenste positie.")
            except Exception as e:
                forms.alert("Kon legenda niet plaatsen: {}".format(str(e)))


if __name__ == "__main__":
    main()
