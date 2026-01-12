# -*- coding: utf-8 -*-
"""
Auto Dimensionering Tool
========================
Voegt automatisch maatvoering toe op views met selecteerbare element types.
Inclusief wanddikte maatvoering en keuze voor plaatsing (links/rechts/boven/onder).

Auteur: JMK
Panel: JMK
"""

__title__ = "Auto\nMaatvoering"
__doc__ = "Automatische maatvoering voor grids, wanden, kolommen etc."
__author__ = "JMK"

# Imports
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms, revit

# UI Template imports
from ui_template import BaseForm, UIFactory, COLORS, DPIScaler, run_dialog
from System.Windows.Forms import DialogResult, GroupBox, CheckBox, ComboBox, TextBox, Label
from System.Drawing import Point, Size

# Document context
doc = revit.doc
uidoc = revit.uidoc

# Conversie factor (mm naar feet)
MM_TO_FEET = 1.0 / 304.8


# =============================================================================
# UI DIALOG
# =============================================================================

class DimensionOptionsDialog(BaseForm):
    """Dialog voor auto maatvoering opties"""
    
    ELEMENT_TYPES = [
        ('Grids', True),
        ('Wanden', False),
        ('Kolommen', True),
        ('Deuren', False),
        ('Ramen', False),
        ('Structurele kolommen', True),
        ('Draagbalken', False)
    ]
    
    def __init__(self):
        super(DimensionOptionsDialog, self).__init__(
            title="Auto Maatvoering",
            width=480,
            height=720,
            resizable=False
        )
        self._setup_ui()
    
    def _setup_ui(self):
        """Bouw de UI"""
        
        # === TITEL ===
        self.add_title("Auto Maatvoering")
        
        # === SECTIE 1: ELEMENT TYPES ===
        self.add_section("1. Element Types", row=1)
        
        self.element_checkboxes = {}
        row = 2
        for name, default in self.ELEMENT_TYPES:
            cb = self.add_checkbox(name, row=row, checked=default, name="elem_" + name)
            self.element_checkboxes[name] = cb
            row += 1
        
        self.add_separator(row=row)
        row += 1
        
        # === SECTIE 2: WAND OPTIES ===
        self.add_section("2. Wand Opties", row=row)
        row += 1
        
        self.wall_center_cb = self.add_checkbox("Hart wand (centerline)", row=row, checked=True)
        row += 1
        self.wall_faces_cb = self.add_checkbox("Wandvlakken (buitenkanten)", row=row, checked=False)
        row += 1
        self.wall_thickness_cb = self.add_checkbox("Wanddikte maatvoering", row=row, checked=False)
        row += 1
        
        self.add_separator(row=row)
        row += 1
        
        # === SECTIE 3: PLAATSING ===
        self.add_section("3. Maatvoering Plaatsing", row=row)
        row += 1
        
        self.add_label("Horizontale maten:", row=row, bold=True)
        row += 1
        
        self.place_top_cb = self.add_checkbox("Bovenkant", row=row, checked=False)
        self.place_bottom_cb = self.add_checkbox("Onderkant", row=row, col=1, checked=True)
        row += 1
        
        self.add_label("Verticale maten:", row=row, bold=True)
        row += 1
        
        self.place_left_cb = self.add_checkbox("Linkerkant", row=row, checked=True)
        self.place_right_cb = self.add_checkbox("Rechterkant", row=row, col=1, checked=False)
        row += 1
        
        self.add_separator(row=row)
        row += 1
        
        # === SECTIE 4: INSTELLINGEN ===
        self.add_section("4. Instellingen", row=row)
        row += 1
        
        self.total_dim_cb = self.add_checkbox("Totaal maatvoering toevoegen", row=row, checked=True)
        row += 1
        
        # Offsets
        self.add_label("Offset detail (mm):", row=row)
        self.detail_offset = self.add_numeric(row=row, col=1, min_val=100, max_val=5000, default=500)
        row += 1
        
        self.add_label("Offset totaal (mm):", row=row)
        self.total_offset = self.add_numeric(row=row, col=1, min_val=100, max_val=5000, default=1000)
        row += 1
        
        self.add_label("Offset wanddikte (mm):", row=row)
        self.thickness_offset = self.add_numeric(row=row, col=1, min_val=100, max_val=2000, default=300)
        row += 1
        
        self.add_separator(row=row)
        row += 1
        
        # === RICHTING ===
        self.add_label("Richting:", row=row)
        self.direction_combo = self.add_combobox(
            items=["Horizontaal", "Verticaal", "Beide"],
            row=row, col=1
        )
        self.direction_combo.SelectedIndex = 2  # Default: Beide
        row += 1
        
        # === BUTTONS ===
        self.skip_row()
        self.add_button_row([
            ("Maatvoering Toepassen", self._on_apply, True),
            ("Annuleren", self._on_cancel, False)
        ], row=row + 1)
    
    def _on_apply(self, sender, args):
        """Handle toepassen button"""
        # Validatie
        selected_types = self.get_selected_types()
        if not selected_types:
            self.show_warning("Selecteer minimaal één element type.")
            return
        
        direction = self.direction_combo.SelectedItem
        placement = self.get_placement_options()
        
        # Check plaatsing
        if direction in ['Horizontaal', 'Beide']:
            if not placement['top'] and not placement['bottom']:
                self.show_warning("Selecteer minstens bovenkant of onderkant\nvoor horizontale maten.")
                return
        
        if direction in ['Verticaal', 'Beide']:
            if not placement['left'] and not placement['right']:
                self.show_warning("Selecteer minstens linkerkant of rechterkant\nvoor verticale maten.")
                return
        
        self.close_ok()
    
    def _on_cancel(self, sender, args):
        """Handle annuleren button"""
        self.close_cancel()
    
    # -------------------------------------------------------------------------
    # Data getters
    # -------------------------------------------------------------------------
    
    def get_selected_types(self):
        """Retourneer lijst van geselecteerde element types"""
        return [name for name, cb in self.element_checkboxes.items() if cb.Checked]
    
    def get_wall_options(self):
        """Retourneer wand opties"""
        return {
            'center': self.wall_center_cb.Checked,
            'faces': self.wall_faces_cb.Checked,
            'thickness': self.wall_thickness_cb.Checked
        }
    
    def get_placement_options(self):
        """Retourneer plaatsing opties"""
        return {
            'top': self.place_top_cb.Checked,
            'bottom': self.place_bottom_cb.Checked,
            'left': self.place_left_cb.Checked,
            'right': self.place_right_cb.Checked
        }
    
    def get_offsets(self):
        """Retourneer offset waarden in feet"""
        return {
            'detail': float(self.detail_offset.Value) * MM_TO_FEET,
            'total': float(self.total_offset.Value) * MM_TO_FEET,
            'thickness': float(self.thickness_offset.Value) * MM_TO_FEET
        }
    
    def get_direction(self):
        """Retourneer geselecteerde richting"""
        return self.direction_combo.SelectedItem
    
    def get_add_total(self):
        """Retourneer of totaal maatvoering moet worden toegevoegd"""
        return self.total_dim_cb.Checked


# =============================================================================
# DIMENSIONING LOGIC
# =============================================================================

def get_elements_on_view(view, selected_types):
    """Verzamel alle relevante elementen op de view"""
    elements_dict = {
        'Grids': [],
        'Wanden': [],
        'Kolommen': [],
        'Deuren': [],
        'Ramen': [],
        'Structurele kolommen': [],
        'Draagbalken': []
    }
    
    type_mapping = {
        'Grids': (Grid, None),
        'Wanden': (None, BuiltInCategory.OST_Walls),
        'Kolommen': (None, BuiltInCategory.OST_Columns),
        'Deuren': (None, BuiltInCategory.OST_Doors),
        'Ramen': (None, BuiltInCategory.OST_Windows),
        'Structurele kolommen': (None, BuiltInCategory.OST_StructuralColumns),
        'Draagbalken': (None, BuiltInCategory.OST_StructuralFraming)
    }
    
    for type_name in selected_types:
        if type_name in type_mapping:
            class_type, category = type_mapping[type_name]
            if class_type:
                elements = FilteredElementCollector(doc, view.Id).OfClass(class_type).ToElements()
            else:
                elements = FilteredElementCollector(doc, view.Id)\
                    .OfCategory(category).WhereElementIsNotElementType().ToElements()
            elements_dict[type_name] = list(elements)
    
    return elements_dict


def get_wall_face_references(wall, view):
    """Haal de binnen- en buitenvlak referenties op voor een wand"""
    faces = {'interior': None, 'exterior': None, 'interior_loc': None, 'exterior_loc': None}
    
    try:
        options = Options()
        options.ComputeReferences = True
        options.IncludeNonVisibleObjects = False
        options.View = view
        
        geom = wall.get_Geometry(options)
        if not geom:
            return faces
        
        loc = wall.Location
        if not isinstance(loc, LocationCurve):
            return faces
        
        curve = loc.Curve
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        wall_dir = (end - start).Normalize()
        wall_normal = XYZ(-wall_dir.Y, wall_dir.X, 0)
        
        side_faces = []
        
        for geom_obj in geom:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                for face in geom_obj.Faces:
                    if isinstance(face, PlanarFace):
                        face_normal = face.FaceNormal
                        dot = abs(face_normal.X * wall_normal.X + face_normal.Y * wall_normal.Y)
                        if dot > 0.9:
                            ref = face.Reference
                            if ref:
                                side_faces.append({
                                    'ref': ref,
                                    'origin': face.Origin,
                                    'normal': face_normal
                                })
        
        if len(side_faces) >= 2:
            side_faces.sort(key=lambda f: f['origin'].X * wall_normal.X + f['origin'].Y * wall_normal.Y)
            faces['interior'] = side_faces[0]['ref']
            faces['interior_loc'] = side_faces[0]['origin']
            faces['exterior'] = side_faces[-1]['ref']
            faces['exterior_loc'] = side_faces[-1]['origin']
    except:
        pass
    
    return faces


def create_wall_thickness_dimensions(view, walls, direction, offset, placement):
    """Maak wanddikte maatvoering voor elke wand"""
    dims = []
    
    for wall in walls:
        try:
            loc = wall.Location
            if not isinstance(loc, LocationCurve):
                continue
            
            curve = loc.Curve
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            is_vertical_wall = abs(start.X - end.X) < abs(start.Y - end.Y)
            
            if direction == 'horizontal' and not is_vertical_wall:
                continue
            if direction == 'vertical' and is_vertical_wall:
                continue
            
            faces = get_wall_face_references(wall, view)
            
            if faces['interior'] and faces['exterior']:
                mid_point = curve.Evaluate(0.5, True)
                z = mid_point.Z
                
                if is_vertical_wall:
                    line_start = XYZ(mid_point.X - 10, mid_point.Y - offset, z)
                    line_end = XYZ(mid_point.X + 10, mid_point.Y - offset, z)
                else:
                    line_start = XYZ(mid_point.X - offset, mid_point.Y - 10, z)
                    line_end = XYZ(mid_point.X - offset, mid_point.Y + 10, z)
                
                dim_line = Line.CreateBound(line_start, line_end)
                
                ref_array = ReferenceArray()
                ref_array.Append(faces['interior'])
                ref_array.Append(faces['exterior'])
                
                try:
                    dim = doc.Create.NewDimension(view, dim_line, ref_array)
                    if dim:
                        dims.append(dim)
                except:
                    pass
        except:
            continue
    
    return dims


def collect_references(elements_dict, direction, wall_options, view):
    """Verzamel alle referenties voor maatvoering"""
    references = []
    
    for element_type, elements in elements_dict.items():
        for elem in elements:
            if isinstance(elem, Grid):
                curve = elem.Curve
                start = curve.GetEndPoint(0)
                end = curve.GetEndPoint(1)
                is_vertical_grid = abs(start.X - end.X) < abs(start.Y - end.Y)
                
                if (direction == 'horizontal' and is_vertical_grid) or \
                   (direction == 'vertical' and not is_vertical_grid):
                    mid_point = XYZ((start.X + end.X) / 2, (start.Y + end.Y) / 2, start.Z)
                    references.append({
                        'ref': Reference(elem),
                        'location': mid_point,
                        'element': elem,
                        'type': 'grid'
                    })
            
            elif isinstance(elem, Wall):
                try:
                    loc = elem.Location
                    if not isinstance(loc, LocationCurve):
                        continue
                    
                    curve = loc.Curve
                    start = curve.GetEndPoint(0)
                    end = curve.GetEndPoint(1)
                    is_vertical_wall = abs(start.X - end.X) < abs(start.Y - end.Y)
                    
                    if (direction == 'horizontal' and not is_vertical_wall) or \
                       (direction == 'vertical' and is_vertical_wall):
                        continue
                    
                    mid_point = curve.Evaluate(0.5, True)
                    
                    if wall_options.get('center', True):
                        references.append({
                            'ref': Reference(elem),
                            'location': mid_point,
                            'element': elem,
                            'type': 'center'
                        })
                    
                    if wall_options.get('faces', False):
                        faces = get_wall_face_references(elem, view)
                        if faces['interior'] and faces['interior_loc']:
                            references.append({
                                'ref': faces['interior'],
                                'location': faces['interior_loc'],
                                'element': elem,
                                'type': 'face_int'
                            })
                        if faces['exterior'] and faces['exterior_loc']:
                            references.append({
                                'ref': faces['exterior'],
                                'location': faces['exterior_loc'],
                                'element': elem,
                                'type': 'face_ext'
                            })
                except:
                    pass
            
            else:
                try:
                    loc = elem.Location
                    if isinstance(loc, LocationPoint):
                        references.append({
                            'ref': Reference(elem),
                            'location': loc.Point,
                            'element': elem,
                            'type': 'point'
                        })
                    elif isinstance(loc, LocationCurve):
                        curve = loc.Curve
                        start = curve.GetEndPoint(0)
                        end = curve.GetEndPoint(1)
                        is_vertical = abs(start.X - end.X) < abs(start.Y - end.Y)
                        
                        if (direction == 'horizontal' and is_vertical) or \
                           (direction == 'vertical' and not is_vertical):
                            mid_point = curve.Evaluate(0.5, True)
                            references.append({
                                'ref': Reference(elem),
                                'location': mid_point,
                                'element': elem,
                                'type': 'curve'
                            })
                except:
                    pass
    
    # Sorteer referenties
    if direction == 'horizontal':
        references.sort(key=lambda x: x['location'].X)
    else:
        references.sort(key=lambda x: x['location'].Y)
    
    return references


def create_dimension_line(view, references, direction, add_total, offsets, placement, side):
    """Maak een maatvoeringslijn aan een specifieke kant"""
    if not references or len(references) < 2:
        return []
    
    dims = []
    all_points = [r['location'] for r in references]
    
    detail_offset = offsets['detail']
    total_offset = offsets['total']
    
    # Bepaal offset richting op basis van side
    if direction == 'horizontal':
        min_y = min(p.Y for p in all_points)
        max_y = max(p.Y for p in all_points)
        max_x = max(p.X for p in all_points)
        min_x = min(p.X for p in all_points)
        z = all_points[0].Z
        
        if side == 'bottom':
            detail_line_y = min_y - detail_offset
            total_line_y = min_y - total_offset
        else:  # top
            detail_line_y = max_y + detail_offset
            total_line_y = max_y + total_offset
        
        detail_line = Line.CreateBound(
            XYZ(min_x - 1, detail_line_y, z),
            XYZ(max_x + 1, detail_line_y, z)
        )
        total_line = Line.CreateBound(
            XYZ(min_x - 1, total_line_y, z),
            XYZ(max_x + 1, total_line_y, z)
        )
    else:  # vertical
        min_x = min(p.X for p in all_points)
        max_x = max(p.X for p in all_points)
        max_y = max(p.Y for p in all_points)
        min_y = min(p.Y for p in all_points)
        z = all_points[0].Z
        
        if side == 'left':
            detail_line_x = min_x - detail_offset
            total_line_x = min_x - total_offset
        else:  # right
            detail_line_x = max_x + detail_offset
            total_line_x = max_x + total_offset
        
        detail_line = Line.CreateBound(
            XYZ(detail_line_x, min_y - 1, z),
            XYZ(detail_line_x, max_y + 1, z)
        )
        total_line = Line.CreateBound(
            XYZ(total_line_x, min_y - 1, z),
            XYZ(total_line_x, max_y + 1, z)
        )
    
    try:
        # Detail maatvoering
        for i in range(len(references) - 1):
            temp_array = ReferenceArray()
            temp_array.Append(references[i]['ref'])
            temp_array.Append(references[i + 1]['ref'])
            
            try:
                dim = doc.Create.NewDimension(view, detail_line, temp_array)
                if dim:
                    dims.append(dim)
            except:
                continue
        
        # Totaal maatvoering
        if add_total and len(references) > 2:
            total_array = ReferenceArray()
            total_array.Append(references[0]['ref'])
            total_array.Append(references[-1]['ref'])
            
            try:
                total_dim = doc.Create.NewDimension(view, total_line, total_array)
                if total_dim:
                    dims.append(total_dim)
            except:
                pass
    except:
        pass
    
    return dims


def create_dimensions(view, elements_dict, direction, add_total, wall_options, offsets, placement):
    """Maak maatvoering voor de geselecteerde elementen"""
    created_dims = []
    
    # Horizontale maatvoering (boven en/of onder)
    if direction in ['Horizontaal', 'Beide']:
        refs_h = collect_references(elements_dict, 'horizontal', wall_options, view)
        if refs_h and len(refs_h) >= 2:
            if placement.get('bottom', True):
                dims = create_dimension_line(view, refs_h, 'horizontal', add_total, offsets, placement, 'bottom')
                created_dims.extend(dims)
            if placement.get('top', False):
                dims = create_dimension_line(view, refs_h, 'horizontal', add_total, offsets, placement, 'top')
                created_dims.extend(dims)
        
        # Wanddikte maatvoering
        if wall_options.get('thickness', False) and elements_dict.get('Wanden'):
            thickness_dims = create_wall_thickness_dimensions(
                view, elements_dict['Wanden'], 'horizontal', offsets['thickness'], placement
            )
            created_dims.extend(thickness_dims)
    
    # Verticale maatvoering (links en/of rechts)
    if direction in ['Verticaal', 'Beide']:
        refs_v = collect_references(elements_dict, 'vertical', wall_options, view)
        if refs_v and len(refs_v) >= 2:
            if placement.get('left', True):
                dims = create_dimension_line(view, refs_v, 'vertical', add_total, offsets, placement, 'left')
                created_dims.extend(dims)
            if placement.get('right', False):
                dims = create_dimension_line(view, refs_v, 'vertical', add_total, offsets, placement, 'right')
                created_dims.extend(dims)
        
        # Wanddikte maatvoering
        if wall_options.get('thickness', False) and elements_dict.get('Wanden'):
            thickness_dims = create_wall_thickness_dimensions(
                view, elements_dict['Wanden'], 'vertical', offsets['thickness'], placement
            )
            created_dims.extend(thickness_dims)
    
    return created_dims


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main entry point"""
    
    # Check actieve view
    active_view = doc.ActiveView
    if not active_view:
        forms.alert('Geen actieve view gevonden.', exitscript=True)
    
    # Check view type
    view_type_ok = False
    if hasattr(active_view, 'ViewType'):
        vt = active_view.ViewType
        if vt in [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan, 
                  ViewType.AreaPlan, ViewType.Section, ViewType.Elevation]:
            view_type_ok = True
    
    if not view_type_ok:
        forms.alert('Selecteer een plattegrond, doorsnede of aanzicht view.', exitscript=True)
    
    # Toon dialog
    result, dialog = run_dialog(DimensionOptionsDialog)
    
    if result != DialogResult.OK:
        return
    
    # Haal opties op
    selected_types = dialog.get_selected_types()
    direction = dialog.get_direction()
    add_total = dialog.get_add_total()
    wall_options = dialog.get_wall_options()
    offsets = dialog.get_offsets()
    placement = dialog.get_placement_options()
    
    # Start transaction
    t = Transaction(doc, 'Auto Maatvoering')
    t.Start()
    
    try:
        elements_dict = get_elements_on_view(active_view, selected_types)
        total_elements = sum(len(v) for v in elements_dict.values())
        
        if total_elements == 0:
            t.RollBack()
            forms.alert(
                'Geen elementen gevonden op de actieve view.\n'
                'Controleer of de geselecteerde element types zichtbaar zijn.',
                title='Waarschuwing'
            )
            return
        
        created_dims = create_dimensions(
            active_view, elements_dict, direction, add_total,
            wall_options, offsets, placement
        )
        
        t.Commit()
        
        if created_dims:
            forms.alert(
                '{} maatvoeringslijnen toegevoegd!\n'
                'Gevonden elementen: {}'.format(len(created_dims), total_elements),
                title='Succes'
            )
        else:
            forms.alert(
                'Geen maatvoering gemaakt.\n'
                'Gevonden elementen: {}\n\n'
                'Mogelijke oorzaken:\n'
                '- Elementen staan niet in de juiste richting\n'
                '- Referenties kunnen niet worden opgehaald'.format(total_elements),
                title='Waarschuwing'
            )
    
    except Exception as e:
        t.RollBack()
        import traceback
        forms.alert('Fout tijdens uitvoering:\n{}\n\n{}'.format(str(e), traceback.format_exc()), title='Error')


# Entry point
if __name__ == "__main__":
    main()
