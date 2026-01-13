# -*- coding: utf-8 -*-
"""
Scan2BIM 
=====================
Extraheer punten uit een pointcloud slice en genereer Revit elementen.

Auteur: JMK
Versie: 1.0
"""

# pyRevit imports
from pyrevit import revit, DB, forms, script

# UI Template import
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from ui_template import (
    BaseForm, UIFactory, COLORS, FONTS, DPIScaler,
    run_dialog
)

# .NET imports
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import DialogResult, Application
from System.Drawing import Point, Size, Color

# Revit imports
from Autodesk.Revit.DB import (
    FilteredElementCollector, PointCloudInstance, 
    BoundingBoxXYZ, XYZ, Transform, Outline,
    Floor, Wall, Level, FloorType, WallType,
    CurveLoop, Line, Transaction, BuiltInCategory,
    ElementId, UnitUtils, SpecTypeId, UnitTypeId
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

# Python math
import math
from collections import defaultdict

# =============================================================================
# CONSTANTEN
# =============================================================================

# Conversie feet naar mm
FEET_TO_MM = 304.8
MM_TO_FEET = 1 / FEET_TO_MM

# =============================================================================
# SELECTION FILTER
# =============================================================================

class PointCloudSelectionFilter(ISelectionFilter):
    """Filter voor alleen pointcloud selectie"""
    
    def AllowElement(self, element):
        return isinstance(element, PointCloudInstance)
    
    def AllowReference(self, reference, position):
        return False


# =============================================================================
# POINTCLOUD UTILITIES
# =============================================================================

class PointCloudProcessor:
    """Hulpklasse voor pointcloud verwerking"""
    
    def __init__(self, doc, pointcloud_instance):
        self.doc = doc
        self.pc_instance = pointcloud_instance
        self.transform = pointcloud_instance.GetTransform()
    
    def get_bounding_box(self):
        """Haal de bounding box van de pointcloud op"""
        return self.pc_instance.get_BoundingBox(None)
    
    def extract_points_in_box(self, min_point, max_point, max_points=100000):
        """
        Extraheer punten binnen een bounding box
        """
        print("=== EXTRACT POINTS ===")
        
        try:
            # Import uit PointClouds namespace
            from Autodesk.Revit.DB.PointClouds import PointCloudFilterFactory
            from System.Collections.Generic import List
            
            # Converteer world naar local coordinates
            inverse_transform = self.transform.Inverse
            local_min = inverse_transform.OfPoint(min_point)
            local_max = inverse_transform.OfPoint(max_point)
            
            actual_min = XYZ(
                min(local_min.X, local_max.X),
                min(local_min.Y, local_max.Y),
                min(local_min.Z, local_max.Z)
            )
            actual_max = XYZ(
                max(local_min.X, local_max.X),
                max(local_min.Y, local_max.Y),
                max(local_min.Z, local_max.Z)
            )
            
            print("Local box: ({:.2f},{:.2f},{:.2f}) -> ({:.2f},{:.2f},{:.2f})".format(
                actual_min.X, actual_min.Y, actual_min.Z,
                actual_max.X, actual_max.Y, actual_max.Z))
            
            # Maak 6 planes voor een box filter
            # Elke plane heeft een normal die naar BINNEN wijst
            planes = List[DB.Plane]()
            
            # +X plane (bij min X, normal wijst naar +X)
            planes.Add(DB.Plane.CreateByNormalAndOrigin(XYZ(1, 0, 0), actual_min))
            # -X plane (bij max X, normal wijst naar -X)
            planes.Add(DB.Plane.CreateByNormalAndOrigin(XYZ(-1, 0, 0), actual_max))
            # +Y plane
            planes.Add(DB.Plane.CreateByNormalAndOrigin(XYZ(0, 1, 0), actual_min))
            # -Y plane
            planes.Add(DB.Plane.CreateByNormalAndOrigin(XYZ(0, -1, 0), actual_max))
            # +Z plane
            planes.Add(DB.Plane.CreateByNormalAndOrigin(XYZ(0, 0, 1), actual_min))
            # -Z plane
            planes.Add(DB.Plane.CreateByNormalAndOrigin(XYZ(0, 0, -1), actual_max))
            
            print("6 planes gemaakt voor box filter")
            
            # Maak filter
            box_filter = PointCloudFilterFactory.CreateMultiPlaneFilter(planes)
            print("MultiPlaneFilter gemaakt: {}".format(type(box_filter)))
            
            # Haal punten op
            cloud_points = self.pc_instance.GetPoints(box_filter, 0.1, max_points)
            
            if not cloud_points:
                print("Geen punten terug")
                return []
            
            print("Punten gevonden: {}".format(cloud_points.Count))
            
            if cloud_points.Count == 0:
                return []
            
            # Converteer naar world coordinates
            points = []
            for i, cp in enumerate(cloud_points):
                local_point = XYZ(cp.X, cp.Y, cp.Z)
                world_point = self.transform.OfPoint(local_point)
                points.append(world_point)
                
                if i < 3:
                    print("Punt {}: ({:.0f}, {:.0f}, {:.0f}) mm".format(
                        i, world_point.X * FEET_TO_MM, 
                        world_point.Y * FEET_TO_MM, 
                        world_point.Z * FEET_TO_MM
                    ))
            
            print("Totaal: {} punten".format(len(points)))
            return points
            
        except Exception as e:
            print("Fout: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []
    
    def _extract_points_fallback(self, min_point, max_point, max_points=100000):
        """
        Fallback methode: haal alle punten en filter handmatig
        """
        print("=== FALLBACK METHODE ===")
        print("Box: ({:.2f}, {:.2f}, {:.2f}) -> ({:.2f}, {:.2f}, {:.2f})".format(
            min_point.X, min_point.Y, min_point.Z,
            max_point.X, max_point.Y, max_point.Z
        ))
        
        try:
            cloud_points = None
            
            # Methode 1: GetPoints met bestaande filter van instance
            try:
                existing_filter = self.pc_instance.GetFilter()
                print("Bestaande filter type: {}".format(type(existing_filter)))
                cloud_points = self.pc_instance.GetPoints(existing_filter, 0.1, max_points)
                print("Methode 1 (GetFilter): {} punten".format(cloud_points.Count if cloud_points else 0))
            except Exception as e1:
                print("Methode 1 fout: {}".format(str(e1)))
            
            # Methode 2: Probeer met None
            if cloud_points is None or cloud_points.Count == 0:
                try:
                    cloud_points = self.pc_instance.GetPoints(None, 0.1, max_points)
                    print("Methode 2 (None): {} punten".format(cloud_points.Count if cloud_points else 0))
                except Exception as e2:
                    print("Methode 2 fout: {}".format(str(e2)))
            
            # Methode 3: Probeer met een box filter direct
            if cloud_points is None or cloud_points.Count == 0:
                try:
                    # Maak een simpele box outline
                    from Autodesk.Revit.DB import PointCloudFilter
                    print("Probeer PointCloudFilter base class...")
                except Exception as e3:
                    print("Methode 3 fout: {}".format(str(e3)))
            
            if cloud_points is None:
                print("Geen cloud_points verkregen!")
                return []
            
            print("Totaal punten opgehaald: {}".format(cloud_points.Count))
            
            # Filter handmatig op bounding box
            points = []
            checked = 0
            for cp in cloud_points:
                local_point = XYZ(cp.X, cp.Y, cp.Z)
                world_point = self.transform.OfPoint(local_point)
                
                checked += 1
                if checked <= 5:
                    print("Punt {}: ({:.2f}, {:.2f}, {:.2f})".format(
                        checked, world_point.X * FEET_TO_MM, 
                        world_point.Y * FEET_TO_MM, world_point.Z * FEET_TO_MM
                    ))
                
                # Check of punt binnen box valt
                if (min_point.X <= world_point.X <= max_point.X and
                    min_point.Y <= world_point.Y <= max_point.Y and
                    min_point.Z <= world_point.Z <= max_point.Z):
                    points.append(world_point)
            
            print("Punten binnen box: {} van {}".format(len(points), checked))
            return points
            
        except Exception as e:
            print("Fallback totale fout: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []
    
    def extract_horizontal_slice(self, z_level, thickness, max_points=100000):
        """
        Extraheer punten uit een horizontale slice
        
        Args:
            z_level: Hoogte van de slice (feet)
            thickness: Dikte van de slice (feet)
            max_points: Maximum aantal punten
            
        Returns:
            List van XYZ punten
        """
        bbox = self.get_bounding_box()
        if not bbox:
            return []
        
        half_thick = thickness / 2.0
        min_point = XYZ(bbox.Min.X, bbox.Min.Y, z_level - half_thick)
        max_point = XYZ(bbox.Max.X, bbox.Max.Y, z_level + half_thick)
        
        return self.extract_points_in_box(min_point, max_point, max_points)
    
    def extract_vertical_slice(self, axis, position, thickness, max_points=100000):
        """
        Extraheer punten uit een verticale slice
        
        Args:
            axis: 'X' of 'Y' - richting van de slice
            position: Positie langs de as (feet)
            thickness: Dikte van de slice (feet)
            max_points: Maximum aantal punten
            
        Returns:
            List van XYZ punten
        """
        bbox = self.get_bounding_box()
        if not bbox:
            return []
        
        half_thick = thickness / 2.0
        
        if axis == 'X':
            min_point = XYZ(position - half_thick, bbox.Min.Y, bbox.Min.Z)
            max_point = XYZ(position + half_thick, bbox.Max.Y, bbox.Max.Z)
        else:  # Y axis
            min_point = XYZ(bbox.Min.X, position - half_thick, bbox.Min.Z)
            max_point = XYZ(bbox.Max.X, position + half_thick, bbox.Max.Z)
        
        return self.extract_points_in_box(min_point, max_point, max_points)


# =============================================================================
# GEOMETRY UTILITIES
# =============================================================================

class GeometryUtils:
    """Hulpfuncties voor geometrie berekeningen"""
    
    @staticmethod
    def points_to_2d_outline(points, grid_size=0.5):
        """
        Converteer 3D punten naar een 2D outline (X, Y vlak)
        Gebruikt een grid-based approach voor robuustheid.
        
        Args:
            points: List van XYZ punten
            grid_size: Grid cel grootte in feet
            
        Returns:
            List van XYZ punten die de outline vormen
        """
        if not points:
            return []
        
        # Vind bounds
        min_x = min(p.X for p in points)
        max_x = max(p.X for p in points)
        min_y = min(p.Y for p in points)
        max_y = max(p.Y for p in points)
        
        # Maak grid
        grid = defaultdict(bool)
        for p in points:
            gx = int((p.X - min_x) / grid_size)
            gy = int((p.Y - min_y) / grid_size)
            grid[(gx, gy)] = True
        
        # Vind boundary cells (cellen met minstens 1 lege buur)
        boundary_cells = []
        for (gx, gy) in grid.keys():
            neighbors = [(gx-1, gy), (gx+1, gy), (gx, gy-1), (gx, gy+1)]
            if any(not grid.get(n, False) for n in neighbors):
                boundary_cells.append((gx, gy))
        
        if not boundary_cells:
            # Als geen boundary gevonden, gebruik bounding box
            return [
                XYZ(min_x, min_y, 0),
                XYZ(max_x, min_y, 0),
                XYZ(max_x, max_y, 0),
                XYZ(min_x, max_y, 0)
            ]
        
        # Sorteer boundary cells voor een continue outline (simplified convex hull)
        # Gebruik bounding box corners voor eenvoud
        avg_z = sum(p.Z for p in points) / len(points)
        
        return [
            XYZ(min_x, min_y, avg_z),
            XYZ(max_x, min_y, avg_z),
            XYZ(max_x, max_y, avg_z),
            XYZ(min_x, max_y, avg_z)
        ]
    
    @staticmethod
    def points_to_wall_line(points, axis='X'):
        """
        Converteer punten naar een wand lijn
        
        Args:
            points: List van XYZ punten
            axis: 'X' of 'Y' - richting van de wand
            
        Returns:
            Tuple van (start_point, end_point, base_z, height)
        """
        if not points:
            return None
        
        min_z = min(p.Z for p in points)
        max_z = max(p.Z for p in points)
        height = max_z - min_z
        
        if axis == 'X':
            # Wand loopt langs Y as
            min_y = min(p.Y for p in points)
            max_y = max(p.Y for p in points)
            avg_x = sum(p.X for p in points) / len(points)
            
            start = XYZ(avg_x, min_y, min_z)
            end = XYZ(avg_x, max_y, min_z)
        else:
            # Wand loopt langs X as
            min_x = min(p.X for p in points)
            max_x = max(p.X for p in points)
            avg_y = sum(p.Y for p in points) / len(points)
            
            start = XYZ(min_x, avg_y, min_z)
            end = XYZ(max_x, avg_y, min_z)
        
        return (start, end, min_z, height)
    
    @staticmethod
    def create_rectangle_curveloop(points):
        """
        Maak een CurveLoop van punten (verwacht 4 punten)
        
        Args:
            points: List van 4 XYZ punten
            
        Returns:
            CurveLoop
        """
        if len(points) < 4:
            return None
        
        curves = []
        for i in range(4):
            start = points[i]
            end = points[(i + 1) % 4]
            line = Line.CreateBound(start, end)
            curves.append(line)
        
        loop = CurveLoop()
        for curve in curves:
            loop.Append(curve)
        
        return loop


# =============================================================================
# ELEMENT CREATION
# =============================================================================

class ElementCreator:
    """Hulpklasse voor het maken van Revit elementen"""
    
    def __init__(self, doc):
        self.doc = doc
    
    def get_levels(self):
        """Haal alle levels op"""
        collector = FilteredElementCollector(self.doc)
        levels = collector.OfClass(Level).ToElements()
        return sorted(levels, key=lambda x: x.Elevation)
    
    def get_floor_types(self):
        """Haal alle floor types op"""
        collector = FilteredElementCollector(self.doc)
        floor_types = collector.OfClass(FloorType).ToElements()
        return [ft for ft in floor_types if ft.IsValidObject]
    
    def get_wall_types(self):
        """Haal alle wall types op"""
        collector = FilteredElementCollector(self.doc)
        wall_types = collector.OfClass(WallType).ToElements()
        return [wt for wt in wall_types if wt.IsValidObject]
    
    def get_nearest_level(self, elevation):
        """Vind het dichtstbijzijnde level bij een elevatie"""
        levels = self.get_levels()
        if not levels:
            return None
        
        return min(levels, key=lambda lvl: abs(lvl.Elevation - elevation))
    
    def create_floor(self, outline_points, floor_type_id, level_id):
        """
        Maak een floor van outline punten
        
        Args:
            outline_points: List van XYZ punten (4 corners)
            floor_type_id: ElementId van floor type
            level_id: ElementId van level
            
        Returns:
            Floor element of None
        """
        try:
            # Zet alle punten op dezelfde Z
            avg_z = sum(p.Z for p in outline_points) / len(outline_points)
            flat_points = [XYZ(p.X, p.Y, avg_z) for p in outline_points]
            
            # Maak curve loop
            curve_loop = GeometryUtils.create_rectangle_curveloop(flat_points)
            if not curve_loop:
                return None
            
            # Maak floor (Revit 2022+ API)
            curve_loops = [curve_loop]
            
            # Probeer nieuwe API eerst
            try:
                floor = Floor.Create(
                    self.doc,
                    curve_loops,
                    floor_type_id,
                    level_id
                )
                return floor
            except:
                # Fallback naar oudere API als nodig
                return None
                
        except Exception as e:
            print("Fout bij maken floor: {}".format(str(e)))
            return None
    
    def create_wall(self, start_point, end_point, wall_type_id, level_id, height):
        """
        Maak een wall van start tot eind punt
        
        Args:
            start_point: XYZ start punt
            end_point: XYZ eind punt
            wall_type_id: ElementId van wall type
            level_id: ElementId van level
            height: Hoogte van de wand (feet)
            
        Returns:
            Wall element of None
        """
        try:
            # Maak lijn voor de wand
            line = Line.CreateBound(
                XYZ(start_point.X, start_point.Y, 0),
                XYZ(end_point.X, end_point.Y, 0)
            )
            
            # Maak wall
            wall = Wall.Create(
                self.doc,
                line,
                wall_type_id,
                level_id,
                height,
                0,  # offset
                False,  # flip
                False   # structural
            )
            
            return wall
            
        except Exception as e:
            print("Fout bij maken wall: {}".format(str(e)))
            return None


# =============================================================================
# UI DIALOG
# =============================================================================

class PointCloudSliceDialog(BaseForm):
    """Hoofddialog voor de Pointcloud Slice Tool"""
    
    def __init__(self, doc, uidoc, pointcloud, saved_state=None):
        super(PointCloudSliceDialog, self).__init__(
            "Pointcloud Slice Tool",
            width=500,
            height=480
        )
        
        # Override DPI scaling voor hoogte (4K fix)
        self.Size = Size(DPIScaler.scale(500), 580)
        
        self.doc = doc
        self.uidoc = uidoc
        self.pointcloud = pointcloud
        self.processor = PointCloudProcessor(doc, pointcloud)
        self.creator = ElementCreator(doc)
        
        # Pick state
        self._pick_requested = False
        self._saved_state = saved_state
        
        # Haal pointcloud bounds op
        self.bbox = self.processor.get_bounding_box()
        
        # Data voor UI
        self.levels = self.creator.get_levels()
        self.floor_types = self.creator.get_floor_types()
        self.wall_types = self.creator.get_wall_types()
        
        # Resultaat
        self.result_element = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Bouw de UI op"""
        
        # Titel + Info op zelfde niveau
        self.add_title("Pointcloud Slice Tool")
        
        # Toon bounds info
        if self.bbox:
            min_z_mm = self.bbox.Min.Z * FEET_TO_MM
            max_z_mm = self.bbox.Max.Z * FEET_TO_MM
            bounds_text = "Z-bereik: {:.0f} - {:.0f} mm".format(min_z_mm, max_z_mm)
        else:
            bounds_text = "Bounds niet beschikbaar"
        
        self.add_label(bounds_text, row=1, col=0)
        
        # Slice Type sectie
        self.add_section("Slice Configuratie", row=2)
        
        # Slice type
        self.add_label("Type:", row=3)
        self.slice_type_combo = self.add_combobox(
            ["Horizontaal (Vloer)", "Verticaal X (Wand)", "Verticaal Y (Wand)"],
            row=3
        )
        self.slice_type_combo.SelectedIndexChanged += self._on_slice_type_changed
        
        # Slice positie met Pick button
        self.add_label("Positie (mm):", row=4)
        default_z = ((self.bbox.Min.Z + self.bbox.Max.Z) / 2 * FEET_TO_MM) if self.bbox else 0
        self.position_input = self.add_numeric(
            row=4, min_val=-100000, max_val=100000, default=int(default_z), decimals=0
        )
        
        # Pick button naast positie input
        self.pick_button = UIFactory.create_button("Pick", primary=False, width=60)
        self.pick_button.Location = Point(
            self.position_input.Location.X + self.position_input.Width + DPIScaler.scale(10),
            self.position_input.Location.Y
        )
        self.pick_button.Click += self._on_pick
        self.Controls.Add(self.pick_button)
        
        # Slice dikte
        self.add_label("Dikte (mm):", row=5)
        self.thickness_input = self.add_numeric(
            row=5, min_val=10, max_val=5000, default=100, decimals=0
        )
        
        # Max punten
        self.add_label("Max punten:", row=6)
        self.max_points_input = self.add_numeric(
            row=6, min_val=1000, max_val=1000000, default=50000, decimals=0
        )
        
        # Element sectie
        self.add_section("Element Generatie", row=7)
        
        # Level selectie
        self.add_label("Level:", row=8)
        level_names = ["{} ({:.0f}mm)".format(l.Name, l.Elevation * FEET_TO_MM) for l in self.levels]
        self.level_combo = self.add_combobox(level_names, row=8)
        
        # Floor type
        self.floor_label = self.add_label("Vloer type:", row=9)
        floor_names = [ft.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or "Unnamed" for ft in self.floor_types]
        self.floor_combo = self.add_combobox(floor_names if floor_names else ["Geen types"], row=9)
        
        # Wall type (zelfde positie, toggle visibility)
        self.wall_label = self.add_label("Wand type:", row=9)
        wall_names = [wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or "Unnamed" for wt in self.wall_types]
        self.wall_combo = self.add_combobox(wall_names if wall_names else ["Geen types"], row=9)
        self.wall_label.Visible = False
        self.wall_combo.Visible = False
        
        # Buttons direct (zonder status panel voor compactheid)
        self.add_button_row([
            ("Preview", self._on_preview, False),
            ("Genereer", self._on_generate, True),
            ("Annuleren", self._on_cancel, False)
        ], row=11)
        
        # Herstel opgeslagen state indien aanwezig
        if self._saved_state:
            self._restore_state(self._saved_state)
    
    def _restore_state(self, state):
        """Herstel opgeslagen waarden"""
        if 'slice_type' in state:
            self.slice_type_combo.SelectedIndex = state['slice_type']
        if 'position' in state:
            self.position_input.Value = state['position']
        if 'thickness' in state:
            self.thickness_input.Value = state['thickness']
        if 'max_points' in state:
            self.max_points_input.Value = state['max_points']
        if 'level_idx' in state and state['level_idx'] >= 0:
            self.level_combo.SelectedIndex = state['level_idx']
        if 'floor_idx' in state and state['floor_idx'] >= 0:
            self.floor_combo.SelectedIndex = state['floor_idx']
        if 'wall_idx' in state and state['wall_idx'] >= 0:
            self.wall_combo.SelectedIndex = state['wall_idx']
        
        # Update visibility
        self._on_slice_type_changed(None, None)
    
    def get_state(self):
        """Haal huidige state op voor opslaan"""
        return {
            'slice_type': self.slice_type_combo.SelectedIndex,
            'position': int(self.position_input.Value),
            'thickness': int(self.thickness_input.Value),
            'max_points': int(self.max_points_input.Value),
            'level_idx': self.level_combo.SelectedIndex,
            'floor_idx': self.floor_combo.SelectedIndex,
            'wall_idx': self.wall_combo.SelectedIndex
        }
    
    def _on_slice_type_changed(self, sender, args):
        """Handle slice type wijziging"""
        is_horizontal = self.slice_type_combo.SelectedIndex == 0
        
        # Toggle floor/wall controls
        self.floor_label.Visible = is_horizontal
        self.floor_combo.Visible = is_horizontal
        self.wall_label.Visible = not is_horizontal
        self.wall_combo.Visible = not is_horizontal
        
        # Update positie label
        if is_horizontal:
            # Update default naar gemiddelde Z
            if self.bbox:
                default_z = (self.bbox.Min.Z + self.bbox.Max.Z) / 2 * FEET_TO_MM
                self.position_input.Value = int(default_z)
        else:
            # Update default naar gemiddelde X of Y
            if self.bbox:
                if self.slice_type_combo.SelectedIndex == 1:  # X
                    default_pos = (self.bbox.Min.X + self.bbox.Max.X) / 2 * FEET_TO_MM
                else:  # Y
                    default_pos = (self.bbox.Min.Y + self.bbox.Max.Y) / 2 * FEET_TO_MM
                self.position_input.Value = int(default_pos)
    
    def _on_pick(self, sender, args):
        """Laat gebruiker een punt selecteren in het model"""
        self._pick_requested = True
        self.DialogResult = DialogResult.Retry  # Speciale code voor "pick"
        self.Close()
    
    def _get_slice_params(self):
        """Haal huidige slice parameters op"""
        slice_type = self.slice_type_combo.SelectedIndex
        position_mm = float(self.position_input.Value)
        thickness_mm = float(self.thickness_input.Value)
        max_points = int(self.max_points_input.Value)
        
        # Conversie naar feet
        position_ft = position_mm * MM_TO_FEET
        thickness_ft = thickness_mm * MM_TO_FEET
        
        return {
            'slice_type': slice_type,  # 0=horizontal, 1=vertical X, 2=vertical Y
            'position': position_ft,
            'thickness': thickness_ft,
            'max_points': max_points
        }
    
    def _extract_points(self, params):
        """Extraheer punten gebaseerd op parameters"""
        if params['slice_type'] == 0:
            # Horizontale slice
            return self.processor.extract_horizontal_slice(
                params['position'],
                params['thickness'],
                params['max_points']
            )
        elif params['slice_type'] == 1:
            # Verticale X slice
            return self.processor.extract_vertical_slice(
                'X',
                params['position'],
                params['thickness'],
                params['max_points']
            )
        else:
            # Verticale Y slice
            return self.processor.extract_vertical_slice(
                'Y',
                params['position'],
                params['thickness'],
                params['max_points']
            )
    
    def _on_preview(self, sender, args):
        """Preview de punten extractie"""
        params = self._get_slice_params()
        points = self._extract_points(params)
        
        if points:
            # Toon info over gevonden punten
            min_x = min(p.X for p in points) * FEET_TO_MM
            max_x = max(p.X for p in points) * FEET_TO_MM
            min_y = min(p.Y for p in points) * FEET_TO_MM
            max_y = max(p.Y for p in points) * FEET_TO_MM
            min_z = min(p.Z for p in points) * FEET_TO_MM
            max_z = max(p.Z for p in points) * FEET_TO_MM
            
            msg = "{} punten gevonden!\n\n".format(len(points))
            msg += "X: {:.0f} - {:.0f} mm\n".format(min_x, max_x)
            msg += "Y: {:.0f} - {:.0f} mm\n".format(min_y, max_y)
            msg += "Z: {:.0f} - {:.0f} mm".format(min_z, max_z)
            
            self.show_info(msg, "Preview Resultaat")
        else:
            self.show_warning(
                "Geen punten gevonden in de opgegeven slice.\n\n"
                "Probeer:\n"
                "- Andere positie\n"
                "- Grotere dikte\n"
                "- Meer max punten",
                "Geen Punten"
            )
    
    def _on_generate(self, sender, args):
        """Genereer het Revit element"""
        params = self._get_slice_params()
        points = self._extract_points(params)
        
        if not points:
            self.show_warning("Geen punten gevonden. Pas de slice parameters aan.")
            return
        
        # Haal level
        if self.level_combo.SelectedIndex < 0 or self.level_combo.SelectedIndex >= len(self.levels):
            self.show_warning("Selecteer een level.")
            return
        
        level = self.levels[self.level_combo.SelectedIndex]
        
        # Genereer element
        try:
            with Transaction(self.doc, "Pointcloud Slice Element") as t:
                t.Start()
                
                if params['slice_type'] == 0:
                    # Horizontale slice -> Floor
                    if self.floor_combo.SelectedIndex < 0 or self.floor_combo.SelectedIndex >= len(self.floor_types):
                        self.show_warning("Selecteer een vloer type.")
                        t.RollBack()
                        return
                    
                    floor_type = self.floor_types[self.floor_combo.SelectedIndex]
                    outline = GeometryUtils.points_to_2d_outline(points)
                    
                    self.result_element = self.creator.create_floor(
                        outline,
                        floor_type.Id,
                        level.Id
                    )
                else:
                    # Verticale slice -> Wall
                    if self.wall_combo.SelectedIndex < 0 or self.wall_combo.SelectedIndex >= len(self.wall_types):
                        self.show_warning("Selecteer een wand type.")
                        t.RollBack()
                        return
                    
                    wall_type = self.wall_types[self.wall_combo.SelectedIndex]
                    axis = 'X' if params['slice_type'] == 1 else 'Y'
                    wall_data = GeometryUtils.points_to_wall_line(points, axis)
                    
                    if wall_data:
                        start, end, base_z, height = wall_data
                        self.result_element = self.creator.create_wall(
                            start,
                            end,
                            wall_type.Id,
                            level.Id,
                            max(height, 2.5)  # Minimaal 2.5 feet hoogte
                        )
                
                if self.result_element:
                    t.Commit()
                    self.show_info(
                        "Element succesvol aangemaakt!\n\n"
                        "Gebaseerd op {} punten.".format(len(points)),
                        "Succes"
                    )
                    self.close_ok()
                else:
                    t.RollBack()
                    self.show_error("Kon element niet maken. Controleer de parameters.")
                    
        except Exception as e:
            self.show_error("Fout bij genereren:\n\n{}".format(str(e)))
    
    def _on_cancel(self, sender, args):
        """Annuleer de dialog"""
        self.close_cancel()


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Hoofdfunctie"""
    doc = revit.doc
    uidoc = revit.uidoc
    
    # Stap 1: Laat gebruiker een pointcloud selecteren
    try:
        # Probeer eerst bestaande selectie
        selection = uidoc.Selection.GetElementIds()
        pointcloud = None
        
        for elem_id in selection:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, PointCloudInstance):
                pointcloud = elem
                break
        
        # Als geen pointcloud geselecteerd, vraag om selectie
        if not pointcloud:
            # Zoek alle pointclouds in het project
            collector = FilteredElementCollector(doc)
            pointclouds = collector.OfClass(PointCloudInstance).ToElements()
            
            if not pointclouds:
                forms.alert(
                    "Geen pointclouds gevonden in dit project.\n\n"
                    "Importeer eerst een pointcloud via:\n"
                    "Insert > Point Cloud",
                    title="Geen Pointcloud"
                )
                return
            
            # Als er maar 1 is, gebruik die
            if len(pointclouds) == 1:
                pointcloud = pointclouds[0]
            else:
                # Laat gebruiker kiezen
                pc_names = [pc.Name or "Pointcloud {}".format(pc.Id.IntegerValue) for pc in pointclouds]
                selected = forms.SelectFromList.show(
                    pc_names,
                    title="Selecteer Pointcloud",
                    button_name="Selecteer"
                )
                
                if not selected:
                    return
                
                idx = pc_names.index(selected)
                pointcloud = pointclouds[idx]
        
        # Dialog loop (voor pick functionaliteit)
        saved_state = None
        
        while True:
            # Open dialog
            dialog = PointCloudSliceDialog(doc, uidoc, pointcloud, saved_state)
            result = dialog.ShowDialog()
            
            # Check resultaat
            if result == DialogResult.Retry:
                # Pick point requested
                saved_state = dialog.get_state()
                slice_type = saved_state['slice_type']
                
                # Bepaal prompt
                if slice_type == 0:
                    prompt = "Klik op een punt om de Z-hoogte te bepalen (ESC om te annuleren)"
                elif slice_type == 1:
                    prompt = "Klik op een punt om de X-positie te bepalen (ESC om te annuleren)"
                else:
                    prompt = "Klik op een punt om de Y-positie te bepalen (ESC om te annuleren)"
                
                try:
                    # Pick point
                    picked_point = uidoc.Selection.PickPoint(prompt)
                    
                    # Bepaal de relevante coordinaat
                    if slice_type == 0:  # Horizontaal -> Z
                        value_mm = picked_point.Z * FEET_TO_MM
                    elif slice_type == 1:  # Verticaal X
                        value_mm = picked_point.X * FEET_TO_MM
                    else:  # Verticaal Y
                        value_mm = picked_point.Y * FEET_TO_MM
                    
                    # Update saved state met nieuwe positie
                    saved_state['position'] = int(value_mm)
                    
                except:
                    # ESC of fout - ga terug naar dialog met oude waarden
                    pass
                
                # Loop door naar volgende iteratie (dialog opent opnieuw)
                continue
            else:
                # OK of Cancel - stop de loop
                break
        
    except Exception as e:
        forms.alert(
            "Fout bij uitvoeren tool:\n\n{}".format(str(e)),
            title="Fout"
        )


# Uitvoeren
if __name__ == '__main__':
    main()
