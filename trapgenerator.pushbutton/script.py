# -*- coding: utf-8 -*-
"""Parametrische Trap Generator
Maakt een parametrische trap als Generic Model met DirectShape geometrie.
Inclusief selecteerbare trapbomen en leuningen met uitgebreide opties.
"""

__title__ = "Trap\nGenerator"
__author__ = "Custom Tool"

import os
from pyrevit import forms, revit, DB, script
from System.Collections.Generic import List
import System.Windows.Media
import System.Windows
import math

doc = revit.doc
uidoc = revit.uidoc

def mm_to_feet(mm):
    return mm / 304.8


class TrapGeneratorWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'TrapGeneratorWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        self.result = None
        self.setup_ui()
        
    def setup_ui(self):
        # Trap afmetingen
        self.hoogte_input.Text = "2800"
        self.breedte_input.Text = "1000"
        self.lengte_input.Text = "3000"
        self.optrede_input.Text = "175"
        self.aantrede_input.Text = "250"
        
        # Trap type
        self.trap_type.ItemsSource = ["Rechte trap", "L-trap", "U-trap", "Spiltrap rechtsom", "Spiltrap linksom"]
        self.trap_type.SelectedIndex = 0
        
        # Materiaal trap
        self.materiaal_type.ItemsSource = ["Beton", "Hout", "Staal"]
        self.materiaal_type.SelectedIndex = 0
        
        # Open/Dicht
        self.trap_style.ItemsSource = ["Dichte trap", "Open trap"]
        self.trap_style.SelectedIndex = 0
        
        # Draairichting voor L-trap en U-trap
        self.draairichting.ItemsSource = ["Rechtsom", "Linksom"]
        self.draairichting.SelectedIndex = 0
        
        # U-trap specifieke velden
        self.utrap_draairichting.ItemsSource = ["Rechtsom", "Linksom"]
        self.utrap_draairichting.SelectedIndex = 0
        self.utrap_bordes_diepte.Text = "1000"
        self.utrap_tussenruimte.Text = "200"
        
        # Trapboom opties
        self.trapboom_breedte.Text = "50"
        self.trapboom_hoogte.Text = "200"
        self.trapboom_materiaal.ItemsSource = ["Hout", "Staal", "RVS", "Aluminium", "Beton"]
        self.trapboom_materiaal.SelectedIndex = 0
        
        # Leuning opties
        self.leuning_vorm.ItemsSource = ["Rond", "Vierkant"]
        self.leuning_vorm.SelectedIndex = 0
        self.leuning_afmeting.Text = "50"
        self.leuning_materiaal.ItemsSource = ["Hout", "Staal", "RVS", "Aluminium"]
        self.leuning_materiaal.SelectedIndex = 0
        self.leuning_hoogte.Text = "900"
        
        # Baluster opties
        self.baluster_vorm.ItemsSource = ["Rond", "Vierkant"]
        self.baluster_vorm.SelectedIndex = 0
        self.baluster_afmeting.Text = "25"
        self.baluster_materiaal.ItemsSource = ["Hout", "Staal", "RVS", "Aluminium"]
        self.baluster_materiaal.SelectedIndex = 0
        self.baluster_hoh.Text = "120"
        self.baluster_offset_start.Text = "50"
        self.baluster_offset_eind.Text = "50"
        
        self.bereken_treden()
        self.update_leuning_visibility()
        self.update_trapboom_visibility()
        self.update_bordes_visibility()
        
    def bereken_treden(self):
        try:
            hoogte = float(self.hoogte_input.Text)
            optrede = float(self.optrede_input.Text)
            aantal = int(math.ceil(hoogte / optrede))
            werkelijke_optrede = hoogte / aantal
            
            self.aantal_treden.Text = "Aantal treden: {0}".format(aantal)
            self.werkelijke_optrede.Text = "Werkelijke optrede: {0:.1f} mm".format(werkelijke_optrede)
            
            aantrede = float(self.aantrede_input.Text)
            blondel = 2 * werkelijke_optrede + aantrede
            
            if 590 <= blondel <= 650:
                self.blondel_check.Text = "Stapmodulus: {0:.0f} mm (OK)".format(blondel)
                self.blondel_check.Foreground = System.Windows.Media.Brushes.Green
            else:
                self.blondel_check.Text = "Stapmodulus: {0:.0f} mm (Optimaal: 590-650)".format(blondel)
                self.blondel_check.Foreground = System.Windows.Media.Brushes.Orange
                
        except (ValueError, AttributeError):
            self.aantal_treden.Text = "Voer geldige waarden in"
            self.werkelijke_optrede.Text = ""
            self.blondel_check.Text = ""
    
    def on_input_changed(self, sender, args):
        self.bereken_treden()
    
    def on_leuning_changed(self, sender, args):
        self.update_leuning_visibility()
    
    def on_trapboom_changed(self, sender, args):
        self.update_trapboom_visibility()
    
    def on_trap_type_changed(self, sender, args):
        self.update_bordes_visibility()
        self.bereken_treden()
    
    def toon_bouwbesluit_info(self, sender, args):
        """Toon informatie over Bouwbesluit eisen voor trappen"""
        info_tekst = """BOUWBESLUIT 2012 - EISEN VOOR TRAPPEN

═══════════════════════════════════════════
ALGEMEEN (Afdeling 2.5)
═══════════════════════════════════════════

▸ Trap verplicht bij hoogteverschil > 210mm
▸ Trap moet over de volle breedte vrij zijn van obstakels

═══════════════════════════════════════════
AFMETINGEN TREDEN
═══════════════════════════════════════════

WONINGEN (Woonfunctie):
▸ Optrede (max): 220 mm
▸ Aantrede (min): 150 mm
▸ Breedte (min): 800 mm
▸ Vrije hoogte (min): 2300 mm

UTILITEIT (Overige functies):
▸ Optrede (max): 185 mm
▸ Aantrede (min): 230 mm
▸ Breedte (min): 1200 mm
▸ Vrije hoogte (min): 2300 mm

GEMEENSCHAPPELIJK (bijv. portiek):
▸ Optrede (max): 185 mm
▸ Aantrede (min): 230 mm
▸ Breedte (min): 1200 mm

═══════════════════════════════════════════
STAPMODULUS (2× optrede + aantrede)
═══════════════════════════════════════════

▸ Optimaal bereik: 590 - 650 mm
▸ Ideaal (comfortabel): 630 mm
▸ Formule: 2×O + A = 590-650 mm

═══════════════════════════════════════════
LEUNINGEN & BALUSTRADES (Afdeling 2.4)
═══════════════════════════════════════════

▸ Leuning verplicht bij hoogteverschil > 1000 mm
▸ Hoogte leuning (min): 850 mm (woningen)
▸ Hoogte leuning (min): 1000 mm (overig)
▸ Openingen max 100 mm (niet beklimbaar voor kinderen)

═══════════════════════════════════════════
OVERIGE EISEN
═══════════════════════════════════════════

▸ Klimlijn mag niet door het tredevlak lopen
▸ Alle treden moeten gelijke afmetingen hebben
▸ Max 18 treden per trap zonder bordes
▸ Bordes min 800 mm diep (woningen)
▸ Bordes min 1200 mm diep (utiliteit)
"""
        forms.alert(info_tekst, title="Bouwbesluit - Trappen")
    
    def update_leuning_visibility(self):
        """Toon/verberg leuning details op basis van checkboxes"""
        heeft_leuning = self.leuning_links.IsChecked or self.leuning_rechts.IsChecked
        self.leuning_details.Visibility = System.Windows.Visibility.Visible if heeft_leuning else System.Windows.Visibility.Collapsed
    
    def update_trapboom_visibility(self):
        """Toon/verberg trapboom details op basis van checkboxes"""
        heeft_trapboom = self.trapboom_links.IsChecked or self.trapboom_rechts.IsChecked
        self.trapboom_details.Visibility = System.Windows.Visibility.Visible if heeft_trapboom else System.Windows.Visibility.Collapsed
    
    def update_bordes_visibility(self):
        """Toon/verberg bordes optie voor L-trap en U-trap"""
        trap_type = str(self.trap_type.SelectedItem) if self.trap_type.SelectedItem else ""
        is_l_trap = trap_type == "L-trap"
        is_u_trap = trap_type == "U-trap"
        
        # L-trap panel
        self.bordes_panel.Visibility = System.Windows.Visibility.Visible if is_l_trap else System.Windows.Visibility.Collapsed
        
        # U-trap panel
        self.utrap_panel.Visibility = System.Windows.Visibility.Visible if is_u_trap else System.Windows.Visibility.Collapsed
        
        # Zet standaard waarde op helft van aantal treden
        try:
            hoogte = float(self.hoogte_input.Text)
            optrede = float(self.optrede_input.Text)
            aantal = int(math.ceil(hoogte / optrede))
            default_bordes = str(aantal // 2)
        except:
            default_bordes = "8"
        
        if is_l_trap and not self.bordes_na_trede.Text:
            self.bordes_na_trede.Text = default_bordes
        if is_u_trap and not self.utrap_bordes_na_trede.Text:
            self.utrap_bordes_na_trede.Text = default_bordes
    
    def maak_trap(self, sender, args):
        try:
            hoogte = float(self.hoogte_input.Text)
            breedte = float(self.breedte_input.Text)
            lengte = float(self.lengte_input.Text)
            optrede = float(self.optrede_input.Text)
            aantrede = float(self.aantrede_input.Text)
            
            trap_type = str(self.trap_type.SelectedItem)
            materiaal = str(self.materiaal_type.SelectedItem)
            is_open = self.trap_style.SelectedIndex == 1
            
            # Trapboom opties
            trapboom_links = self.trapboom_links.IsChecked
            trapboom_rechts = self.trapboom_rechts.IsChecked
            trapboom_breedte = float(self.trapboom_breedte.Text)
            trapboom_hoogte = float(self.trapboom_hoogte.Text)
            trapboom_materiaal = str(self.trapboom_materiaal.SelectedItem)
            
            # Leuning opties
            leuning_links = self.leuning_links.IsChecked
            leuning_rechts = self.leuning_rechts.IsChecked
            
            # Balusters per zijde
            balusters_links = self.balusters_links.IsChecked
            balusters_rechts = self.balusters_rechts.IsChecked
            
            # Leuning parameters
            leuning_vorm = str(self.leuning_vorm.SelectedItem)
            leuning_afmeting = float(self.leuning_afmeting.Text)
            leuning_materiaal = str(self.leuning_materiaal.SelectedItem)
            leuning_hoogte = float(self.leuning_hoogte.Text)
            
            # Baluster parameters
            baluster_vorm = str(self.baluster_vorm.SelectedItem)
            baluster_afmeting = float(self.baluster_afmeting.Text)
            baluster_materiaal = str(self.baluster_materiaal.SelectedItem)
            baluster_hoh = float(self.baluster_hoh.Text)
            baluster_offset_start = float(self.baluster_offset_start.Text)
            baluster_offset_eind = float(self.baluster_offset_eind.Text)
            
            aantal_treden = int(math.ceil(hoogte / optrede))
            
            # Bordes positie voor L-trap
            try:
                bordes_na_trede = int(self.bordes_na_trede.Text)
            except:
                bordes_na_trede = aantal_treden // 2
            
            # Draairichting (direct gebruiken zonder inversie)
            draairichting = str(self.draairichting.SelectedItem) if self.draairichting.SelectedItem else "Rechtsom"
            
            # U-trap specifieke parameters
            try:
                utrap_bordes_na_trede = int(self.utrap_bordes_na_trede.Text)
            except:
                utrap_bordes_na_trede = aantal_treden // 2
            try:
                utrap_bordes_diepte = float(self.utrap_bordes_diepte.Text)
            except:
                utrap_bordes_diepte = 1000
            try:
                utrap_tussenruimte = float(self.utrap_tussenruimte.Text)
            except:
                utrap_tussenruimte = 200
            utrap_draairichting = str(self.utrap_draairichting.SelectedItem) if self.utrap_draairichting.SelectedItem else "Rechtsom"
            
            # Looplijn optie
            toon_looplijn = self.toon_looplijn.IsChecked
            
            self.result = {
                'hoogte': hoogte,
                'breedte': breedte,
                'lengte': lengte,
                'optrede': optrede,
                'aantrede': aantrede,
                'aantal_treden': aantal_treden,
                'trap_type': trap_type,
                'materiaal': materiaal,
                'is_open': is_open,
                'bordes_na_trede': bordes_na_trede,
                'draairichting': draairichting,
                # U-trap specifiek
                'utrap_bordes_na_trede': utrap_bordes_na_trede,
                'utrap_bordes_diepte': utrap_bordes_diepte,
                'utrap_tussenruimte': utrap_tussenruimte,
                'utrap_draairichting': utrap_draairichting,
                # Looplijn
                'toon_looplijn': toon_looplijn,
                # Trapbomen
                'trapboom_links': trapboom_links,
                'trapboom_rechts': trapboom_rechts,
                'trapboom_breedte': trapboom_breedte,
                'trapboom_hoogte': trapboom_hoogte,
                'trapboom_materiaal': trapboom_materiaal,
                # Leuningen
                'leuning_links': leuning_links,
                'leuning_rechts': leuning_rechts,
                'balusters_links': balusters_links,
                'balusters_rechts': balusters_rechts,
                # Leuning
                'leuning_vorm': leuning_vorm,
                'leuning_afmeting': leuning_afmeting,
                'leuning_materiaal': leuning_materiaal,
                'leuning_hoogte': leuning_hoogte,
                # Baluster
                'baluster_vorm': baluster_vorm,
                'baluster_afmeting': baluster_afmeting,
                'baluster_materiaal': baluster_materiaal,
                'baluster_hoh': baluster_hoh,
                'baluster_offset_start': baluster_offset_start,
                'baluster_offset_eind': baluster_offset_eind
            }
            
            self.Close()
                
        except ValueError:
            forms.alert("Voer geldige numerieke waarden in", title="Invoerfout")
    
    def annuleer(self, sender, args):
        self.result = None
        self.Close()


def create_box_solid(origin, length_x, length_y, length_z):
    """Maak een eenvoudige box solid"""
    p0 = origin
    p1 = DB.XYZ(origin.X + length_x, origin.Y, origin.Z)
    p2 = DB.XYZ(origin.X + length_x, origin.Y + length_y, origin.Z)
    p3 = DB.XYZ(origin.X, origin.Y + length_y, origin.Z)
    
    profile = DB.CurveLoop()
    profile.Append(DB.Line.CreateBound(p0, p1))
    profile.Append(DB.Line.CreateBound(p1, p2))
    profile.Append(DB.Line.CreateBound(p2, p3))
    profile.Append(DB.Line.CreateBound(p3, p0))
    
    profiles = List[DB.CurveLoop]()
    profiles.Add(profile)
    
    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        profiles,
        DB.XYZ(0, 0, 1),
        length_z
    )
    
    return solid


def create_trapboom_solid(start_point, end_point, boom_breedte, boom_hoogte):
    """Maak een trapboom (stringer) als schuine balk"""
    direction = end_point.Subtract(start_point)
    length = direction.GetLength()
    
    if length < 0.01:
        return None
    
    dir_norm = direction.Normalize()
    
    half_b = boom_breedte / 2
    half_h = boom_hoogte / 2
    
    up = DB.XYZ(0, 0, 1)
    right = dir_norm.CrossProduct(up).Normalize()
    if right.GetLength() < 0.01:
        right = DB.XYZ(0, 1, 0)
    
    p0 = start_point.Add(right.Multiply(-half_b)).Add(DB.XYZ(0, 0, -half_h))
    p1 = start_point.Add(right.Multiply(half_b)).Add(DB.XYZ(0, 0, -half_h))
    p2 = start_point.Add(right.Multiply(half_b)).Add(DB.XYZ(0, 0, half_h))
    p3 = start_point.Add(right.Multiply(-half_b)).Add(DB.XYZ(0, 0, half_h))
    
    profile = DB.CurveLoop()
    profile.Append(DB.Line.CreateBound(p0, p1))
    profile.Append(DB.Line.CreateBound(p1, p2))
    profile.Append(DB.Line.CreateBound(p2, p3))
    profile.Append(DB.Line.CreateBound(p3, p0))
    
    profiles = List[DB.CurveLoop]()
    profiles.Add(profile)
    
    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        profiles,
        dir_norm,
        length
    )
    
    return solid


def create_leuning_solid(start_point, end_point, afmeting, vorm="Rond"):
    """Maak een leuning als schuine balk (rond of vierkant)"""
    direction = end_point.Subtract(start_point)
    length = direction.GetLength()
    
    if length < 0.01:
        return None
    
    dir_norm = direction.Normalize()
    
    up = DB.XYZ(0, 0, 1)
    right = dir_norm.CrossProduct(up).Normalize()
    if right.GetLength() < 0.01:
        right = DB.XYZ(0, 1, 0)
    
    if vorm == "Rond":
        profile = DB.CurveLoop()
        segments = 12
        radius = afmeting / 2
        
        local_up = right.CrossProduct(dir_norm).Normalize()
        
        points = []
        for i in range(segments):
            angle = 2 * math.pi * i / segments
            offset = right.Multiply(radius * math.cos(angle)).Add(local_up.Multiply(radius * math.sin(angle)))
            points.append(start_point.Add(offset))
        
        for i in range(segments):
            p1 = points[i]
            p2 = points[(i + 1) % segments]
            profile.Append(DB.Line.CreateBound(p1, p2))
    else:
        half = afmeting / 2
        local_up = right.CrossProduct(dir_norm).Normalize()
        
        p0 = start_point.Add(right.Multiply(-half)).Add(local_up.Multiply(-half))
        p1 = start_point.Add(right.Multiply(half)).Add(local_up.Multiply(-half))
        p2 = start_point.Add(right.Multiply(half)).Add(local_up.Multiply(half))
        p3 = start_point.Add(right.Multiply(-half)).Add(local_up.Multiply(half))
        
        profile = DB.CurveLoop()
        profile.Append(DB.Line.CreateBound(p0, p1))
        profile.Append(DB.Line.CreateBound(p1, p2))
        profile.Append(DB.Line.CreateBound(p2, p3))
        profile.Append(DB.Line.CreateBound(p3, p0))
    
    profiles = List[DB.CurveLoop]()
    profiles.Add(profile)
    
    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        profiles,
        dir_norm,
        length
    )
    
    return solid


def create_baluster(origin, afmeting, hoogte, vorm="Rond"):
    """Maak een baluster (verticale stijl) - rond of vierkant"""
    if vorm == "Rond":
        profile = DB.CurveLoop()
        segments = 12
        radius = afmeting / 2
        
        points = []
        for i in range(segments):
            angle = 2 * math.pi * i / segments
            x = origin.X + radius * math.cos(angle)
            y = origin.Y + radius * math.sin(angle)
            points.append(DB.XYZ(x, y, origin.Z))
        
        for i in range(segments):
            p1 = points[i]
            p2 = points[(i + 1) % segments]
            profile.Append(DB.Line.CreateBound(p1, p2))
    else:
        half = afmeting / 2
        p0 = DB.XYZ(origin.X - half, origin.Y - half, origin.Z)
        p1 = DB.XYZ(origin.X + half, origin.Y - half, origin.Z)
        p2 = DB.XYZ(origin.X + half, origin.Y + half, origin.Z)
        p3 = DB.XYZ(origin.X - half, origin.Y + half, origin.Z)
        
        profile = DB.CurveLoop()
        profile.Append(DB.Line.CreateBound(p0, p1))
        profile.Append(DB.Line.CreateBound(p1, p2))
        profile.Append(DB.Line.CreateBound(p2, p3))
        profile.Append(DB.Line.CreateBound(p3, p0))
    
    profiles = List[DB.CurveLoop]()
    profiles.Add(profile)
    
    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        profiles,
        DB.XYZ(0, 0, 1),
        hoogte
    )
    
    return solid


def create_balusters_along_run(start_x, start_z, end_x, end_z, leuning_y, params):
    """Maak balusters langs een run met offset (Run 1 - langs X-as)"""
    solids = []
    
    baluster_vorm = params.get('baluster_vorm', 'Rond')
    baluster_afmeting = mm_to_feet(params.get('baluster_afmeting', 25))
    baluster_hoh = mm_to_feet(params.get('baluster_hoh', 120))
    leuning_hoogte = mm_to_feet(params.get('leuning_hoogte', 900))
    offset_start = mm_to_feet(params.get('baluster_offset_start', 50))
    offset_eind = mm_to_feet(params.get('baluster_offset_eind', 50))
    
    dx = end_x - start_x
    dz = end_z - start_z
    trap_lengte = math.sqrt(dx**2 + dz**2)
    
    effectieve_lengte = trap_lengte - offset_start - offset_eind
    
    if effectieve_lengte <= 0:
        return solids
    
    num_balusters = max(2, int(effectieve_lengte / baluster_hoh) + 1)
    
    for i in range(num_balusters):
        if num_balusters > 1:
            t_local = float(i) / (num_balusters - 1)
        else:
            t_local = 0.5
        
        pos_along = offset_start + t_local * effectieve_lengte
        t = pos_along / trap_lengte if trap_lengte > 0 else 0
        
        pos_x = start_x + t * dx
        pos_z = start_z + t * dz
        
        bal = create_baluster(
            DB.XYZ(pos_x, leuning_y, pos_z),
            baluster_afmeting,
            leuning_hoogte,
            baluster_vorm
        )
        if bal:
            solids.append(bal)
    
    return solids


def create_balusters_along_run2(start_y, start_z, end_y, end_z, leuning_x, params):
    """Maak balusters langs Run 2 (langs Y-as met hoogteverschil)"""
    solids = []
    
    baluster_vorm = params.get('baluster_vorm', 'Rond')
    baluster_afmeting = mm_to_feet(params.get('baluster_afmeting', 25))
    baluster_hoh = mm_to_feet(params.get('baluster_hoh', 120))
    leuning_hoogte = mm_to_feet(params.get('leuning_hoogte', 900))
    offset_start = mm_to_feet(params.get('baluster_offset_start', 50))
    offset_eind = mm_to_feet(params.get('baluster_offset_eind', 50))
    
    dy = end_y - start_y
    dz = end_z - start_z
    trap_lengte = math.sqrt(dy**2 + dz**2)
    
    effectieve_lengte = trap_lengte - offset_start - offset_eind
    
    if effectieve_lengte <= 0:
        return solids
    
    num_balusters = max(2, int(effectieve_lengte / baluster_hoh) + 1)
    
    for i in range(num_balusters):
        if num_balusters > 1:
            t_local = float(i) / (num_balusters - 1)
        else:
            t_local = 0.5
        
        pos_along = offset_start + t_local * effectieve_lengte
        t = pos_along / trap_lengte if trap_lengte > 0 else 0
        
        pos_y = start_y + t * dy
        pos_z = start_z + t * dz
        
        bal = create_baluster(
            DB.XYZ(leuning_x, pos_y, pos_z),
            baluster_afmeting,
            leuning_hoogte,
            baluster_vorm
        )
        if bal:
            solids.append(bal)
    
    return solids


def create_balusters_along_bordes(start_x, start_y, end_x, end_y, z_pos, params):
    """Maak balusters langs het bordes (horizontaal, geen hoogteverschil)"""
    solids = []
    
    baluster_vorm = params.get('baluster_vorm', 'Rond')
    baluster_afmeting = mm_to_feet(params.get('baluster_afmeting', 25))
    baluster_hoh = mm_to_feet(params.get('baluster_hoh', 120))
    leuning_hoogte = mm_to_feet(params.get('leuning_hoogte', 900))
    offset_start = mm_to_feet(params.get('baluster_offset_start', 50))
    offset_eind = mm_to_feet(params.get('baluster_offset_eind', 50))
    
    dx = end_x - start_x
    dy = end_y - start_y
    bordes_lengte = math.sqrt(dx**2 + dy**2)
    
    effectieve_lengte = bordes_lengte - offset_start - offset_eind
    
    if effectieve_lengte <= 0:
        return solids
    
    num_balusters = max(2, int(effectieve_lengte / baluster_hoh) + 1)
    
    for i in range(num_balusters):
        if num_balusters > 1:
            t_local = float(i) / (num_balusters - 1)
        else:
            t_local = 0.5
        
        pos_along = offset_start + t_local * effectieve_lengte
        t = pos_along / bordes_lengte if bordes_lengte > 0 else 0
        
        pos_x = start_x + t * dx
        pos_y = start_y + t * dy
        
        bal = create_baluster(
            DB.XYZ(pos_x, pos_y, z_pos),
            baluster_afmeting,
            leuning_hoogte,
            baluster_vorm
        )
        if bal:
            solids.append(bal)
    
    return solids


def create_pie_trede(center, inner_radius, outer_radius, start_angle, end_angle, z_base, thickness):
    """Maak een taartpuntvormige trede voor spiltrap"""
    segments = 8  # Aantal segmenten voor de boog
    
    # Maak profiel punten
    points = []
    
    # Binnenste boog (van start naar eind)
    for i in range(segments + 1):
        t = float(i) / segments
        angle = start_angle + t * (end_angle - start_angle)
        x = center.X + inner_radius * math.cos(angle)
        y = center.Y + inner_radius * math.sin(angle)
        points.append(DB.XYZ(x, y, z_base))
    
    # Buitenste boog (van eind naar start)
    for i in range(segments, -1, -1):
        t = float(i) / segments
        angle = start_angle + t * (end_angle - start_angle)
        x = center.X + outer_radius * math.cos(angle)
        y = center.Y + outer_radius * math.sin(angle)
        points.append(DB.XYZ(x, y, z_base))
    
    # Maak CurveLoop
    profile = DB.CurveLoop()
    for i in range(len(points)):
        p1 = points[i]
        p2 = points[(i + 1) % len(points)]
        if p1.DistanceTo(p2) > 0.001:  # Voorkom te korte lijnen
            profile.Append(DB.Line.CreateBound(p1, p2))
    
    profiles = List[DB.CurveLoop]()
    profiles.Add(profile)
    
    try:
        solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            profiles,
            DB.XYZ(0, 0, 1),
            thickness
        )
        return solid
    except:
        return None


def create_trap_geometry(params, start_point):
    """Maak de trap geometrie als lijst van solids"""
    solids = []
    
    # Parameters in feet
    totale_breedte = mm_to_feet(params['breedte'])
    aantrede = mm_to_feet(params['aantrede'])
    optrede = mm_to_feet(params['hoogte'] / params['aantal_treden'])
    trede_dikte = mm_to_feet(50)
    stootbord_dikte = mm_to_feet(20)
    aantal = params['aantal_treden']
    is_open = params['is_open']
    trap_type = params['trap_type']
    
    # Trapboom opties
    trapboom_links = params.get('trapboom_links', True)
    trapboom_rechts = params.get('trapboom_rechts', True)
    boom_breedte = mm_to_feet(params.get('trapboom_breedte', 50))
    boom_hoogte = mm_to_feet(params.get('trapboom_hoogte', 200))
    
    # Leuning opties
    leuning_links = params.get('leuning_links', True)
    leuning_rechts = params.get('leuning_rechts', True)
    balusters_links = params.get('balusters_links', True)
    balusters_rechts = params.get('balusters_rechts', True)
    
    # Leuning parameters
    leuning_vorm = params.get('leuning_vorm', 'Rond')
    leuning_afmeting = mm_to_feet(params.get('leuning_afmeting', 50))
    leuning_hoogte = mm_to_feet(params.get('leuning_hoogte', 900))
    
    # Bereken of er trapbomen zijn en pas trede breedte aan
    heeft_trapboom_links = trapboom_links
    heeft_trapboom_rechts = trapboom_rechts
    
    # Trede offset en breedte berekenen
    trede_y_offset = boom_breedte if heeft_trapboom_links else 0
    trede_breedte_aftrek = 0
    if heeft_trapboom_links:
        trede_breedte_aftrek += boom_breedte
    if heeft_trapboom_rechts:
        trede_breedte_aftrek += boom_breedte
    trede_breedte = totale_breedte - trede_breedte_aftrek
    
    # ============ RECHTE TRAP ============
    if trap_type == "Rechte trap":
        # Treden en stootborden
        for i in range(aantal):
            trede_x = start_point.X + (i * aantrede)
            trede_z = start_point.Z + ((i + 1) * optrede)
            
            trede = create_box_solid(
                DB.XYZ(trede_x, start_point.Y + trede_y_offset, trede_z - trede_dikte),
                aantrede, trede_breedte, trede_dikte
            )
            if trede:
                solids.append(trede)
            
            if not is_open:
                stootbord = create_box_solid(
                    DB.XYZ(trede_x, start_point.Y + trede_y_offset, start_point.Z + (i * optrede)),
                    stootbord_dikte, trede_breedte, optrede
                )
                if stootbord:
                    solids.append(stootbord)
        
        # Trapboom LINKS
        if trapboom_links:
            boom_left = create_trapboom_solid(
                DB.XYZ(start_point.X, start_point.Y + boom_breedte/2, start_point.Z + boom_hoogte/2),
                DB.XYZ(start_point.X + aantal * aantrede, start_point.Y + boom_breedte/2, start_point.Z + aantal * optrede + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom_left:
                solids.append(boom_left)
        
        # Trapboom RECHTS
        if trapboom_rechts:
            boom_right = create_trapboom_solid(
                DB.XYZ(start_point.X, start_point.Y + totale_breedte - boom_breedte/2, start_point.Z + boom_hoogte/2),
                DB.XYZ(start_point.X + aantal * aantrede, start_point.Y + totale_breedte - boom_breedte/2, start_point.Z + aantal * optrede + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom_right:
                solids.append(boom_right)
        
        # Leuning LINKS (vanuit loopperspectief = hoge Y kant)
        if leuning_links:
            if trapboom_rechts:
                leuning_y = start_point.Y + totale_breedte - boom_breedte/2
            else:
                leuning_y = start_point.Y + totale_breedte
            
            leuning_l = create_leuning_solid(
                DB.XYZ(start_point.X, leuning_y, start_point.Z + leuning_hoogte),
                DB.XYZ(start_point.X + aantal * aantrede, leuning_y, start_point.Z + aantal * optrede + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning_l:
                solids.append(leuning_l)
            
            # Balusters links
            if balusters_links:
                balusters = create_balusters_along_run(
                    start_point.X, start_point.Z,
                    start_point.X + aantal * aantrede, start_point.Z + aantal * optrede,
                    leuning_y, params
                )
                solids.extend(balusters)
        
        # Leuning RECHTS (vanuit loopperspectief = lage Y kant)
        if leuning_rechts:
            if trapboom_links:
                leuning_y = start_point.Y + boom_breedte/2
            else:
                leuning_y = start_point.Y
            
            leuning_r = create_leuning_solid(
                DB.XYZ(start_point.X, leuning_y, start_point.Z + leuning_hoogte),
                DB.XYZ(start_point.X + aantal * aantrede, leuning_y, start_point.Z + aantal * optrede + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning_r:
                solids.append(leuning_r)
            
            # Balusters rechts
            if balusters_rechts:
                balusters = create_balusters_along_run(
                    start_point.X, start_point.Z,
                    start_point.X + aantal * aantrede, start_point.Z + aantal * optrede,
                    leuning_y, params
                )
                solids.extend(balusters)
    
    # ============ L-TRAP ============
    elif trap_type == "L-trap":
        bordes_na_trede = params.get('bordes_na_trede', aantal // 2)
        # Begrens de waarde
        bordes_na_trede = max(1, min(bordes_na_trede, aantal - 1))
        
        draairichting = params.get('draairichting', 'Rechtsom')
        # BELANGRIJK: Vanuit het perspectief van iemand die de trap OPLOOPT (in +X richting):
        # - Rechtsom = draai naar rechts = Run 2 gaat naar -Y
        # - Linksom = draai naar links = Run 2 gaat naar +Y
        is_rechtsom = draairichting == "Rechtsom"
        
        treden_run1 = bordes_na_trede
        treden_run2 = aantal - treden_run1
        
        # Run 1 - treden
        for i in range(treden_run1):
            trede_x = start_point.X + (i * aantrede)
            trede_z = start_point.Z + ((i + 1) * optrede)
            
            trede = create_box_solid(
                DB.XYZ(trede_x, start_point.Y + trede_y_offset, trede_z - trede_dikte),
                aantrede, trede_breedte, trede_dikte
            )
            if trede:
                solids.append(trede)
            
            if not is_open:
                stootbord = create_box_solid(
                    DB.XYZ(trede_x, start_point.Y + trede_y_offset, start_point.Z + (i * optrede)),
                    stootbord_dikte, trede_breedte, optrede
                )
                if stootbord:
                    solids.append(stootbord)
        
        # Landing
        landing_x = start_point.X + (treden_run1 * aantrede)
        landing_z = start_point.Z + (treden_run1 * optrede)
        
        # Het bordes moet ALTIJD aansluiten op Run 1 (Y van start_point.Y tot start_point.Y + totale_breedte)
        # Bordes is vierkant (totale_breedte x totale_breedte)
        if is_rechtsom:
            # Rechtsom: bordes naar -Y kant (rechts kijkend in looprichting)
            # Bordes van Y=start_point.Y tot Y=start_point.Y+totale_breedte (zelfde als Run 1)
            landing = create_box_solid(
                DB.XYZ(landing_x, start_point.Y, landing_z - trede_dikte),
                totale_breedte, totale_breedte, trede_dikte
            )
        else:
            # Linksom: bordes naar +Y kant (links kijkend in looprichting)  
            # Bordes van Y=start_point.Y tot Y=start_point.Y+totale_breedte (zelfde als Run 1)
            landing = create_box_solid(
                DB.XYZ(landing_x, start_point.Y, landing_z - trede_dikte),
                totale_breedte, totale_breedte, trede_dikte
            )
        if landing:
            solids.append(landing)
        
        # Run 2 - treden (gedraaid 90 graden)
        # Run 2 begint direct aan de rand van het bordes
        for i in range(treden_run2):
            if is_rechtsom:
                # Rechtsom: Run 2 gaat naar -Y, startend bij Y = start_point.Y
                # Eerste trede begint bij Y = start_point.Y - aantrede
                trede_y = start_point.Y - aantrede - (i * aantrede)
            else:
                # Linksom: Run 2 gaat naar +Y, startend bij Y = start_point.Y + totale_breedte
                trede_y = start_point.Y + totale_breedte + (i * aantrede)
            trede_z = landing_z + ((i + 1) * optrede)
            
            trede = create_box_solid(
                DB.XYZ(landing_x + trede_y_offset, trede_y, trede_z - trede_dikte),
                trede_breedte, aantrede, trede_dikte
            )
            if trede:
                solids.append(trede)
            
            if not is_open:
                if is_rechtsom:
                    stootbord = create_box_solid(
                        DB.XYZ(landing_x + trede_y_offset, trede_y + aantrede, landing_z + (i * optrede)),
                        trede_breedte, stootbord_dikte, optrede
                    )
                else:
                    stootbord = create_box_solid(
                        DB.XYZ(landing_x + trede_y_offset, trede_y, landing_z + (i * optrede)),
                        trede_breedte, stootbord_dikte, optrede
                    )
                if stootbord:
                    solids.append(stootbord)
        
        end_z = landing_z + treden_run2 * optrede
        if is_rechtsom:
            end_y = start_point.Y - treden_run2 * aantrede
        else:
            end_y = start_point.Y + totale_breedte + treden_run2 * aantrede
        
        # Trapbomen Run 1
        if trapboom_links:
            boom1_l = create_trapboom_solid(
                DB.XYZ(start_point.X, start_point.Y + boom_breedte/2, start_point.Z + boom_hoogte/2),
                DB.XYZ(landing_x, start_point.Y + boom_breedte/2, landing_z + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom1_l:
                solids.append(boom1_l)
        
        if trapboom_rechts:
            boom1_r = create_trapboom_solid(
                DB.XYZ(start_point.X, start_point.Y + totale_breedte - boom_breedte/2, start_point.Z + boom_hoogte/2),
                DB.XYZ(landing_x, start_point.Y + totale_breedte - boom_breedte/2, landing_z + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom1_r:
                solids.append(boom1_r)
        
        # Trapbomen Run 2
        if is_rechtsom:
            # Rechtsom: Run 2 gaat naar -Y, start bij Y = start_point.Y
            if trapboom_links:
                boom2_l = create_trapboom_solid(
                    DB.XYZ(landing_x + boom_breedte/2, start_point.Y, landing_z + boom_hoogte/2),
                    DB.XYZ(landing_x + boom_breedte/2, end_y, end_z + boom_hoogte/2),
                    boom_breedte, boom_hoogte
                )
                if boom2_l:
                    solids.append(boom2_l)
            
            if trapboom_rechts:
                boom2_r = create_trapboom_solid(
                    DB.XYZ(landing_x + totale_breedte - boom_breedte/2, start_point.Y, landing_z + boom_hoogte/2),
                    DB.XYZ(landing_x + totale_breedte - boom_breedte/2, end_y, end_z + boom_hoogte/2),
                    boom_breedte, boom_hoogte
                )
                if boom2_r:
                    solids.append(boom2_r)
        else:
            # Linksom: Run 2 gaat naar +Y, start bij Y = start_point.Y + totale_breedte
            if trapboom_links:
                boom2_l = create_trapboom_solid(
                    DB.XYZ(landing_x + boom_breedte/2, start_point.Y + totale_breedte, landing_z + boom_hoogte/2),
                    DB.XYZ(landing_x + boom_breedte/2, end_y, end_z + boom_hoogte/2),
                    boom_breedte, boom_hoogte
                )
                if boom2_l:
                    solids.append(boom2_l)
            
            if trapboom_rechts:
                boom2_r = create_trapboom_solid(
                    DB.XYZ(landing_x + totale_breedte - boom_breedte/2, start_point.Y + totale_breedte, landing_z + boom_hoogte/2),
                    DB.XYZ(landing_x + totale_breedte - boom_breedte/2, end_y, end_z + boom_hoogte/2),
                    boom_breedte, boom_hoogte
                )
                if boom2_r:
                    solids.append(boom2_r)
        
        # Leuningen Run 1
        # BELANGRIJK: "Links" en "Rechts" zijn vanuit het perspectief van iemand die de trap OPLOOPT
        # Bij Run 1 (loopt in +X richting): links = hoge Y, rechts = lage Y
        if leuning_links:
            # Links vanuit loopperspectief = hoge Y kant
            leuning_y = start_point.Y + totale_breedte - boom_breedte/2 if trapboom_rechts else start_point.Y + totale_breedte
            leuning1_l = create_leuning_solid(
                DB.XYZ(start_point.X, leuning_y, start_point.Z + leuning_hoogte),
                DB.XYZ(landing_x, leuning_y, landing_z + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning1_l:
                solids.append(leuning1_l)
            
            if balusters_links:
                balusters = create_balusters_along_run(
                    start_point.X, start_point.Z,
                    landing_x, landing_z,
                    leuning_y, params
                )
                solids.extend(balusters)
        
        if leuning_rechts:
            # Rechts vanuit loopperspectief = lage Y kant
            leuning_y_r1 = start_point.Y + boom_breedte/2 if trapboom_links else start_point.Y
            leuning1_r = create_leuning_solid(
                DB.XYZ(start_point.X, leuning_y_r1, start_point.Z + leuning_hoogte),
                DB.XYZ(landing_x, leuning_y_r1, landing_z + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning1_r:
                solids.append(leuning1_r)
            
            if balusters_rechts:
                balusters = create_balusters_along_run(
                    start_point.X, start_point.Z,
                    landing_x, landing_z,
                    leuning_y_r1, params
                )
                solids.extend(balusters)
        
        # Leuningen over bordes (aan de 2 buitenranden waar geen trap aangrenst)
        if is_rechtsom:
            # Rechtsom: bordes van Y=start_point.Y tot Y=start_point.Y+totale_breedte
            # Buitenranden: +X kant en +Y kant
            if leuning_rechts:
                r1_leuning_y = start_point.Y + totale_breedte - boom_breedte/2 if trapboom_rechts else start_point.Y + totale_breedte
                # Run 2 leuning rechts X-positie
                r2_leuning_x = landing_x + totale_breedte - boom_breedte/2 if trapboom_rechts else landing_x + totale_breedte
                
                # Buitenrand langs +X - moet op zelfde lijn als Run 2 rechter leuning
                leuning_bordes_1 = create_leuning_solid(
                    DB.XYZ(r2_leuning_x, start_point.Y, landing_z + leuning_hoogte),
                    DB.XYZ(r2_leuning_x, r1_leuning_y, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_1:
                    solids.append(leuning_bordes_1)
                
                # Buitenrand langs +Y (van Run 1 naar buitenrand)
                leuning_bordes_2 = create_leuning_solid(
                    DB.XYZ(landing_x, r1_leuning_y, landing_z + leuning_hoogte),
                    DB.XYZ(r2_leuning_x, r1_leuning_y, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_2:
                    solids.append(leuning_bordes_2)
            
            if leuning_links:
                l2_leuning_x = landing_x + boom_breedte/2 if trapboom_links else landing_x
                l1_leuning_y = start_point.Y + boom_breedte/2 if trapboom_links else start_point.Y
                
                # Verbinding naar Run 2 links (langs -Y kant)
                leuning_bordes_3 = create_leuning_solid(
                    DB.XYZ(l2_leuning_x, start_point.Y, landing_z + leuning_hoogte),
                    DB.XYZ(l2_leuning_x, l1_leuning_y, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_3:
                    solids.append(leuning_bordes_3)
        else:
            # Linksom: bordes van Y=start_point.Y tot Y=start_point.Y+totale_breedte
            # Buitenranden: +X kant en -Y kant (gespiegeld van rechtsom)
            if leuning_links:
                l1_leuning_y = start_point.Y + boom_breedte/2 if trapboom_links else start_point.Y
                # Run 2 leuning links X-positie
                l2_leuning_x = landing_x + boom_breedte/2 if trapboom_links else landing_x
                
                # Buitenrand langs -Y (van Run 1 links naar buitenrand)
                leuning_bordes_2 = create_leuning_solid(
                    DB.XYZ(landing_x, l1_leuning_y, landing_z + leuning_hoogte),
                    DB.XYZ(l2_leuning_x, l1_leuning_y, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_2:
                    solids.append(leuning_bordes_2)
            
            if leuning_rechts:
                r1_leuning_y = start_point.Y + totale_breedte - boom_breedte/2 if trapboom_rechts else start_point.Y + totale_breedte
                # Run 2 leuning rechts X-positie - dit is de BUITENKANT bij linksom
                r2_leuning_x = landing_x + totale_breedte - boom_breedte/2 if trapboom_rechts else landing_x + totale_breedte
                l1_leuning_y = start_point.Y + boom_breedte/2 if trapboom_links else start_point.Y
                
                # Buitenrand langs +X - van Run 1 links (l1_leuning_y) naar Run 2 start (+totale_breedte)
                leuning_bordes_1 = create_leuning_solid(
                    DB.XYZ(r2_leuning_x, l1_leuning_y, landing_z + leuning_hoogte),
                    DB.XYZ(r2_leuning_x, start_point.Y + totale_breedte, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_1:
                    solids.append(leuning_bordes_1)
                
                # Buitenrand langs -Y (van Run 1 links naar buitenrand +X)
                leuning_bordes_4 = create_leuning_solid(
                    DB.XYZ(landing_x, l1_leuning_y, landing_z + leuning_hoogte),
                    DB.XYZ(r2_leuning_x, l1_leuning_y, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_4:
                    solids.append(leuning_bordes_4)
                
                # Verbinding naar Run 2 rechts (langs +Y kant bordes)
                leuning_bordes_3 = create_leuning_solid(
                    DB.XYZ(r2_leuning_x, r1_leuning_y, landing_z + leuning_hoogte),
                    DB.XYZ(r2_leuning_x, start_point.Y + totale_breedte, landing_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_bordes_3:
                    solids.append(leuning_bordes_3)
        
        # Leuningen Run 2
        if is_rechtsom:
            # Rechtsom: Run 2 gaat naar -Y, start bij Y = start_point.Y
            if leuning_links:
                leuning_x = landing_x + boom_breedte/2 if trapboom_links else landing_x
                leuning2_l = create_leuning_solid(
                    DB.XYZ(leuning_x, start_point.Y, landing_z + leuning_hoogte),
                    DB.XYZ(leuning_x, end_y, end_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning2_l:
                    solids.append(leuning2_l)
            
            if leuning_rechts:
                leuning_x = landing_x + totale_breedte - boom_breedte/2 if trapboom_rechts else landing_x + totale_breedte
                leuning2_r = create_leuning_solid(
                    DB.XYZ(leuning_x, start_point.Y, landing_z + leuning_hoogte),
                    DB.XYZ(leuning_x, end_y, end_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning2_r:
                    solids.append(leuning2_r)
        else:
            # Linksom: Run 2 gaat naar +Y, start bij Y = start_point.Y + totale_breedte
            if leuning_links:
                leuning_x = landing_x + boom_breedte/2 if trapboom_links else landing_x
                leuning2_l = create_leuning_solid(
                    DB.XYZ(leuning_x, start_point.Y + totale_breedte, landing_z + leuning_hoogte),
                    DB.XYZ(leuning_x, end_y, end_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning2_l:
                    solids.append(leuning2_l)
            
            if leuning_rechts:
                leuning_x = landing_x + totale_breedte - boom_breedte/2 if trapboom_rechts else landing_x + totale_breedte
                leuning2_r = create_leuning_solid(
                    DB.XYZ(leuning_x, start_point.Y + totale_breedte, landing_z + leuning_hoogte),
                    DB.XYZ(leuning_x, end_y, end_z + leuning_hoogte),
                    leuning_afmeting, leuning_vorm
                )
                if leuning2_r:
                    solids.append(leuning2_r)
        
        # Balusters voor bordes en Run 2
        if is_rechtsom:
            # Rechtsom: buitenkant is +X en +Y
            r2_leuning_x = landing_x + totale_breedte - boom_breedte/2 if trapboom_rechts else landing_x + totale_breedte
            r1_leuning_y = start_point.Y + totale_breedte - boom_breedte/2 if trapboom_rechts else start_point.Y + totale_breedte
            l2_leuning_x = landing_x + boom_breedte/2 if trapboom_links else landing_x
            
            # Balusters bordes rechts - langs +X kant
            if balusters_rechts:
                balusters = create_balusters_along_bordes(
                    r2_leuning_x, start_point.Y,
                    r2_leuning_x, r1_leuning_y,
                    landing_z, params
                )
                solids.extend(balusters)
                
                # Balusters bordes rechts - langs +Y kant
                balusters = create_balusters_along_bordes(
                    landing_x, r1_leuning_y,
                    r2_leuning_x, r1_leuning_y,
                    landing_z, params
                )
                solids.extend(balusters)
            
            # Balusters Run 2 links
            if balusters_links:
                balusters = create_balusters_along_run2(
                    start_point.Y, landing_z,
                    end_y, end_z,
                    l2_leuning_x, params
                )
                solids.extend(balusters)
            
            # Balusters Run 2 rechts
            if balusters_rechts:
                balusters = create_balusters_along_run2(
                    start_point.Y, landing_z,
                    end_y, end_z,
                    r2_leuning_x, params
                )
                solids.extend(balusters)
        else:
            # Linksom: buitenkant is +X en -Y
            r2_leuning_x = landing_x + totale_breedte - boom_breedte/2 if trapboom_rechts else landing_x + totale_breedte
            l1_leuning_y = start_point.Y + boom_breedte/2 if trapboom_links else start_point.Y
            l2_leuning_x = landing_x + boom_breedte/2 if trapboom_links else landing_x
            
            # Balusters bordes rechts - langs +X kant
            if balusters_rechts:
                balusters = create_balusters_along_bordes(
                    r2_leuning_x, l1_leuning_y,
                    r2_leuning_x, start_point.Y + totale_breedte,
                    landing_z, params
                )
                solids.extend(balusters)
                
                # Balusters bordes rechts - langs -Y kant
                balusters = create_balusters_along_bordes(
                    landing_x, l1_leuning_y,
                    r2_leuning_x, l1_leuning_y,
                    landing_z, params
                )
                solids.extend(balusters)
            
            # Balusters Run 2 links
            if balusters_links:
                balusters = create_balusters_along_run2(
                    start_point.Y + totale_breedte, landing_z,
                    end_y, end_z,
                    l2_leuning_x, params
                )
                solids.extend(balusters)
            
            # Balusters Run 2 rechts
            if balusters_rechts:
                balusters = create_balusters_along_run2(
                    start_point.Y + totale_breedte, landing_z,
                    end_y, end_z,
                    r2_leuning_x, params
                )
                solids.extend(balusters)
    
    # ============ U-TRAP ============
    elif trap_type == "U-trap":
        bordes_na_trede = params.get('bordes_na_trede', aantal // 2)
        # Begrens de waarde
        bordes_na_trede = max(1, min(bordes_na_trede, aantal - 1))
        
        draairichting = params.get('draairichting', 'Rechtsom')
        is_rechtsom = draairichting == "Rechtsom"
        
        treden_run1 = bordes_na_trede
        treden_run2 = aantal - treden_run1
        gap = mm_to_feet(200)
        
        # Run 1 - treden
        for i in range(treden_run1):
            trede_x = start_point.X + (i * aantrede)
            trede_z = start_point.Z + ((i + 1) * optrede)
            
            trede = create_box_solid(
                DB.XYZ(trede_x, start_point.Y + trede_y_offset, trede_z - trede_dikte),
                aantrede, trede_breedte, trede_dikte
            )
            if trede:
                solids.append(trede)
            
            if not is_open:
                stootbord = create_box_solid(
                    DB.XYZ(trede_x, start_point.Y + trede_y_offset, start_point.Z + (i * optrede)),
                    stootbord_dikte, trede_breedte, optrede
                )
                if stootbord:
                    solids.append(stootbord)
        
        # Landing
        landing_x = start_point.X + (treden_run1 * aantrede)
        landing_z = start_point.Z + (treden_run1 * optrede)
        
        landing = create_box_solid(
            DB.XYZ(landing_x - aantrede, start_point.Y, landing_z - trede_dikte),
            aantrede * 2, totale_breedte * 2 + gap, trede_dikte
        )
        if landing:
            solids.append(landing)
        
        # Run 2 - treden (terug)
        for i in range(treden_run2):
            trede_x = landing_x - (i * aantrede)
            trede_z = landing_z + ((i + 1) * optrede)
            
            trede = create_box_solid(
                DB.XYZ(trede_x - aantrede, start_point.Y + totale_breedte + gap + trede_y_offset, trede_z - trede_dikte),
                aantrede, trede_breedte, trede_dikte
            )
            if trede:
                solids.append(trede)
            
            if not is_open:
                stootbord = create_box_solid(
                    DB.XYZ(trede_x, start_point.Y + totale_breedte + gap + trede_y_offset, landing_z + (i * optrede)),
                    stootbord_dikte, trede_breedte, optrede
                )
                if stootbord:
                    solids.append(stootbord)
        
        end_z = landing_z + treden_run2 * optrede
        end_x = landing_x - treden_run2 * aantrede
        
        # Trapbomen Run 1
        if trapboom_links:
            boom1_l = create_trapboom_solid(
                DB.XYZ(start_point.X, start_point.Y + boom_breedte/2, start_point.Z + boom_hoogte/2),
                DB.XYZ(landing_x, start_point.Y + boom_breedte/2, landing_z + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom1_l:
                solids.append(boom1_l)
        
        if trapboom_rechts:
            boom1_r = create_trapboom_solid(
                DB.XYZ(start_point.X, start_point.Y + totale_breedte - boom_breedte/2, start_point.Z + boom_hoogte/2),
                DB.XYZ(landing_x, start_point.Y + totale_breedte - boom_breedte/2, landing_z + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom1_r:
                solids.append(boom1_r)
        
        # Trapbomen Run 2
        if trapboom_links:
            boom2_l = create_trapboom_solid(
                DB.XYZ(landing_x, start_point.Y + totale_breedte + gap + boom_breedte/2, landing_z + boom_hoogte/2),
                DB.XYZ(end_x, start_point.Y + totale_breedte + gap + boom_breedte/2, end_z + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom2_l:
                solids.append(boom2_l)
        
        if trapboom_rechts:
            boom2_r = create_trapboom_solid(
                DB.XYZ(landing_x, start_point.Y + totale_breedte * 2 + gap - boom_breedte/2, landing_z + boom_hoogte/2),
                DB.XYZ(end_x, start_point.Y + totale_breedte * 2 + gap - boom_breedte/2, end_z + boom_hoogte/2),
                boom_breedte, boom_hoogte
            )
            if boom2_r:
                solids.append(boom2_r)
        
        # Leuningen Run 1
        if leuning_links:
            leuning_y = start_point.Y + boom_breedte/2 if trapboom_links else start_point.Y
            leuning1_l = create_leuning_solid(
                DB.XYZ(start_point.X, leuning_y, start_point.Z + leuning_hoogte),
                DB.XYZ(landing_x, leuning_y, landing_z + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning1_l:
                solids.append(leuning1_l)
            
            if balusters_links:
                balusters = create_balusters_along_run(
                    start_point.X, start_point.Z,
                    landing_x, landing_z,
                    leuning_y, params
                )
                solids.extend(balusters)
        
        if leuning_rechts:
            leuning_y = start_point.Y + totale_breedte - boom_breedte/2 if trapboom_rechts else start_point.Y + totale_breedte
            leuning1_r = create_leuning_solid(
                DB.XYZ(start_point.X, leuning_y, start_point.Z + leuning_hoogte),
                DB.XYZ(landing_x, leuning_y, landing_z + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning1_r:
                solids.append(leuning1_r)
            
            if balusters_rechts:
                balusters = create_balusters_along_run(
                    start_point.X, start_point.Z,
                    landing_x, landing_z,
                    leuning_y, params
                )
                solids.extend(balusters)
        
        # Leuningen Run 2
        if leuning_links:
            leuning_y = start_point.Y + totale_breedte + gap + boom_breedte/2 if trapboom_links else start_point.Y + totale_breedte + gap
            leuning2_l = create_leuning_solid(
                DB.XYZ(landing_x, leuning_y, landing_z + leuning_hoogte),
                DB.XYZ(end_x, leuning_y, end_z + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning2_l:
                solids.append(leuning2_l)
            
            if balusters_links:
                balusters = create_balusters_along_run(
                    landing_x, landing_z,
                    end_x, end_z,
                    leuning_y, params
                )
                solids.extend(balusters)
        
        if leuning_rechts:
            leuning_y = start_point.Y + totale_breedte * 2 + gap - boom_breedte/2 if trapboom_rechts else start_point.Y + totale_breedte * 2 + gap
            leuning2_r = create_leuning_solid(
                DB.XYZ(landing_x, leuning_y, landing_z + leuning_hoogte),
                DB.XYZ(end_x, leuning_y, end_z + leuning_hoogte),
                leuning_afmeting, leuning_vorm
            )
            if leuning2_r:
                solids.append(leuning2_r)
            
            if balusters_rechts:
                balusters = create_balusters_along_run(
                    landing_x, landing_z,
                    end_x, end_z,
                    leuning_y, params
                )
                solids.extend(balusters)
    
    # ============ SPILTRAP ============
    elif trap_type == "Spiltrap rechtsom" or trap_type == "Spiltrap linksom":
        # Draairichting: rechtsom = positief (tegen klok in van boven), linksom = negatief
        richting = 1 if trap_type == "Spiltrap rechtsom" else -1
        
        # Spiltrap parameters
        # Breedte is de breedte van de treden (van spil naar buiten)
        trede_breedte_spil = totale_breedte
        spil_radius = mm_to_feet(100)  # Centrale spil diameter
        binnen_radius = spil_radius  # Treden sluiten direct aan op spil
        buiten_radius = binnen_radius + trede_breedte_spil
        
        # Als er een trapboom aan de buitenzijde is, pas de trede breedte aan
        if trapboom_rechts:  # Buitenzijde
            buiten_radius_trede = buiten_radius - boom_breedte
        else:
            buiten_radius_trede = buiten_radius
        
        # Hoek per trede (360 graden / aantal treden voor volledige draai)
        # Standaard: 1 volledige rotatie over de hoogte
        totale_rotatie = 2 * math.pi * richting  # 360 graden met richting
        hoek_per_trede = totale_rotatie / aantal
        
        # Centrale spil
        spil = create_baluster(
            DB.XYZ(start_point.X, start_point.Y, start_point.Z),
            spil_radius * 2,
            aantal * optrede + trede_dikte,
            "Rond"
        )
        if spil:
            solids.append(spil)
        
        # Treden (taartpuntvormig)
        for i in range(aantal):
            hoek_start = i * hoek_per_trede
            hoek_eind = (i + 1) * hoek_per_trede
            trede_z = start_point.Z + ((i + 1) * optrede)
            
            # Maak taartpunt trede
            trede = create_pie_trede(
                start_point,
                binnen_radius,
                buiten_radius_trede,
                hoek_start,
                hoek_eind,
                trede_z - trede_dikte,
                trede_dikte
            )
            if trede:
                solids.append(trede)
        
        # Trapboom aan buitenzijde (spiraalvormig)
        if trapboom_rechts:
            for i in range(aantal):
                hoek_start = i * hoek_per_trede
                hoek_eind = (i + 1) * hoek_per_trede
                z_start = start_point.Z + (i * optrede) + boom_hoogte/2
                z_eind = start_point.Z + ((i + 1) * optrede) + boom_hoogte/2
                
                # Trapboom op buitenste radius
                boom_radius = buiten_radius - boom_breedte/2
                x1 = start_point.X + boom_radius * math.cos(hoek_start)
                y1 = start_point.Y + boom_radius * math.sin(hoek_start)
                x2 = start_point.X + boom_radius * math.cos(hoek_eind)
                y2 = start_point.Y + boom_radius * math.sin(hoek_eind)
                
                boom_seg = create_trapboom_solid(
                    DB.XYZ(x1, y1, z_start),
                    DB.XYZ(x2, y2, z_eind),
                    boom_breedte, boom_hoogte
                )
                if boom_seg:
                    solids.append(boom_seg)
        
        # Buitenste leuning (spiraalvormig)
        if leuning_rechts:  # Buitenkant
            for i in range(aantal):
                hoek_start = i * hoek_per_trede
                hoek_eind = (i + 1) * hoek_per_trede
                z_start = start_point.Z + (i * optrede) + leuning_hoogte
                z_eind = start_point.Z + ((i + 1) * optrede) + leuning_hoogte
                
                # Start en eindpunt van dit segment
                x1 = start_point.X + buiten_radius * math.cos(hoek_start)
                y1 = start_point.Y + buiten_radius * math.sin(hoek_start)
                x2 = start_point.X + buiten_radius * math.cos(hoek_eind)
                y2 = start_point.Y + buiten_radius * math.sin(hoek_eind)
                
                leuning_seg = create_leuning_solid(
                    DB.XYZ(x1, y1, z_start),
                    DB.XYZ(x2, y2, z_eind),
                    leuning_afmeting, leuning_vorm
                )
                if leuning_seg:
                    solids.append(leuning_seg)
                
                # Baluster aan buitenkant
                if balusters_rechts:
                    bal = create_baluster(
                        DB.XYZ(x1, y1, start_point.Z + (i * optrede)),
                        mm_to_feet(params.get('baluster_afmeting', 25)),
                        leuning_hoogte,
                        params.get('baluster_vorm', 'Rond')
                    )
                    if bal:
                        solids.append(bal)
    
    return solids


def create_directshape_stair(params, start_point):
    """Maak de trap als DirectShape"""
    
    solids = create_trap_geometry(params, start_point)
    
    if not solids:
        forms.alert("Kon geen geometrie maken", title="Fout")
        return None
    
    with revit.Transaction("Maak Trap (DirectShape)"):
        cat_id = DB.ElementId(DB.BuiltInCategory.OST_Stairs)
        ds = DB.DirectShape.CreateElement(doc, cat_id)
        
        geom_list = List[DB.GeometryObject]()
        for solid in solids:
            if solid:
                geom_list.Add(solid)
        
        ds.SetShape(geom_list)
        
        style_name = "Open" if params['is_open'] else "Dicht"
        ds.Name = "Trap_{0}_{1}".format(params['trap_type'].replace(" ", ""), style_name)
        
        return ds.Id
    
    return None


def pick_point():
    try:
        return uidoc.Selection.PickPoint("Selecteer startpunt voor de trap")
    except:
        return None


# Start
if __name__ == '__main__':
    window = TrapGeneratorWindow()
    window.ShowDialog()
    
    if window.result:
        forms.alert("Klik op een punt in de view om de trap te plaatsen.", title="Plaats Trap")
        start_point = pick_point()
        
        if start_point:
            result = create_directshape_stair(window.result, start_point)
            if result:
                forms.alert("Trap succesvol aangemaakt!", title="Gereed")
        else:
            forms.alert("Geen punt geselecteerd.", title="Geannuleerd")
