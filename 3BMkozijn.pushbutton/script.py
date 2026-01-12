# -*- coding: utf-8 -*-
"""
3BM-Style Kozijn Family Creator voor Revit
============================================
Dit script creëert parametrische kozijn families in Revit zonder wall hosting.
Gebaseerd op 3BM kozijn specificaties en Nederlandse bouwpraktijk.

Gebruik: Run dit script in pyRevit of Revit Python Shell
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
import System.Windows.Forms as WinForms
from System.Drawing import Point, Size, Font, FontStyle
from System.Collections.Generic import List
import System.IO
import sys

# Document en UI Context
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application


class VakConfig:
    """Configuratie voor een enkel vak in het kozijn"""
    def __init__(self):
        self.breedte = 800  # mm
        self.hoogte = 1200  # mm
        self.vulling = "Glas"  # Glas, Paneel, Deur, Raam
        self.draairichting = "Geen"  # Geen, Links, Rechts, Kantelbaar
        self.is_vast = True


class KozijnParameters:
    """Alle parameters voor het kozijn"""
    def __init__(self):
        # Indeling
        self.aantal_vakken_h = 2
        self.aantal_vakken_v = 1
        
        # Kozijnhout afmetingen (mm)
        self.randstijl_breedte = 80
        self.randstijl_dikte = 90
        self.tussenstijl_breedte = 65
        self.middenregel_hoogte = 65
        
        # Totale afmetingen (mm)
        self.totale_breedte = 1600
        self.totale_hoogte = 1200
        
        # Positie
        self.offset_vloer = 100  # mm boven vloerpeil
        
        # Sponning
        self.sponning_type = "Binnen"  # Binnen, Buiten
        self.randsponning_diepte = 15  # mm
        self.randsponning_breedte = 12  # mm
        
        # Spouwlat en folie
        self.heeft_spouwlat = True
        self.spouwlat_breedte = 18
        self.spouwlat_dikte = 50
        self.heeft_folie = True
        self.folie_dikte = 0.15  # mm (PE-folie 120 Mu)
        
        # Vakken configuratie
        self.vakken = []
        self._initialize_vakken()
    
    def _initialize_vakken(self):
        """Initialiseer vakken configuratie"""
        self.vakken = []
        for row in range(self.aantal_vakken_v):
            row_vakken = []
            for col in range(self.aantal_vakken_h):
                row_vakken.append(VakConfig())
            self.vakken.append(row_vakken)
    
    def bereken_vak_afmetingen(self):
        """Bereken automatisch vak afmetingen"""
        # Totale breedte - randstijlen - tussenstijlen
        netto_breedte = self.totale_breedte - (2 * self.randstijl_breedte)
        if self.aantal_vakken_h > 1:
            netto_breedte -= (self.aantal_vakken_h - 1) * self.tussenstijl_breedte
        vak_breedte = netto_breedte / self.aantal_vakken_h
        
        # Totale hoogte - bovenstijl - onderdorpel - middenregels
        netto_hoogte = self.totale_hoogte - (2 * self.randstijl_breedte)
        if self.aantal_vakken_v > 1:
            netto_hoogte -= (self.aantal_vakken_v - 1) * self.middenregel_hoogte
        vak_hoogte = netto_hoogte / self.aantal_vakken_v
        
        # Update alle vakken
        for row in self.vakken:
            for vak in row:
                vak.breedte = vak_breedte
                vak.hoogte = vak_hoogte


class KozijnFamilyCreator:
    """Hoofd class voor het maken van kozijn families"""
    
    def __init__(self, params):
        self.params = params
        self.family_doc = None
        
    def create_family(self, family_name="3BM_Kozijn"):
        """Creëer een nieuwe family"""
        
        # Start een nieuwe family - zoek Window template voor kozijnen
        family_template_path = None
        
        # Haal template locatie op
        template_folder = app.FamilyTemplatePath
        if not template_folder:
            template_folder = r"C:\ProgramData\Autodesk\RVT " + app.VersionNumber + r"\Family Templates\Dutch"
        
        # Zoek Window template (voor kozijnen) of Generic Model als fallback
        templates = [
            # Nederlandse templates
            "Metrisch raam.rft",
            "Metrisch kozijn.rft",
            "Metrische algemene opbouw.rft",
            # Engelse templates
            "Metric Window.rft",
            "Window.rft",
            "Metric Generic Model.rft",
            "Generic Model.rft"
        ]
        
        # Zoek in verschillende mogelijke locaties
        possible_folders = [
            template_folder,
            app.FamilyTemplatePath if app.FamilyTemplatePath else "",
            r"C:\ProgramData\Autodesk\RVT " + app.VersionNumber + r"\Family Templates\Dutch",
            r"C:\ProgramData\Autodesk\RVT " + app.VersionNumber + r"\Family Templates\English",
            r"C:\ProgramData\Autodesk\RVT " + app.VersionNumber + r"\Family Templates"
        ]
        
        for folder in possible_folders:
            if not folder:
                continue
            for template_name in templates:
                try:
                    template_path = folder + "\\" + template_name
                    if System.IO.File.Exists(template_path):
                        family_template_path = template_path
                        break
                except:
                    continue
            if family_template_path:
                break
        
        if not family_template_path:
            # Laat gebruiker zelf template selecteren
            open_dialog = WinForms.OpenFileDialog()
            open_dialog.Title = "Selecteer Family Template (Window/Kozijn)"
            open_dialog.Filter = "Revit Family Template (*.rft)|*.rft"
            open_dialog.InitialDirectory = r"C:\ProgramData\Autodesk"
            
            if open_dialog.ShowDialog() == WinForms.DialogResult.OK:
                family_template_path = open_dialog.FileName
            else:
                TaskDialog.Show("Error", "Kon geen Family template vinden.\n\nZoek locaties geprobeerd:\n" + "\n".join(possible_folders))
                return None
        
        # Maak nieuwe family document
        self.family_doc = app.NewFamilyDocument(family_template_path)
        
        if not self.family_doc:
            TaskDialog.Show("Error", "Kon family document niet aanmaken")
            return None
        
        # Start transactie
        trans = Transaction(self.family_doc, "Create Kozijn Family")
        trans.Start()
        
        try:
            # Maak family parameters
            self._create_family_parameters()
            
            # Maak geometrie
            self._create_kozijn_geometry()
            
            trans.Commit()
            
            # Sla family op
            save_options = SaveAsOptions()
            save_options.OverwriteExistingFile = True
            
            # Bepaal opslaglocatie: 70_BIM submap van projectlocatie
            project_path = doc.PathName
            save_path = None
            
            if project_path:
                # Project is opgeslagen, zoek 70_BIM map
                project_folder = System.IO.Path.GetDirectoryName(project_path)
                
                # Zoek naar 70_BIM map (kan in huidige map of parent mappen zijn)
                search_folder = project_folder
                bim_folder = None
                
                # Zoek maximaal 3 niveaus omhoog
                for i in range(4):
                    test_path = System.IO.Path.Combine(search_folder, "70_BIM")
                    if System.IO.Directory.Exists(test_path):
                        bim_folder = test_path
                        break
                    # Ga een niveau omhoog
                    parent = System.IO.Directory.GetParent(search_folder)
                    if parent:
                        search_folder = parent.FullName
                    else:
                        break
                
                # Als 70_BIM niet gevonden, maak aan in projectmap
                if not bim_folder:
                    bim_folder = System.IO.Path.Combine(project_folder, "70_BIM")
                    if not System.IO.Directory.Exists(bim_folder):
                        System.IO.Directory.CreateDirectory(bim_folder)
                
                save_path = System.IO.Path.Combine(bim_folder, family_name + ".rfa")
            
            # Als geen projectpad of gebruiker wil andere locatie
            if not save_path or not project_path:
                # Fallback: vraag gebruiker
                save_dialog = WinForms.SaveFileDialog()
                save_dialog.Filter = "Revit Family Files (*.rfa)|*.rfa"
                save_dialog.FileName = family_name + ".rfa"
                save_dialog.Title = "Kozijn Family Opslaan"
                
                if save_dialog.ShowDialog() == WinForms.DialogResult.OK:
                    save_path = save_dialog.FileName
                else:
                    trans.RollBack()
                    return None
            
            # Sla op
            self.family_doc.SaveAs(save_path, save_options)
            TaskDialog.Show("Succes", 
                "Kozijn family opgeslagen in:\n" + save_path)
            
            # Optioneel: load in current project
            result = TaskDialog.Show("Family Laden", 
                "Wilt u deze family laden in het huidige project?",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No)
            
            if result == TaskDialogResult.Yes:
                self._load_into_project(save_path)
            
            return save_path
            
        except Exception as e:
            trans.RollBack()
            TaskDialog.Show("Error", "Fout bij aanmaken family:\n" + str(e))
            return None
        
        return None
    
    def _create_family_parameters(self):
        """Maak alle family parameters aan"""
        fm = self.family_doc.FamilyManager
        
        # Detecteer Revit versie voor juiste API
        revit_version = int(app.VersionNumber)
        
        if revit_version >= 2022:
            # Nieuwe API met SpecTypeId (Revit 2022+)
            self._create_parameters_new_api(fm)
        else:
            # Oude API met ParameterType (Revit 2021 en eerder)
            self._create_parameters_old_api(fm)
    
    def _create_parameters_new_api(self, fm):
        """Parameters aanmaken met nieuwe API (Revit 2022+)"""
        try:
            # Afmetingen parameters (Length)
            self._add_param_new(fm, "Totale_Breedte", SpecTypeId.Length, 
                               self.params.totale_breedte / 304.8)
            self._add_param_new(fm, "Totale_Hoogte", SpecTypeId.Length, 
                               self.params.totale_hoogte / 304.8)
            
            # Kozijnhout parameters
            self._add_param_new(fm, "Randstijl_Breedte", SpecTypeId.Length, 
                               self.params.randstijl_breedte / 304.8)
            self._add_param_new(fm, "Randstijl_Dikte", SpecTypeId.Length, 
                               self.params.randstijl_dikte / 304.8)
            self._add_param_new(fm, "Tussenstijl_Breedte", SpecTypeId.Length, 
                               self.params.tussenstijl_breedte / 304.8)
            
            # Positie parameters
            self._add_param_new(fm, "Offset_Vloer", SpecTypeId.Length, 
                               self.params.offset_vloer / 304.8)
            
            # Vakindeling (Integer)
            self._add_param_new(fm, "Aantal_Vakken_Horizontaal", SpecTypeId.Int.Integer, 
                               self.params.aantal_vakken_h)
            self._add_param_new(fm, "Aantal_Vakken_Verticaal", SpecTypeId.Int.Integer, 
                               self.params.aantal_vakken_v)
            
            # Sponning
            self._add_param_new(fm, "Sponning_Diepte", SpecTypeId.Length, 
                               self.params.randsponning_diepte / 304.8)
        except Exception as e:
            TaskDialog.Show("Parameter Error", str(e))
    
    def _add_param_new(self, fm, name, spec_type_id, default_value):
        """Voeg parameter toe met nieuwe API"""
        try:
            # Maak parameter met ForgeTypeId
            param = fm.AddParameter(name, GroupTypeId.Geometry, spec_type_id, False)
            if param and default_value is not None:
                if spec_type_id == SpecTypeId.Int.Integer:
                    fm.Set(param, int(default_value))
                else:
                    fm.Set(param, float(default_value))
        except Exception as e:
            # Parameter bestaat mogelijk al of andere fout
            pass
    
    def _create_parameters_old_api(self, fm):
        """Parameters aanmaken met oude API (Revit 2021 en eerder)"""
        # Afmetingen parameters
        self._add_parameter_old(fm, "Totale_Breedte", "Length", 
                           self.params.totale_breedte / 304.8)
        self._add_parameter_old(fm, "Totale_Hoogte", "Length", 
                           self.params.totale_hoogte / 304.8)
        
        # Kozijnhout parameters
        self._add_parameter_old(fm, "Randstijl_Breedte", "Length", 
                           self.params.randstijl_breedte / 304.8)
        self._add_parameter_old(fm, "Randstijl_Dikte", "Length", 
                           self.params.randstijl_dikte / 304.8)
        self._add_parameter_old(fm, "Tussenstijl_Breedte", "Length", 
                           self.params.tussenstijl_breedte / 304.8)
        
        # Positie parameters
        self._add_parameter_old(fm, "Offset_Vloer", "Length", 
                           self.params.offset_vloer / 304.8)
        
        # Vakindeling
        self._add_parameter_old(fm, "Aantal_Vakken_Horizontaal", "Integer", 
                           self.params.aantal_vakken_h)
        self._add_parameter_old(fm, "Aantal_Vakken_Verticaal", "Integer", 
                           self.params.aantal_vakken_v)
        
        # Sponning parameters
        self._add_parameter_old(fm, "Sponning_Diepte", "Length", 
                           self.params.randsponning_diepte / 304.8)
    
    def _add_parameter_old(self, fm, name, param_type_str, default_value):
        """Hulpfunctie om parameter toe te voegen (oude API)"""
        try:
            # Haal ParameterType enum waarde op
            param_type = getattr(ParameterType, param_type_str)
            param = fm.AddParameter(name, BuiltInParameterGroup.PG_GEOMETRY, param_type, False)
            if param and default_value is not None:
                if param_type_str == "Integer":
                    fm.Set(param, int(default_value))
                else:
                    fm.Set(param, float(default_value))
        except Exception as e:
            # Parameter bestaat mogelijk al
            pass
    
    def _create_kozijn_geometry(self):
        """Maak de kozijn geometrie"""
        
        # Converteer mm naar Revit internal units (feet)
        mm_to_feet = 1.0 / 304.8
        
        # Bereken vak afmetingen
        self.params.bereken_vak_afmetingen()
        
        # Basis punten
        origin = XYZ(0, 0, 0)
        
        # Maak kozijnframe (randstijlen en dorpels)
        self._create_frame()
        
        # Maak tussenstijlen en middenregels
        if self.params.aantal_vakken_h > 1:
            self._create_tussenstijlen()
        
        if self.params.aantal_vakken_v > 1:
            self._create_middenregels()
        
        # Maak glas/vulling per vak
        self._create_vak_vullingen()
        
        # Voeg spouwlat toe indien nodig
        if self.params.heeft_spouwlat:
            self._create_spouwlat()
    
    def _create_frame(self):
        """Maak het kozijnframe (randstijlen en dorpels)"""
        mm_to_feet = 1.0 / 304.8
        
        # Maak extrusion profile voor kozijnhout
        # Dit is een vereenvoudigde versie - in praktijk complexere profielen
        
        b = self.params.totale_breedte * mm_to_feet
        h = self.params.totale_hoogte * mm_to_feet
        rs_b = self.params.randstijl_breedte * mm_to_feet
        rs_d = self.params.randstijl_dikte * mm_to_feet
        
        # Linker stijl
        self._create_rectangular_extrusion(
            XYZ(0, 0, 0),
            rs_b, h, rs_d,
            "Kozijnhout"
        )
        
        # Rechter stijl
        self._create_rectangular_extrusion(
            XYZ(b - rs_b, 0, 0),
            rs_b, h, rs_d,
            "Kozijnhout"
        )
        
        # Onder dorpel
        self._create_rectangular_extrusion(
            XYZ(0, 0, 0),
            b, rs_b, rs_d,
            "Kozijnhout"
        )
        
        # Boven dorpel
        self._create_rectangular_extrusion(
            XYZ(0, h - rs_b, 0),
            b, rs_b, rs_d,
            "Kozijnhout"
        )
    
    def _create_tussenstijlen(self):
        """Maak verticale tussenstijlen"""
        mm_to_feet = 1.0 / 304.8
        
        b = self.params.totale_breedte * mm_to_feet
        h = self.params.totale_hoogte * mm_to_feet
        rs_b = self.params.randstijl_breedte * mm_to_feet
        ts_b = self.params.tussenstijl_breedte * mm_to_feet
        rs_d = self.params.randstijl_dikte * mm_to_feet
        
        # Bereken netto ruimte voor vakken
        netto_breedte = b - 2 * rs_b
        vak_breedte = (netto_breedte - (self.params.aantal_vakken_h - 1) * ts_b) / self.params.aantal_vakken_h
        
        # Plaats tussenstijlen
        for i in range(1, self.params.aantal_vakken_h):
            x_pos = rs_b + i * vak_breedte + (i - 0.5) * ts_b
            self._create_rectangular_extrusion(
                XYZ(x_pos, rs_b, 0),
                ts_b, h - 2 * rs_b, rs_d,
                "Kozijnhout"
            )
    
    def _create_middenregels(self):
        """Maak horizontale middenregels"""
        mm_to_feet = 1.0 / 304.8
        
        b = self.params.totale_breedte * mm_to_feet
        h = self.params.totale_hoogte * mm_to_feet
        rs_b = self.params.randstijl_breedte * mm_to_feet
        mr_h = self.params.middenregel_hoogte * mm_to_feet
        rs_d = self.params.randstijl_dikte * mm_to_feet
        
        # Bereken netto ruimte voor vakken
        netto_hoogte = h - 2 * rs_b
        vak_hoogte = (netto_hoogte - (self.params.aantal_vakken_v - 1) * mr_h) / self.params.aantal_vakken_v
        
        # Plaats middenregels
        for i in range(1, self.params.aantal_vakken_v):
            y_pos = rs_b + i * vak_hoogte + (i - 0.5) * mr_h
            self._create_rectangular_extrusion(
                XYZ(rs_b, y_pos, 0),
                b - 2 * rs_b, mr_h, rs_d,
                "Kozijnhout"
            )
    
    def _create_vak_vullingen(self):
        """Maak glas/paneel vulling per vak"""
        mm_to_feet = 1.0 / 304.8
        
        b = self.params.totale_breedte * mm_to_feet
        h = self.params.totale_hoogte * mm_to_feet
        rs_b = self.params.randstijl_breedte * mm_to_feet
        rs_d = self.params.randstijl_dikte * mm_to_feet
        ts_b = self.params.tussenstijl_breedte * mm_to_feet
        mr_h = self.params.middenregel_hoogte * mm_to_feet
        
        # Bereken vak afmetingen
        netto_breedte = b - 2 * rs_b
        vak_breedte = (netto_breedte - (self.params.aantal_vakken_h - 1) * ts_b) / self.params.aantal_vakken_h
        
        netto_hoogte = h - 2 * rs_b
        vak_hoogte = (netto_hoogte - (self.params.aantal_vakken_v - 1) * mr_h) / self.params.aantal_vakken_v
        
        # Glas dikte (HR+++ triple glas: 4-14-4-15-4 = 41mm)
        glas_dikte = 0.041 * mm_to_feet * 304.8  # 41mm
        sponning = self.params.randsponning_diepte * mm_to_feet
        
        # Plaats glas per vak
        for row in range(self.params.aantal_vakken_v):
            for col in range(self.params.aantal_vakken_h):
                x_pos = rs_b + col * (vak_breedte + ts_b) + sponning
                y_pos = rs_b + row * (vak_hoogte + mr_h) + sponning
                z_pos = (rs_d - glas_dikte) / 2
                
                glass_width = vak_breedte - 2 * sponning
                glass_height = vak_hoogte - 2 * sponning
                
                # Maak glas/vulling
                vak = self.params.vakken[row][col]
                if vak.vulling == "Glas":
                    self._create_rectangular_extrusion(
                        XYZ(x_pos, y_pos, z_pos),
                        glass_width, glass_height, glas_dikte,
                        "Glas"
                    )
    
    def _create_spouwlat(self):
        """Maak spouwlat rondom kozijn"""
        mm_to_feet = 1.0 / 304.8
        
        b = self.params.totale_breedte * mm_to_feet
        h = self.params.totale_hoogte * mm_to_feet
        lat_b = self.params.spouwlat_breedte * mm_to_feet
        lat_d = self.params.spouwlat_dikte * mm_to_feet
        rs_d = self.params.randstijl_dikte * mm_to_feet
        
        # Spouwlat aan achterzijde kozijn
        z_pos = -(lat_d / 2)
        
        # Links
        self._create_rectangular_extrusion(
            XYZ(-lat_b, 0, z_pos),
            lat_b, h, lat_d,
            "Hout"
        )
        
        # Rechts
        self._create_rectangular_extrusion(
            XYZ(b, 0, z_pos),
            lat_b, h, lat_d,
            "Hout"
        )
        
        # Boven
        self._create_rectangular_extrusion(
            XYZ(0, h, z_pos),
            b, lat_b, lat_d,
            "Hout"
        )
        
        # Onder
        self._create_rectangular_extrusion(
            XYZ(0, -lat_b, z_pos),
            b, lat_b, lat_d,
            "Hout"
        )
    
    def _create_rectangular_extrusion(self, origin, width, height, depth, material_name):
        """Hulpfunctie om rechthoekige extrusion te maken"""
        try:
            # Maak een rechthoekig profiel
            profile_loops = List[CurveLoop]()
            profile = CurveLoop()
            
            # Definieer rechthoek in XY vlak
            p1 = origin
            p2 = origin + XYZ(width, 0, 0)
            p3 = origin + XYZ(width, height, 0)
            p4 = origin + XYZ(0, height, 0)
            
            profile.Append(Line.CreateBound(p1, p2))
            profile.Append(Line.CreateBound(p2, p3))
            profile.Append(Line.CreateBound(p3, p4))
            profile.Append(Line.CreateBound(p4, p1))
            
            profile_loops.Add(profile)
            
            # Maak solid extrusion in Z richting
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                profile_loops,
                XYZ(0, 0, 1),  # Extrusion richting
                depth
            )
            
            # Maak DirectShape om solid toe te voegen
            ds = DirectShape.CreateElement(self.family_doc, ElementId(BuiltInCategory.OST_GenericModel))
            shape_list = List[GeometryObject]()
            shape_list.Add(solid)
            ds.SetShape(shape_list)
            
            return ds
            
        except Exception as e:
            TaskDialog.Show("Geometrie Error", str(e))
            return None
    
    def _load_into_project(self, family_path):
        """Laad family in het huidige project"""
        try:
            trans = Transaction(doc, "Load Kozijn Family")
            trans.Start()
            
            # Load family
            family = doc.LoadFamily(family_path)
            
            trans.Commit()
            
            TaskDialog.Show("Succes", "Family geladen in project!")
            
        except Exception as e:
            TaskDialog.Show("Error", "Kon family niet laden:\n" + str(e))


class KozijnConfigDialog(WinForms.Form):
    """GUI Dialog voor kozijn configuratie"""
    
    def __init__(self):
        self.params = KozijnParameters()
        self.result = None
        self._init_ui()
    
    def _init_ui(self):
        """Initialiseer de user interface"""
        self.Text = "3BM Kozijn Configurator"
        self.Size = Size(600, 800)
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        
        y_pos = 20
        
        # Title
        title = WinForms.Label()
        title.Text = "Kozijn Family Generator"
        title.Font = Font("Arial", 14, FontStyle.Bold)
        title.Location = Point(20, y_pos)
        title.Size = Size(550, 30)
        self.Controls.Add(title)
        y_pos += 40
        
        # === INDELING SECTIE ===
        section1 = WinForms.Label()
        section1.Text = "1. VAKINDELING"
        section1.Font = Font("Arial", 10, FontStyle.Bold)
        section1.Location = Point(20, y_pos)
        section1.Size = Size(550, 20)
        self.Controls.Add(section1)
        y_pos += 25
        
        # Aantal vakken horizontaal
        y_pos = self._add_numeric_input("Aantal vakken horizontaal:", 
                                        self.params.aantal_vakken_h, 1, 10, y_pos, "vakken_h")
        
        # Aantal vakken verticaal
        y_pos = self._add_numeric_input("Aantal vakken verticaal:", 
                                        self.params.aantal_vakken_v, 1, 5, y_pos, "vakken_v")
        
        y_pos += 10
        
        # === AFMETINGEN SECTIE ===
        section2 = WinForms.Label()
        section2.Text = "2. TOTALE AFMETINGEN (mm)"
        section2.Font = Font("Arial", 10, FontStyle.Bold)
        section2.Location = Point(20, y_pos)
        section2.Size = Size(550, 20)
        self.Controls.Add(section2)
        y_pos += 25
        
        # Totale breedte
        y_pos = self._add_numeric_input("Totale breedte:", 
                                        self.params.totale_breedte, 400, 5000, y_pos, "totaal_b")
        
        # Totale hoogte
        y_pos = self._add_numeric_input("Totale hoogte:", 
                                        self.params.totale_hoogte, 400, 3000, y_pos, "totaal_h")
        
        y_pos += 10
        
        # === KOZIJNHOUT SECTIE ===
        section3 = WinForms.Label()
        section3.Text = "3. KOZIJNHOUT AFMETINGEN (mm)"
        section3.Font = Font("Arial", 10, FontStyle.Bold)
        section3.Location = Point(20, y_pos)
        section3.Size = Size(550, 20)
        self.Controls.Add(section3)
        y_pos += 25
        
        # Randstijl breedte
        y_pos = self._add_numeric_input("Randstijl breedte:", 
                                        self.params.randstijl_breedte, 60, 120, y_pos, "rand_b")
        
        # Randstijl dikte
        y_pos = self._add_numeric_input("Randstijl dikte:", 
                                        self.params.randstijl_dikte, 60, 160, y_pos, "rand_d")
        
        # Tussenstijl breedte
        y_pos = self._add_numeric_input("Tussenstijl breedte:", 
                                        self.params.tussenstijl_breedte, 50, 100, y_pos, "tussen_b")
        
        y_pos += 10
        
        # === SPONNING SECTIE ===
        section4 = WinForms.Label()
        section4.Text = "4. SPONNING EN DETAILS"
        section4.Font = Font("Arial", 10, FontStyle.Bold)
        section4.Location = Point(20, y_pos)
        section4.Size = Size(550, 20)
        self.Controls.Add(section4)
        y_pos += 25
        
        # Sponning type
        label = WinForms.Label()
        label.Text = "Sponning type:"
        label.Location = Point(20, y_pos)
        label.Size = Size(200, 20)
        self.Controls.Add(label)
        
        self.sponning_combo = WinForms.ComboBox()
        self.sponning_combo.Items.Add("Binnen")
        self.sponning_combo.Items.Add("Buiten")
        self.sponning_combo.SelectedItem = "Binnen"
        self.sponning_combo.Location = Point(230, y_pos)
        self.sponning_combo.Size = Size(150, 25)
        self.sponning_combo.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.sponning_combo)
        y_pos += 30
        
        # Offset vloer
        y_pos = self._add_numeric_input("Offset vanaf vloerpeil:", 
                                        self.params.offset_vloer, 0, 500, y_pos, "offset")
        
        # Spouwlat checkbox
        self.spouwlat_check = WinForms.CheckBox()
        self.spouwlat_check.Text = "Spouwlat toevoegen"
        self.spouwlat_check.Checked = True
        self.spouwlat_check.Location = Point(20, y_pos)
        self.spouwlat_check.Size = Size(200, 25)
        self.Controls.Add(self.spouwlat_check)
        y_pos += 30
        
        # Folie checkbox
        self.folie_check = WinForms.CheckBox()
        self.folie_check.Text = "PE-folie bescherming"
        self.folie_check.Checked = True
        self.folie_check.Location = Point(20, y_pos)
        self.folie_check.Size = Size(200, 25)
        self.Controls.Add(self.folie_check)
        y_pos += 40
        
        # === MATERIAAL KEUZE ===
        section5 = WinForms.Label()
        section5.Text = "5. MATERIAAL"
        section5.Font = Font("Arial", 10, FontStyle.Bold)
        section5.Location = Point(20, y_pos)
        section5.Size = Size(550, 20)
        self.Controls.Add(section5)
        y_pos += 25
        
        label = WinForms.Label()
        label.Text = "Kozijnmateriaal:"
        label.Location = Point(20, y_pos)
        label.Size = Size(200, 20)
        self.Controls.Add(label)
        
        self.materiaal_combo = WinForms.ComboBox()
        self.materiaal_combo.Items.Add("Meranti")
        self.materiaal_combo.Items.Add("Grenen")
        self.materiaal_combo.Items.Add("Kunststof")
        self.materiaal_combo.SelectedItem = "Meranti"
        self.materiaal_combo.Location = Point(230, y_pos)
        self.materiaal_combo.Size = Size(150, 25)
        self.materiaal_combo.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.materiaal_combo)
        y_pos += 35
        
        # === BUTTONS ===
        y_pos += 20
        
        # Create button
        create_btn = WinForms.Button()
        create_btn.Text = "Family Aanmaken"
        create_btn.Location = Point(150, y_pos)
        create_btn.Size = Size(150, 35)
        create_btn.Click += self._on_create_click
        self.Controls.Add(create_btn)
        
        # Cancel button
        cancel_btn = WinForms.Button()
        cancel_btn.Text = "Annuleren"
        cancel_btn.Location = Point(320, y_pos)
        cancel_btn.Size = Size(120, 35)
        cancel_btn.Click += self._on_cancel_click
        self.Controls.Add(cancel_btn)
    
    def _add_numeric_input(self, label_text, default_value, min_val, max_val, y_pos, control_name):
        """Hulpfunctie om numerieke input toe te voegen"""
        label = WinForms.Label()
        label.Text = label_text
        label.Location = Point(20, y_pos)
        label.Size = Size(200, 20)
        self.Controls.Add(label)
        
        numeric = WinForms.NumericUpDown()
        numeric.Minimum = min_val
        numeric.Maximum = max_val
        numeric.Value = default_value
        numeric.Location = Point(230, y_pos)
        numeric.Size = Size(100, 25)
        numeric.Name = control_name
        self.Controls.Add(numeric)
        
        unit_label = WinForms.Label()
        unit_label.Text = "mm" if "breedte" in label_text.lower() or "hoogte" in label_text.lower() or "dikte" in label_text.lower() or "offset" in label_text.lower() else ""
        unit_label.Location = Point(340, y_pos)
        unit_label.Size = Size(40, 20)
        self.Controls.Add(unit_label)
        
        return y_pos + 30
    
    def _on_create_click(self, sender, event):
        """Handle create button click"""
        # Update parameters from controls
        self.params.aantal_vakken_h = int(self.Controls.Find("vakken_h", True)[0].Value)
        self.params.aantal_vakken_v = int(self.Controls.Find("vakken_v", True)[0].Value)
        self.params.totale_breedte = float(self.Controls.Find("totaal_b", True)[0].Value)
        self.params.totale_hoogte = float(self.Controls.Find("totaal_h", True)[0].Value)
        self.params.randstijl_breedte = float(self.Controls.Find("rand_b", True)[0].Value)
        self.params.randstijl_dikte = float(self.Controls.Find("rand_d", True)[0].Value)
        self.params.tussenstijl_breedte = float(self.Controls.Find("tussen_b", True)[0].Value)
        self.params.offset_vloer = float(self.Controls.Find("offset", True)[0].Value)
        self.params.sponning_type = self.sponning_combo.SelectedItem
        self.params.heeft_spouwlat = self.spouwlat_check.Checked
        self.params.heeft_folie = self.folie_check.Checked
        
        # Herinitialiseer vakken met nieuwe aantallen
        self.params._initialize_vakken()
        
        self.result = self.params
        self.DialogResult = WinForms.DialogResult.OK
        self.Close()
    
    def _on_cancel_click(self, sender, event):
        """Handle cancel button click"""
        self.result = None
        self.DialogResult = WinForms.DialogResult.Cancel
        self.Close()


def main():
    """Main functie - start de kozijn creator"""
    
    # Toon configuratie dialog
    dialog = KozijnConfigDialog()
    
    if dialog.ShowDialog() == WinForms.DialogResult.OK and dialog.result:
        params = dialog.result
        
        # Maak family creator
        creator = KozijnFamilyCreator(params)
        
        # Genereer family naam
        family_name = "3BM_Kozijn_{}x{}_{}x{}mm".format(
            params.aantal_vakken_h,
            params.aantal_vakken_v,
            int(params.totale_breedte),
            int(params.totale_hoogte)
        )
        
        # Maak family
        result = creator.create_family(family_name)
        
        if result:
            TaskDialog.Show("Voltooid", 
                "Kozijn family succesvol aangemaakt!\n\n" +
                "Parameters:\n" +
                "- Vakken: {}H x {}V\n".format(params.aantal_vakken_h, params.aantal_vakken_v) +
                "- Afmetingen: {} x {} mm\n".format(int(params.totale_breedte), int(params.totale_hoogte)) +
                "- Kozijnhout: {} x {} mm\n".format(int(params.randstijl_breedte), int(params.randstijl_dikte)) +
                "- Spouwlat: {}\n".format("Ja" if params.heeft_spouwlat else "Nee"))


# Start het script
if __name__ == "__main__":
    main()
