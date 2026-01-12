# -*- coding: utf-8 -*-
"""Family Manager - Export en Import families naar/van Dropbox bibliotheek met thumbnails"""

__title__ = "Family\nManager"
__author__ = "Your Name"

from pyrevit import forms, script, revit, DB
import os
import shutil
import json
import ctypes
from System.Collections.Generic import List
from System.Windows.Forms import FolderBrowserDialog, DialogResult

# WPF imports voor custom dialoog met thumbnails
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System.Windows import Window, SizeToContent, WindowStartupLocation, Thickness, GridLength, GridUnitType, HorizontalAlignment, VerticalAlignment
from System.Windows.Controls import (
    ListBox, ListBoxItem, StackPanel, WrapPanel, Image, TextBlock, 
    Button, ScrollViewer, CheckBox, Grid, ColumnDefinition, RowDefinition,
    Orientation, SelectionMode, Border, DockPanel, Dock, TextBox, Label,
    ScrollBarVisibility
)
from System.Windows.Media import Brushes, Stretch
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind

output = script.get_output()
doc = revit.doc

# Config file locatie
CONFIG_FILE = os.path.join(
    os.path.dirname(__file__),
    'family_manager_config.json'
)

# Thumbnail cache folder
THUMB_CACHE = os.path.join(os.path.dirname(__file__), 'thumb_cache')

# Constante voor nieuwe map optie
NEW_FOLDER_OPTION = "[+] Nieuwe map aanmaken..."


def get_library_path():
    """Haal de opgeslagen bibliotheek path op"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                return config.get('library_path', None)
        except:
            pass
    return None


def save_library_path(path):
    """Sla de bibliotheek path op"""
    config = {'library_path': path}
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)


def configure_library_path():
    """Laat gebruiker de bibliotheek locatie kiezen"""
    current_path = get_library_path()
    
    if current_path:
        message = "Huidige bibliotheek locatie:\n{}\n\nWil je een nieuwe locatie kiezen?".format(current_path)
        if not forms.alert(message, yes=True, no=True):
            return current_path
    
    dialog = FolderBrowserDialog()
    dialog.Description = "Selecteer de Dropbox Family Library folder"
    if current_path and os.path.exists(current_path):
        dialog.SelectedPath = current_path
    
    result = dialog.ShowDialog()
    
    if result == DialogResult.OK:
        new_path = dialog.SelectedPath
        save_library_path(new_path)
        forms.alert("Bibliotheek locatie opgeslagen:\n{}".format(new_path))
        return new_path
    
    return current_path


def ensure_library_exists():
    """Zorg dat de bibliotheek folder bestaat"""
    library_path = get_library_path()
    
    if not library_path:
        forms.alert(
            "Geen bibliotheek locatie geconfigureerd!\n\nKies eerst een bibliotheek locatie.",
            exitscript=True
        )
        return None
    
    if not os.path.exists(library_path):
        os.makedirs(library_path)
    
    return library_path


def get_subfolders(library_path):
    """Haal alleen directe submappen op (niet recursief, geen verborgen mappen)"""
    subfolders = []
    
    if not os.path.exists(library_path):
        return subfolders
    
    for item in os.listdir(library_path):
        item_path = os.path.join(library_path, item)
        if os.path.isdir(item_path) and not item.startswith('.'):
            subfolders.append(item)
    
    return sorted(subfolders)


def get_families_in_project():
    """Haal alle families uit het project op"""
    collector = DB.FilteredElementCollector(doc)
    families = collector.OfClass(DB.Family).ToElements()
    
    editable_families = []
    for fam in families:
        if fam.IsEditable:
            editable_families.append(fam)
    
    return editable_families


def get_families_in_library():
    """Haal alle .rfa files uit de bibliotheek"""
    library_path = ensure_library_exists()
    if not library_path:
        return []
    
    family_files = []
    
    for root, dirs, files in os.walk(library_path):
        for file in files:
            if file.lower().endswith('.rfa'):
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, library_path)
                family_files.append({
                    'name': file,
                    'path': full_path,
                    'relative': rel_path
                })
    
    return family_files


def extract_rfa_thumbnail(rfa_path, output_path):
    """Extract thumbnail uit een .rfa bestand"""
    try:
        with open(rfa_path, 'rb') as f:
            data = f.read()
        
        # Revit families slaan thumbnails op als EMF of PNG in de OLE storage
        # Zoek naar verschillende image signatures
        
        # PNG signature
        png_sig = b'\x89PNG\r\n\x1a\n'
        png_start = data.find(png_sig)
        
        if png_start != -1:
            # Zoek alle mogelijke PNG einden
            search_start = png_start
            best_png = None
            
            while True:
                iend_pos = data.find(b'IEND', search_start)
                if iend_pos == -1:
                    break
                
                # PNG IEND chunk eindigt met CRC (4 bytes na IEND)
                png_data = data[png_start:iend_pos + 8]
                
                # Neem de grootste geldige PNG (vaak zijn er meerdere)
                if len(png_data) > 500:  # Minimaal 500 bytes voor een echte thumbnail
                    if best_png is None or len(png_data) > len(best_png):
                        best_png = png_data
                
                search_start = iend_pos + 8
                
                # Stop na eerste goede PNG
                if best_png and len(best_png) > 1000:
                    break
            
            if best_png:
                with open(output_path, 'wb') as out:
                    out.write(best_png)
                return True
        
        # BMP signature (soms gebruikt)
        bmp_sig = b'BM'
        bmp_start = data.find(bmp_sig)
        if bmp_start != -1 and bmp_start < len(data) - 54:  # BMP header is 54 bytes
            # Lees BMP grootte uit header (little endian)
            try:
                import struct
                size_bytes = data[bmp_start+2:bmp_start+6]
                bmp_size = struct.unpack('<I', size_bytes)[0]
                if 1000 < bmp_size < 500000:  # Redelijke grootte
                    bmp_data = data[bmp_start:bmp_start+bmp_size]
                    bmp_output = output_path.replace('.png', '.bmp')
                    with open(bmp_output, 'wb') as out:
                        out.write(bmp_data)
                    return True
            except:
                pass
        
        return False
        
    except Exception as e:
        return False


def get_family_thumbnail_from_revit(family):
    """Haal thumbnail op van een geladen family via Revit API"""
    try:
        # Probeer een family symbol te vinden voor de preview
        family_symbols = list(family.GetFamilySymbolIds())
        if family_symbols:
            symbol_id = family_symbols[0]
            symbol = doc.GetElement(symbol_id)
            if symbol:
                # Probeer preview image te krijgen
                size = DB.ImageExportOptions.GetPixelSize(DB.ImageResolution.DPI_150)
                # Dit werkt niet altijd, maar we proberen het
                pass
    except:
        pass
    return None


def get_or_create_thumbnail(rfa_path):
    """Haal thumbnail op uit cache of extract uit RFA"""
    if not os.path.exists(THUMB_CACHE):
        os.makedirs(THUMB_CACHE)
    
    # Maak unieke cache naam
    file_name = os.path.basename(rfa_path)
    file_hash = str(hash(rfa_path) & 0xffffffff)
    cache_base = "{}_{}".format(file_hash, file_name.replace('.rfa', '').replace('.RFA', ''))
    cache_path_png = os.path.join(THUMB_CACHE, cache_base + ".png")
    cache_path_jpg = os.path.join(THUMB_CACHE, cache_base + ".jpg")
    cache_path_bmp = os.path.join(THUMB_CACHE, cache_base + ".bmp")
    
    # Check cache (PNG, JPG of BMP)
    for cache_path in [cache_path_png, cache_path_jpg, cache_path_bmp]:
        if os.path.exists(cache_path):
            # Check of cache nieuwer is dan bestand
            try:
                if os.path.getmtime(cache_path) >= os.path.getmtime(rfa_path):
                    return cache_path
                else:
                    # Verwijder oude cache
                    os.remove(cache_path)
            except:
                pass
    
    # Extract thumbnail
    if extract_rfa_thumbnail(rfa_path, cache_path_png):
        # Check welk bestand is aangemaakt
        for cache_path in [cache_path_png, cache_path_jpg, cache_path_bmp]:
            if os.path.exists(cache_path):
                return cache_path
    
    return None


class ThumbnailSelectDialog(Window):
    """Custom WPF dialoog met thumbnail weergave"""
    
    def __init__(self, items, title="Selecteer items", multiselect=True, show_thumbnails=True):
        self.Title = title
        self.Width = 850
        self.Height = 650
        self.MinWidth = 600
        self.MinHeight = 400
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        
        self.items = items  # List of dicts with 'name', 'path' (optional), 'thumbnail' (optional)
        self.selected_items = []
        self.multiselect = multiselect
        self.show_thumbnails = show_thumbnails
        
        self._build_ui()
    
    def _build_ui(self):
        # Main grid
        main_grid = Grid()
        row1 = RowDefinition()
        row1.Height = GridLength(1, GridUnitType.Star)  # Neemt alle beschikbare ruimte
        main_grid.RowDefinitions.Add(row1)
        row2 = RowDefinition()
        row2.Height = GridLength(60, GridUnitType.Pixel)  # Vaste hoogte voor knoppen
        main_grid.RowDefinitions.Add(row2)
        
        # Scrollable wrap panel voor items
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll.Margin = Thickness(5)
        
        self.wrap_panel = WrapPanel()
        self.wrap_panel.Orientation = Orientation.Horizontal
        self.wrap_panel.Margin = Thickness(5)
        
        self.checkboxes = []
        
        for item in self.items:
            item_panel = self._create_item_panel(item)
            self.wrap_panel.Children.Add(item_panel)
        
        scroll.Content = self.wrap_panel
        Grid.SetRow(scroll, 0)
        main_grid.Children.Add(scroll)
        
        # Button panel
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.VerticalAlignment = VerticalAlignment.Center
        button_panel.Margin = Thickness(10)
        
        # Select All button
        if self.multiselect:
            select_all_btn = Button()
            select_all_btn.Content = "Alles selecteren"
            select_all_btn.Width = 120
            select_all_btn.Height = 30
            select_all_btn.Margin = Thickness(5)
            select_all_btn.Click += self._select_all_click
            button_panel.Children.Add(select_all_btn)
            
            deselect_all_btn = Button()
            deselect_all_btn.Content = "Niets selecteren"
            deselect_all_btn.Width = 120
            deselect_all_btn.Height = 30
            deselect_all_btn.Margin = Thickness(5)
            deselect_all_btn.Click += self._deselect_all_click
            button_panel.Children.Add(deselect_all_btn)
        
        # OK button
        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 100
        ok_btn.Height = 30
        ok_btn.Margin = Thickness(5)
        ok_btn.Click += self._ok_click
        button_panel.Children.Add(ok_btn)
        
        # Cancel button
        cancel_btn = Button()
        cancel_btn.Content = "Annuleren"
        cancel_btn.Width = 100
        cancel_btn.Height = 30
        cancel_btn.Margin = Thickness(5)
        cancel_btn.Click += self._cancel_click
        button_panel.Children.Add(cancel_btn)
        
        Grid.SetRow(button_panel, 1)
        main_grid.Children.Add(button_panel)
        
        self.Content = main_grid
    
    def _create_item_panel(self, item):
        """Maak een panel voor een enkel item met thumbnail en checkbox"""
        border = Border()
        border.BorderBrush = Brushes.LightGray
        border.BorderThickness = Thickness(1)
        border.Margin = Thickness(5)
        border.Padding = Thickness(5)
        border.Width = 140
        border.Height = 160 if self.show_thumbnails else 50
        
        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        
        # Thumbnail
        if self.show_thumbnails:
            img = Image()
            img.Width = 120
            img.Height = 100
            img.Stretch = Stretch.Uniform
            img.Margin = Thickness(0, 0, 0, 5)
            
            thumb_path = item.get('thumbnail')
            if thumb_path and os.path.exists(thumb_path):
                try:
                    bitmap = BitmapImage()
                    bitmap.BeginInit()
                    bitmap.UriSource = Uri(thumb_path, UriKind.Absolute)
                    bitmap.CacheOption = BitmapCacheOption.OnLoad
                    bitmap.EndInit()
                    img.Source = bitmap
                except:
                    pass
            
            stack.Children.Add(img)
        
        # Checkbox met naam
        cb = CheckBox()
        cb.Content = item.get('name', 'Unknown')[:20]  # Truncate lange namen
        cb.Tag = item
        cb.ToolTip = item.get('name', '')  # Volledige naam in tooltip
        self.checkboxes.append(cb)
        stack.Children.Add(cb)
        
        border.Child = stack
        
        # Click op border selecteert ook
        border.MouseLeftButtonDown += lambda s, e: self._toggle_checkbox(cb)
        
        return border
    
    def _toggle_checkbox(self, cb):
        cb.IsChecked = not cb.IsChecked
    
    def _select_all_click(self, sender, e):
        for cb in self.checkboxes:
            cb.IsChecked = True
    
    def _deselect_all_click(self, sender, e):
        for cb in self.checkboxes:
            cb.IsChecked = False
    
    def _ok_click(self, sender, e):
        self.selected_items = []
        for cb in self.checkboxes:
            if cb.IsChecked:
                self.selected_items.append(cb.Tag)
        self.DialogResult = True
        self.Close()
    
    def _cancel_click(self, sender, e):
        self.DialogResult = False
        self.Close()


def select_export_folder(library_path):
    """Laat gebruiker een bestaande map kiezen of nieuwe aanmaken"""
    subfolders = get_subfolders(library_path)
    
    folder_options = [NEW_FOLDER_OPTION]
    folder_options.append("[Root] - Hoofdmap bibliotheek")
    
    for folder in subfolders:
        folder_options.append(folder)
    
    selected = forms.SelectFromList.show(
        folder_options,
        title="Selecteer doelmap voor export",
        multiselect=False,
        button_name="Selecteren"
    )
    
    if not selected:
        return None
    
    if selected == NEW_FOLDER_OPTION:
        new_folder_name = forms.ask_for_string(
            default="",
            prompt="Voer de naam in voor de nieuwe map:",
            title="Nieuwe map aanmaken"
        )
        
        if not new_folder_name:
            return None
        
        new_folder_path = os.path.join(library_path, new_folder_name)
        if not os.path.exists(new_folder_path):
            os.makedirs(new_folder_path)
            print("Nieuwe map aangemaakt: {}".format(new_folder_name))
        
        return new_folder_path
    
    elif selected.startswith("[Root]"):
        return library_path
    
    else:
        return os.path.join(library_path, selected)


def select_import_folder(library_path):
    """Laat gebruiker een map kiezen om uit te importeren"""
    subfolders = get_subfolders(library_path)
    
    folder_options = []
    
    root_rfa_count = len([f for f in os.listdir(library_path) if f.lower().endswith('.rfa')])
    if root_rfa_count > 0:
        folder_options.append("[Root] ({} families)".format(root_rfa_count))
    else:
        folder_options.append("[Root] - Hoofdmap")
    
    for folder in subfolders:
        folder_path = os.path.join(library_path, folder)
        rfa_count = len([f for f in os.listdir(folder_path) if f.lower().endswith('.rfa')])
        if rfa_count > 0:
            folder_options.append("{} ({} families)".format(folder, rfa_count))
        else:
            folder_options.append(folder)
    
    selected = forms.SelectFromList.show(
        folder_options,
        title="Selecteer bronmap voor import",
        multiselect=False,
        button_name="Selecteren"
    )
    
    if not selected:
        return None
    
    if selected.startswith("[Root]"):
        return library_path
    
    else:
        folder_name = selected.split(" (")[0]
        return os.path.join(library_path, folder_name)


def export_families():
    """Export geselecteerde families naar de bibliotheek"""
    library_path = ensure_library_exists()
    if not library_path:
        return
    
    families = get_families_in_project()
    
    if not families:
        forms.alert("Geen exporteerbare families gevonden in dit project.", exitscript=True)
    
    families = sorted(families, key=lambda x: x.Name)
    
    # Bouw items lijst voor thumbnail dialoog
    items = []
    for fam in families:
        items.append({
            'name': fam.Name,
            'family': fam,
            'thumbnail': None  # Revit families hebben geen makkelijk te extraheren thumbnails in-project
        })
    
    # Toon selectie dialoog (zonder thumbnails voor export, want die zijn moeilijk te krijgen)
    dialog = ThumbnailSelectDialog(
        items,
        title="Selecteer families om te exporteren",
        multiselect=True,
        show_thumbnails=False
    )
    
    result = dialog.ShowDialog()
    
    if not result or not dialog.selected_items:
        script.exit()
    
    selected_families = [item['family'] for item in dialog.selected_items]
    
    # Laat gebruiker doelmap kiezen
    export_path = select_export_folder(library_path)
    
    if not export_path:
        script.exit()
    
    if not os.path.exists(export_path):
        os.makedirs(export_path)
    
    # Export families
    exported_count = 0
    failed_exports = []
    
    print("")
    print("Exporteren naar: {}".format(export_path))
    print("-" * 60)
    
    for family in selected_families:
        fam_doc = None
        try:
            fam_doc = doc.EditFamily(family)
            
            if fam_doc:
                save_path = os.path.join(export_path, family.Name + ".rfa")
                save_options = DB.SaveAsOptions()
                save_options.OverwriteExistingFile = True
                
                fam_doc.SaveAs(save_path, save_options)
                fam_doc.Close(False)
                
                exported_count += 1
                print("  [OK] {}".format(family.Name))
            else:
                failed_exports.append(family.Name)
                print("  [FOUT] Kon niet openen: {}".format(family.Name))
            
        except Exception as e:
            failed_exports.append(family.Name)
            print("  [FOUT] {}: {}".format(family.Name, str(e)))
            if fam_doc and fam_doc.IsValidObject:
                try:
                    fam_doc.Close(False)
                except:
                    pass
    
    print("")
    print("=" * 60)
    output.print_md("## Export Voltooid")
    output.print_md("**Succesvol:** {} families".format(exported_count))
    output.print_md("**Locatie:** {}".format(export_path))
    
    if failed_exports:
        output.print_md("**Mislukt:** {}".format(len(failed_exports)))
        for name in failed_exports:
            print("  - {}".format(name))


def import_families():
    """Importeer families vanuit de bibliotheek met thumbnail preview"""
    library_path = ensure_library_exists()
    if not library_path:
        return
    
    source_path = select_import_folder(library_path)
    
    if not source_path:
        script.exit()
    
    # Verzamel families
    family_files = []
    
    if os.path.exists(source_path):
        for file in os.listdir(source_path):
            if file.lower().endswith('.rfa'):
                full_path = os.path.join(source_path, file)
                thumb_path = get_or_create_thumbnail(full_path)
                family_files.append({
                    'name': file.replace('.rfa', ''),
                    'path': full_path,
                    'thumbnail': thumb_path
                })
    
    if not family_files:
        forms.alert(
            "Geen families gevonden in de geselecteerde map!\n\nLocatie: {}".format(source_path),
            exitscript=True
        )
    
    # Sorteer op naam
    family_files = sorted(family_files, key=lambda x: x['name'])
    
    # Toon thumbnail selectie dialoog
    dialog = ThumbnailSelectDialog(
        family_files,
        title="Selecteer families om te importeren - {}".format(os.path.basename(source_path)),
        multiselect=True,
        show_thumbnails=True
    )
    
    result = dialog.ShowDialog()
    
    if not result or not dialog.selected_items:
        script.exit()
    
    # Import families
    imported_count = 0
    skipped_count = 0
    failed_imports = []
    
    print("")
    print("Importeren uit: {}".format(source_path))
    print("-" * 60)
    
    with revit.Transaction("Import Families"):
        for item in dialog.selected_items:
            try:
                family_loaded = doc.LoadFamily(item['path'])
                
                if family_loaded:
                    imported_count += 1
                    print("  [OK] {}".format(item['name']))
                else:
                    print("  [--] Al aanwezig: {}".format(item['name']))
                    skipped_count += 1
                    
            except Exception as e:
                failed_imports.append(item['name'])
                print("  [FOUT] {}: {}".format(item['name'], str(e)))
    
    print("")
    print("=" * 60)
    output.print_md("## Import Voltooid")
    output.print_md("**Nieuw geimporteerd:** {} families".format(imported_count))
    if skipped_count > 0:
        output.print_md("**Al aanwezig:** {} families".format(skipped_count))
    
    if failed_imports:
        output.print_md("**Mislukt:** {}".format(len(failed_imports)))
        for name in failed_imports:
            print("  - {}".format(name))


def main():
    """Main functie met keuzemenu"""
    
    library_path = get_library_path()
    
    if library_path:
        print("Bibliotheek locatie: {}".format(library_path))
        if os.path.exists(library_path):
            subfolders = get_subfolders(library_path)
            families = get_families_in_library()
            print("  {} submappen, {} families".format(len(subfolders), len(families)))
    else:
        print("Geen bibliotheek locatie geconfigureerd")
    print("-" * 60)
    
    options = [
        "Exporteer families naar bibliotheek",
        "Importeer families vanuit bibliotheek",
        "Open bibliotheek folder",
        "Configureer bibliotheek locatie",
        "Cache wissen (thumbnails)"
    ]
    
    selected = forms.CommandSwitchWindow.show(
        options,
        message="Kies een actie:"
    )
    
    if not selected:
        script.exit()
    
    if "Exporteer" in selected:
        export_families()
    elif "Importeer" in selected:
        import_families()
    elif "Open" in selected:
        lib_path = ensure_library_exists()
        if lib_path:
            os.startfile(lib_path)
    elif "Configureer" in selected:
        configure_library_path()
    elif "Cache" in selected:
        if os.path.exists(THUMB_CACHE):
            shutil.rmtree(THUMB_CACHE)
            forms.alert("Thumbnail cache gewist!")
        else:
            forms.alert("Geen cache om te wissen.")


if __name__ == "__main__":
    main()
