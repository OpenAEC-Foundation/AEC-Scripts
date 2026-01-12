# -*- coding: utf-8 -*-
"""Toolbar Manager - Beheer welke toolbars en knoppen worden geladen.

Hiermee kun je:
- Toolbars (panels) aan/uit zetten
- De volgorde van panels aanpassen
- Knoppen binnen panels herordenen
- Configuraties opslaan per gebruiker
"""

__title__ = "Toolbar\nManager"
__author__ = "3BM"
__doc__ = "Beheer toolbar zichtbaarheid en volgorde"

import os
import sys
import json
import clr
import shutil
from collections import OrderedDict

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System.Windows import Window, Application, ResourceDictionary, Thickness, MessageBox, MessageBoxButton, MessageBoxImage, WindowStartupLocation, VerticalAlignment, HorizontalAlignment
from System.Windows.Controls import (
    StackPanel, Grid, TextBlock, CheckBox, Button, ScrollViewer,
    Expander, ListBox, ListBoxItem, Border, DockPanel, Separator,
    ComboBox, ComboBoxItem, TabControl, TabItem, TextBox, Orientation, Dock
)
from System.Windows.Media import SolidColorBrush, Color, Colors, BrushConverter
from System.Windows import FontWeights
from System.Windows.Input import MouseButtonState

# pyRevit imports
from pyrevit import script, forms, HOST_APP
from pyrevit.coreutils import yaml

# Configuratie pad
CONFIG_FOLDER = os.path.join(os.environ.get('APPDATA', ''), '3BM', 'ToolbarManager')
CONFIG_FILE = os.path.join(CONFIG_FOLDER, 'toolbar_config.json')

# 3BM Kleuren
JMK_BLAUW = "#003F5F"
JMK_BLAUW_LIGHT = "#005580"
JMK_GRIJS = "#4A4A4A"
JMK_WIT = "#FFFFFF"
JMK_ACCENT = "#0078D4"


class ExtensionInfo:
    """Informatie over een pyRevit extensie."""
    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path).replace('.extension', '')
        self.tabs = []
        self.enabled = True
        self._scan_tabs()
    
    def _scan_tabs(self):
        """Scan alle tabs in deze extensie."""
        for item in os.listdir(self.path):
            if item.endswith('.tab'):
                tab_path = os.path.join(self.path, item)
                if os.path.isdir(tab_path):
                    self.tabs.append(TabInfo(tab_path))


class TabInfo:
    """Informatie over een tab."""
    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path).replace('.tab', '')
        self.panels = []
        self.panel_order = []
        self.enabled = True
        self._scan_panels()
        self._load_bundle_order()
    
    def _scan_panels(self):
        """Scan alle panels in deze tab."""
        for item in os.listdir(self.path):
            if item.endswith('.panel'):
                panel_path = os.path.join(self.path, item)
                if os.path.isdir(panel_path):
                    self.panels.append(PanelInfo(panel_path))
    
    def _load_bundle_order(self):
        """Laad de volgorde uit bundle.yaml."""
        bundle_path = os.path.join(self.path, 'bundle.yaml')
        if os.path.exists(bundle_path):
            try:
                with open(bundle_path, 'r') as f:
                    data = yaml.safe_load(f)
                    if data and 'layout' in data:
                        self.panel_order = data['layout']
            except:
                pass
        
        # Als geen order, gebruik alfabetisch
        if not self.panel_order:
            self.panel_order = [p.name for p in self.panels]
    
    def save_bundle_order(self):
        """Sla de volgorde op in bundle.yaml."""
        bundle_path = os.path.join(self.path, 'bundle.yaml')
        data = {'layout': self.panel_order}
        try:
            with open(bundle_path, 'w') as f:
                yaml.safe_dump(data, f, default_flow_style=False)
            return True
        except Exception as e:
            print("Fout bij opslaan bundle.yaml: {}".format(e))
            return False


class PanelInfo:
    """Informatie over een panel."""
    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path).replace('.panel', '')
        self.buttons = []
        self.button_order = []
        self.enabled = True
        self._scan_buttons()
        self._load_bundle_order()
    
    def _scan_buttons(self):
        """Scan alle knoppen in dit panel."""
        for item in os.listdir(self.path):
            item_path = os.path.join(self.path, item)
            if os.path.isdir(item_path):
                if item.endswith('.pushbutton'):
                    self.buttons.append(ButtonInfo(item_path, 'pushbutton'))
                elif item.endswith('.pulldown'):
                    self.buttons.append(ButtonInfo(item_path, 'pulldown'))
                elif item.endswith('.splitbutton'):
                    self.buttons.append(ButtonInfo(item_path, 'splitbutton'))
                elif item.endswith('.stack'):
                    self.buttons.append(ButtonInfo(item_path, 'stack'))
    
    def _load_bundle_order(self):
        """Laad de volgorde uit bundle.yaml."""
        bundle_path = os.path.join(self.path, 'bundle.yaml')
        if os.path.exists(bundle_path):
            try:
                with open(bundle_path, 'r') as f:
                    data = yaml.safe_load(f)
                    if data and 'layout' in data:
                        self.button_order = data['layout']
            except:
                pass
        
        if not self.button_order:
            self.button_order = [b.name for b in self.buttons]
    
    def save_bundle_order(self):
        """Sla de volgorde op in bundle.yaml."""
        bundle_path = os.path.join(self.path, 'bundle.yaml')
        data = {'layout': self.button_order}
        try:
            with open(bundle_path, 'w') as f:
                yaml.safe_dump(data, f, default_flow_style=False)
            return True
        except Exception as e:
            print("Fout bij opslaan bundle.yaml: {}".format(e))
            return False


class ButtonInfo:
    """Informatie over een knop."""
    def __init__(self, path, button_type):
        self.path = path
        self.button_type = button_type
        self.folder_name = os.path.basename(path)
        self.name = self.folder_name.split('.')[0]
        self.enabled = True
        self.title = self._get_title()
    
    def _get_title(self):
        """Haal de titel uit het script."""
        script_path = os.path.join(self.path, 'script.py')
        if os.path.exists(script_path):
            try:
                with open(script_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    for line in content.split('\n'):
                        if '__title__' in line:
                            # Extract title value
                            if '=' in line:
                                title = line.split('=')[1].strip()
                                title = title.strip('"').strip("'")
                                return title.replace('\\n', ' ')
            except:
                pass
        return self.name


class ToolbarManagerWindow(Window):
    """Hoofdvenster voor de Toolbar Manager."""
    
    def __init__(self, extensions):
        self.extensions = extensions
        self.config = self._load_config()
        self.selected_tab = None
        self.selected_panel = None
        self._init_window()
        self._apply_config()
    
    def _init_window(self):
        """Initialiseer het venster."""
        self.Title = "3BM Toolbar Manager"
        self.Width = 900
        self.Height = 700
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = BrushConverter().ConvertFrom(JMK_WIT)
        
        # Hoofdcontainer
        main_grid = Grid()
        main_grid.RowDefinitions.Add(self._row_def("Auto"))  # Header
        main_grid.RowDefinitions.Add(self._row_def("*"))     # Content
        main_grid.RowDefinitions.Add(self._row_def("Auto"))  # Footer
        
        # Header
        header = self._create_header()
        Grid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        # Content met tabs
        content = self._create_content()
        Grid.SetRow(content, 1)
        main_grid.Children.Add(content)
        
        # Footer met knoppen
        footer = self._create_footer()
        Grid.SetRow(footer, 2)
        main_grid.Children.Add(footer)
        
        self.Content = main_grid
    
    def _row_def(self, height):
        """Maak een RowDefinition."""
        from System.Windows.Controls import RowDefinition
        from System.Windows import GridLength, GridUnitType
        rd = RowDefinition()
        if height == "*":
            rd.Height = GridLength(1, GridUnitType.Star)
        elif height == "Auto":
            rd.Height = GridLength.Auto
        else:
            rd.Height = GridLength(float(height))
        return rd
    
    def _col_def(self, width):
        """Maak een ColumnDefinition."""
        from System.Windows.Controls import ColumnDefinition
        from System.Windows import GridLength, GridUnitType
        cd = ColumnDefinition()
        if width == "*":
            cd.Width = GridLength(1, GridUnitType.Star)
        elif width == "Auto":
            cd.Width = GridLength.Auto
        else:
            cd.Width = GridLength(float(width))
        return cd
    
    def _create_header(self):
        """Maak de header."""
        border = Border()
        border.Background = BrushConverter().ConvertFrom(JMK_BLAUW)
        border.Padding = Thickness(20, 15, 20, 15)
        
        stack = StackPanel()
        stack.Orientation = Orientation.Horizontal
        
        title = TextBlock()
        title.Text = "Toolbar Manager"
        title.FontSize = 24
        title.FontWeight = FontWeights.Bold
        title.Foreground = BrushConverter().ConvertFrom(JMK_WIT)
        stack.Children.Add(title)
        
        subtitle = TextBlock()
        subtitle.Text = "  |  Beheer je pyRevit toolbars"
        subtitle.FontSize = 14
        subtitle.Foreground = BrushConverter().ConvertFrom("#AACCDD")
        subtitle.VerticalAlignment = VerticalAlignment.Center
        subtitle.Margin = Thickness(10, 0, 0, 0)
        stack.Children.Add(subtitle)
        
        border.Child = stack
        return border
    
    def _create_content(self):
        """Maak de content area met tabs."""
        tab_control = TabControl()
        tab_control.Margin = Thickness(15)
        
        # Tab 1: Extensies & Tabs
        tab1 = TabItem()
        tab1.Header = "Extensies & Tabs"
        tab1.Content = self._create_extensions_tab()
        tab_control.Items.Add(tab1)
        
        # Tab 2: Panels
        tab2 = TabItem()
        tab2.Header = "Panels"
        tab2.Content = self._create_panels_tab()
        tab_control.Items.Add(tab2)
        
        # Tab 3: Knoppen
        tab3 = TabItem()
        tab3.Header = "Knoppen"
        tab3.Content = self._create_buttons_tab()
        tab_control.Items.Add(tab3)
        
        return tab_control
    
    def _create_extensions_tab(self):
        """Maak de extensies tab."""
        grid = Grid()
        grid.ColumnDefinitions.Add(self._col_def("*"))
        grid.ColumnDefinitions.Add(self._col_def("*"))
        
        # Linker kolom: Extensies
        left_panel = self._create_section("Extensies", "Selecteer welke extensies geladen worden")
        self.extensions_list = ListBox()
        self.extensions_list.Margin = Thickness(10)
        self.extensions_list.SelectionChanged += self._on_extension_selected
        
        for ext in self.extensions:
            item = self._create_checkbox_item(ext.name, ext.enabled, ext)
            self.extensions_list.Items.Add(item)
        
        left_panel.Children.Add(self.extensions_list)
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        # Rechter kolom: Tabs van geselecteerde extensie
        right_panel = self._create_section("Tabs", "Tabs in geselecteerde extensie")
        self.tabs_list = ListBox()
        self.tabs_list.Margin = Thickness(10)
        right_panel.Children.Add(self.tabs_list)
        Grid.SetColumn(right_panel, 1)
        grid.Children.Add(right_panel)
        
        return grid
    
    def _create_panels_tab(self):
        """Maak de panels tab."""
        grid = Grid()
        grid.ColumnDefinitions.Add(self._col_def("300"))
        grid.ColumnDefinitions.Add(self._col_def("*"))
        grid.ColumnDefinitions.Add(self._col_def("120"))
        
        # Links: Tab selectie
        left_panel = self._create_section("Selecteer Tab", "")
        self.tab_selector = ComboBox()
        self.tab_selector.Margin = Thickness(10)
        self.tab_selector.SelectionChanged += self._on_tab_selected_for_panels
        
        # Vul met alle tabs
        for ext in self.extensions:
            for tab in ext.tabs:
                item = ComboBoxItem()
                item.Content = "{} > {}".format(ext.name, tab.name)
                item.Tag = tab
                self.tab_selector.Items.Add(item)
        
        left_panel.Children.Add(self.tab_selector)
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        # Midden: Panels lijst
        middle_panel = self._create_section("Panels", "Volgorde en zichtbaarheid")
        self.panels_list = ListBox()
        self.panels_list.Margin = Thickness(10)
        middle_panel.Children.Add(self.panels_list)
        Grid.SetColumn(middle_panel, 1)
        grid.Children.Add(middle_panel)
        
        # Rechts: Sorteer knoppen
        right_panel = StackPanel()
        right_panel.VerticalAlignment = VerticalAlignment.Center
        right_panel.Margin = Thickness(10)
        
        btn_up = Button()
        btn_up.Content = "▲ Omhoog"
        btn_up.Margin = Thickness(0, 5, 0, 5)
        btn_up.Padding = Thickness(10, 8, 10, 8)
        btn_up.Click += self._move_panel_up
        right_panel.Children.Add(btn_up)
        
        btn_down = Button()
        btn_down.Content = "▼ Omlaag"
        btn_down.Margin = Thickness(0, 5, 0, 5)
        btn_down.Padding = Thickness(10, 8, 10, 8)
        btn_down.Click += self._move_panel_down
        right_panel.Children.Add(btn_down)
        
        Grid.SetColumn(right_panel, 2)
        grid.Children.Add(right_panel)
        
        return grid
    
    def _create_buttons_tab(self):
        """Maak de knoppen tab."""
        grid = Grid()
        grid.ColumnDefinitions.Add(self._col_def("300"))
        grid.ColumnDefinitions.Add(self._col_def("*"))
        grid.ColumnDefinitions.Add(self._col_def("120"))
        
        # Links: Panel selectie
        left_panel = self._create_section("Selecteer Panel", "")
        self.panel_selector = ComboBox()
        self.panel_selector.Margin = Thickness(10)
        self.panel_selector.SelectionChanged += self._on_panel_selected_for_buttons
        
        # Vul met alle panels
        for ext in self.extensions:
            for tab in ext.tabs:
                for panel in tab.panels:
                    item = ComboBoxItem()
                    item.Content = "{} > {} > {}".format(ext.name, tab.name, panel.name)
                    item.Tag = panel
                    self.panel_selector.Items.Add(item)
        
        left_panel.Children.Add(self.panel_selector)
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        # Midden: Knoppen lijst
        middle_panel = self._create_section("Knoppen", "Volgorde")
        self.buttons_list = ListBox()
        self.buttons_list.Margin = Thickness(10)
        middle_panel.Children.Add(self.buttons_list)
        Grid.SetColumn(middle_panel, 1)
        grid.Children.Add(middle_panel)
        
        # Rechts: Sorteer knoppen
        right_panel = StackPanel()
        right_panel.VerticalAlignment = VerticalAlignment.Center
        right_panel.Margin = Thickness(10)
        
        btn_up = Button()
        btn_up.Content = "▲ Omhoog"
        btn_up.Margin = Thickness(0, 5, 0, 5)
        btn_up.Padding = Thickness(10, 8, 10, 8)
        btn_up.Click += self._move_button_up
        right_panel.Children.Add(btn_up)
        
        btn_down = Button()
        btn_down.Content = "▼ Omlaag"
        btn_down.Margin = Thickness(0, 5, 0, 5)
        btn_down.Padding = Thickness(10, 8, 10, 8)
        btn_down.Click += self._move_button_down
        right_panel.Children.Add(btn_down)
        
        Grid.SetColumn(right_panel, 2)
        grid.Children.Add(right_panel)
        
        return grid
    
    def _create_section(self, title, description):
        """Maak een sectie met titel."""
        stack = StackPanel()
        stack.Margin = Thickness(10)
        
        header = TextBlock()
        header.Text = title
        header.FontSize = 16
        header.FontWeight = FontWeights.Bold
        header.Foreground = BrushConverter().ConvertFrom(JMK_BLAUW)
        stack.Children.Add(header)
        
        if description:
            desc = TextBlock()
            desc.Text = description
            desc.FontSize = 11
            desc.Foreground = BrushConverter().ConvertFrom(JMK_GRIJS)
            desc.Margin = Thickness(0, 2, 0, 10)
            stack.Children.Add(desc)
        
        return stack
    
    def _create_checkbox_item(self, text, is_checked, tag):
        """Maak een listbox item met checkbox."""
        stack = StackPanel()
        stack.Orientation = Orientation.Horizontal
        stack.Margin = Thickness(5)
        
        cb = CheckBox()
        cb.IsChecked = is_checked
        cb.Tag = tag
        cb.Checked += self._on_item_checked
        cb.Unchecked += self._on_item_unchecked
        stack.Children.Add(cb)
        
        txt = TextBlock()
        txt.Text = text
        txt.Margin = Thickness(8, 0, 0, 0)
        txt.VerticalAlignment = VerticalAlignment.Center
        stack.Children.Add(txt)
        
        item = ListBoxItem()
        item.Content = stack
        item.Tag = tag
        return item
    
    def _create_footer(self):
        """Maak de footer met knoppen."""
        border = Border()
        border.Background = BrushConverter().ConvertFrom("#F5F5F5")
        border.Padding = Thickness(20, 15, 20, 15)
        
        dock = DockPanel()
        dock.LastChildFill = False
        
        # Info links
        info = TextBlock()
        info.Text = "Wijzigingen worden opgeslagen in bundle.yaml bestanden"
        info.FontSize = 11
        info.Foreground = BrushConverter().ConvertFrom(JMK_GRIJS)
        info.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(info, Dock.Left)
        dock.Children.Add(info)
        
        # Knoppen rechts
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        DockPanel.SetDock(btn_panel, Dock.Right)
        
        btn_cancel = Button()
        btn_cancel.Content = "Annuleren"
        btn_cancel.Padding = Thickness(20, 10, 20, 10)
        btn_cancel.Margin = Thickness(0, 0, 10, 0)
        btn_cancel.Click += self._on_cancel
        btn_panel.Children.Add(btn_cancel)
        
        btn_save = Button()
        btn_save.Content = "Opslaan & Sluiten"
        btn_save.Padding = Thickness(20, 10, 20, 10)
        btn_save.Background = BrushConverter().ConvertFrom(JMK_BLAUW)
        btn_save.Foreground = BrushConverter().ConvertFrom(JMK_WIT)
        btn_save.Click += self._on_save
        btn_panel.Children.Add(btn_save)
        
        dock.Children.Add(btn_panel)
        
        border.Child = dock
        return border
    
    # Event handlers
    def _on_extension_selected(self, sender, args):
        """Handler voor extensie selectie."""
        if sender.SelectedItem:
            ext = sender.SelectedItem.Tag
            self.tabs_list.Items.Clear()
            for tab in ext.tabs:
                item = self._create_checkbox_item(tab.name, tab.enabled, tab)
                self.tabs_list.Items.Add(item)
    
    def _on_tab_selected_for_panels(self, sender, args):
        """Handler voor tab selectie in panels tab."""
        if sender.SelectedItem:
            tab = sender.SelectedItem.Tag
            self.selected_tab = tab
            self._refresh_panels_list()
    
    def _refresh_panels_list(self):
        """Ververs de panels lijst."""
        if not self.selected_tab:
            return
        
        self.panels_list.Items.Clear()
        
        # Sorteer panels op basis van panel_order
        ordered_panels = []
        for name in self.selected_tab.panel_order:
            for panel in self.selected_tab.panels:
                if panel.name == name:
                    ordered_panels.append(panel)
                    break
        
        # Voeg panels toe die niet in order staan
        for panel in self.selected_tab.panels:
            if panel not in ordered_panels:
                ordered_panels.append(panel)
        
        for panel in ordered_panels:
            item = self._create_checkbox_item(panel.name, panel.enabled, panel)
            self.panels_list.Items.Add(item)
    
    def _on_panel_selected_for_buttons(self, sender, args):
        """Handler voor panel selectie in buttons tab."""
        if sender.SelectedItem:
            panel = sender.SelectedItem.Tag
            self.selected_panel = panel
            self._refresh_buttons_list()
    
    def _refresh_buttons_list(self):
        """Ververs de buttons lijst."""
        if not self.selected_panel:
            return
        
        self.buttons_list.Items.Clear()
        
        # Sorteer buttons op basis van button_order
        ordered_buttons = []
        for name in self.selected_panel.button_order:
            for btn in self.selected_panel.buttons:
                if btn.name == name:
                    ordered_buttons.append(btn)
                    break
        
        # Voeg buttons toe die niet in order staan
        for btn in self.selected_panel.buttons:
            if btn not in ordered_buttons:
                ordered_buttons.append(btn)
        
        for btn in ordered_buttons:
            item = ListBoxItem()
            stack = StackPanel()
            stack.Orientation = Orientation.Horizontal
            stack.Margin = Thickness(5)
            
            # Type indicator
            type_txt = TextBlock()
            type_txt.Text = "[{}]".format(btn.button_type[:4])
            type_txt.FontSize = 10
            type_txt.Foreground = BrushConverter().ConvertFrom(JMK_GRIJS)
            type_txt.Width = 50
            stack.Children.Add(type_txt)
            
            # Naam
            name_txt = TextBlock()
            name_txt.Text = btn.title or btn.name
            name_txt.Margin = Thickness(5, 0, 0, 0)
            stack.Children.Add(name_txt)
            
            item.Content = stack
            item.Tag = btn
            self.buttons_list.Items.Add(item)
    
    def _move_panel_up(self, sender, args):
        """Verplaats geselecteerd panel omhoog."""
        if not self.selected_tab or self.panels_list.SelectedIndex < 1:
            return
        
        idx = self.panels_list.SelectedIndex
        order = self.selected_tab.panel_order
        if idx < len(order):
            order[idx], order[idx-1] = order[idx-1], order[idx]
            self._refresh_panels_list()
            self.panels_list.SelectedIndex = idx - 1
    
    def _move_panel_down(self, sender, args):
        """Verplaats geselecteerd panel omlaag."""
        if not self.selected_tab or self.panels_list.SelectedIndex < 0:
            return
        
        idx = self.panels_list.SelectedIndex
        order = self.selected_tab.panel_order
        if idx < len(order) - 1:
            order[idx], order[idx+1] = order[idx+1], order[idx]
            self._refresh_panels_list()
            self.panels_list.SelectedIndex = idx + 1
    
    def _move_button_up(self, sender, args):
        """Verplaats geselecteerde button omhoog."""
        if not self.selected_panel or self.buttons_list.SelectedIndex < 1:
            return
        
        idx = self.buttons_list.SelectedIndex
        order = self.selected_panel.button_order
        if idx < len(order):
            order[idx], order[idx-1] = order[idx-1], order[idx]
            self._refresh_buttons_list()
            self.buttons_list.SelectedIndex = idx - 1
    
    def _move_button_down(self, sender, args):
        """Verplaats geselecteerde button omlaag."""
        if not self.selected_panel or self.buttons_list.SelectedIndex < 0:
            return
        
        idx = self.buttons_list.SelectedIndex
        order = self.selected_panel.button_order
        if idx < len(order) - 1:
            order[idx], order[idx+1] = order[idx+1], order[idx]
            self._refresh_buttons_list()
            self.buttons_list.SelectedIndex = idx + 1
    
    def _on_item_checked(self, sender, args):
        """Handler voor checkbox aanvinken."""
        if sender.Tag:
            sender.Tag.enabled = True
    
    def _on_item_unchecked(self, sender, args):
        """Handler voor checkbox uitvinken."""
        if sender.Tag:
            sender.Tag.enabled = False
    
    def _on_cancel(self, sender, args):
        """Annuleer en sluit."""
        self.Close()
    
    def _on_save(self, sender, args):
        """Sla wijzigingen op."""
        saved_count = 0
        errors = []
        
        for ext in self.extensions:
            for tab in ext.tabs:
                try:
                    if tab.save_bundle_order():
                        saved_count += 1
                except Exception as e:
                    errors.append("Tab {}: {}".format(tab.name, str(e)))
                
                for panel in tab.panels:
                    try:
                        if panel.save_bundle_order():
                            saved_count += 1
                    except Exception as e:
                        errors.append("Panel {}: {}".format(panel.name, str(e)))
        
        # Sla config op
        self._save_config()
        
        if errors:
            MessageBox.Show(
                "Opgeslagen met {} fouten:\n{}".format(len(errors), "\n".join(errors[:5])),
                "Waarschuwing",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            )
        else:
            MessageBox.Show(
                "Wijzigingen opgeslagen!\n\nHerlaad pyRevit om de wijzigingen te zien.",
                "Succes",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
        
        self.Close()
    
    def _load_config(self):
        """Laad configuratie."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {}
    
    def _save_config(self):
        """Sla configuratie op."""
        config = {
            'disabled_extensions': [],
            'disabled_tabs': [],
            'disabled_panels': []
        }
        
        for ext in self.extensions:
            if not ext.enabled:
                config['disabled_extensions'].append(ext.name)
            for tab in ext.tabs:
                if not tab.enabled:
                    config['disabled_tabs'].append("{}/{}".format(ext.name, tab.name))
                for panel in tab.panels:
                    if not panel.enabled:
                        config['disabled_panels'].append(
                            "{}/{}/{}".format(ext.name, tab.name, panel.name)
                        )
        
        if not os.path.exists(CONFIG_FOLDER):
            os.makedirs(CONFIG_FOLDER)
        
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    
    def _apply_config(self):
        """Pas opgeslagen configuratie toe."""
        for ext in self.extensions:
            if ext.name in self.config.get('disabled_extensions', []):
                ext.enabled = False
            for tab in ext.tabs:
                key = "{}/{}".format(ext.name, tab.name)
                if key in self.config.get('disabled_tabs', []):
                    tab.enabled = False
                for panel in tab.panels:
                    key = "{}/{}/{}".format(ext.name, tab.name, panel.name)
                    if key in self.config.get('disabled_panels', []):
                        panel.enabled = False


def scan_extensions():
    """Scan alle pyRevit extensies."""
    extensions = []
    
    # Standaard pyRevit extensies
    pyrevit_extensions = r"C:\Program Files\pyRevit-Master\extensions"
    if os.path.exists(pyrevit_extensions):
        for item in os.listdir(pyrevit_extensions):
            if item.endswith('.extension'):
                ext_path = os.path.join(pyrevit_extensions, item)
                if os.path.isdir(ext_path):
                    extensions.append(ExtensionInfo(ext_path))
    
    # Zoek ook in AppData
    appdata_extensions = os.path.join(
        os.environ.get('APPDATA', ''), 
        'pyRevit-Master', 
        'Extensions'
    )
    if os.path.exists(appdata_extensions):
        for item in os.listdir(appdata_extensions):
            if item.endswith('.extension'):
                ext_path = os.path.join(appdata_extensions, item)
                if os.path.isdir(ext_path):
                    extensions.append(ExtensionInfo(ext_path))
    
    return extensions


# Main
if __name__ == '__main__':
    try:
        extensions = scan_extensions()
        
        if not extensions:
            forms.alert(
                "Geen pyRevit extensies gevonden!",
                title="Toolbar Manager"
            )
        else:
            window = ToolbarManagerWindow(extensions)
            window.ShowDialog()
    except Exception as e:
        forms.alert(
            "Fout: {}".format(str(e)),
            title="Toolbar Manager"
        )
