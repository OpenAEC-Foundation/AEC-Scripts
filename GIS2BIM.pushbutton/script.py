# -*- coding: utf-8 -*-
"""
GIS2BIM - Nederlandse Geodata voor Revit v6.4 (Standalone)
"""

__title__ = "GIS2BIM"
__author__ = "3BM / JMK"
__doc__ = "Nederlandse GIS data workflow - Locatie, Selectie, Import"

import clr
import os
import math
import json
import tempfile
import time

clr.AddReference('System')
clr.AddReference('System.Net')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Net import WebClient, WebRequest
from System.Text import Encoding
import System.Drawing as Drawing
from System.Drawing import Bitmap, Graphics, Pen, SolidBrush, Rectangle, Font, FontStyle
from System.Drawing import Color as DrawingColor
from System.Drawing.Drawing2D import SmoothingMode, DashStyle
from System.Drawing.Imaging import ImageFormat
from System.Collections.Generic import List

import System.Windows.Forms as WinForms
from System.Windows.Forms import (
    Application, DialogResult, ListViewItem, PictureBox,
    PictureBoxSizeMode, BorderStyle, Cursors, MouseButtons,
    FormBorderStyle, FormStartPosition,
    Label, TextBox, Button, ComboBox, NumericUpDown, CheckBox,
    ListView, ColumnHeader, ProgressBar,
    GroupBox, FlatStyle, AnchorStyles, HorizontalAlignment
)
# Note: View and Panel not imported directly to avoid conflict with Autodesk.Revit.DB - use WinForms.View and WinForms.Panel instead

# Helper functions to avoid Point/Size struct instantiation issues in IronPython
def set_location(control, x, y):
    """Set control location using Left and Top properties instead of Point."""
    control.Left = int(x)
    control.Top = int(y)

def set_size(control, width, height):
    """Set control size using Width and Height properties instead of Size."""
    control.Width = int(width)
    control.Height = int(height)

def set_bounds(control, x, y, width, height):
    """Set control bounds using individual properties."""
    control.Left = int(x)
    control.Top = int(y)
    control.Width = int(width)
    control.Height = int(height)

from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog


# =============================================================================
# UI STYLING (gebruik System.Drawing.Color)
# =============================================================================

# Define colors and fonts as module-level variables (lazy initialization)
_colors_initialized = False
_fonts_initialized = False

COLOR_PRIMARY = None
COLOR_PRIMARY_DARK = None
COLOR_BACKGROUND = None
COLOR_SURFACE = None
COLOR_TEXT_PRIMARY = None
COLOR_TEXT_SECONDARY = None
COLOR_BORDER = None

FONT_HEADER = None
FONT_NORMAL = None
FONT_BUTTON = None

def init_colors():
    global _colors_initialized, COLOR_PRIMARY, COLOR_PRIMARY_DARK, COLOR_BACKGROUND
    global COLOR_SURFACE, COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_BORDER
    if not _colors_initialized:
        COLOR_PRIMARY = DrawingColor.FromArgb(0, 122, 204)
        COLOR_PRIMARY_DARK = DrawingColor.FromArgb(0, 92, 153)
        COLOR_BACKGROUND = DrawingColor.FromArgb(250, 250, 250)
        COLOR_SURFACE = DrawingColor.White
        COLOR_TEXT_PRIMARY = DrawingColor.FromArgb(33, 33, 33)
        COLOR_TEXT_SECONDARY = DrawingColor.FromArgb(117, 117, 117)
        COLOR_BORDER = DrawingColor.FromArgb(200, 200, 200)
        _colors_initialized = True

def init_fonts():
    global _fonts_initialized, FONT_HEADER, FONT_NORMAL, FONT_BUTTON
    if not _fonts_initialized:
        FONT_HEADER = Font("Segoe UI", 11, FontStyle.Bold)
        FONT_NORMAL = Font("Segoe UI", 9, FontStyle.Regular)
        FONT_BUTTON = Font("Segoe UI", 9, FontStyle.Regular)
        _fonts_initialized = True

def init_ui():
    init_colors()
    init_fonts()


class DPIScaler:
    _scale_factor = None
    
    @classmethod
    def get_scale_factor(cls):
        if cls._scale_factor is None:
            try:
                import ctypes
                user32 = ctypes.windll.user32
                user32.SetProcessDPIAware()
                dc = user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)
                user32.ReleaseDC(0, dc)
                cls._scale_factor = dpi / 96.0
            except:
                cls._scale_factor = 1.0
        return cls._scale_factor
    
    @classmethod
    def scale(cls, value):
        return int(value * cls.get_scale_factor())


class UIFactory:
    @staticmethod
    def create_label(text, bold=False, color=None):
        lbl = Label()
        lbl.Text = text
        lbl.Font = FONT_HEADER if bold else FONT_NORMAL
        lbl.ForeColor = color if color else COLOR_TEXT_PRIMARY
        lbl.AutoSize = True
        return lbl

    @staticmethod
    def create_textbox(multiline=False, readonly=False):
        txt = TextBox()
        txt.Font = FONT_NORMAL
        txt.Multiline = multiline
        txt.ReadOnly = readonly
        txt.BorderStyle = BorderStyle.FixedSingle
        if readonly:
            txt.BackColor = COLOR_BACKGROUND
        return txt

    @staticmethod
    def create_button(text, primary=True, width=100):
        btn = Button()
        btn.Text = text
        btn.Font = FONT_BUTTON
        set_size(btn, DPIScaler.scale(width), DPIScaler.scale(28))
        btn.FlatStyle = FlatStyle.Flat
        btn.Cursor = Cursors.Hand
        if primary:
            btn.BackColor = COLOR_PRIMARY
            btn.ForeColor = DrawingColor.White
            btn.FlatAppearance.BorderColor = COLOR_PRIMARY_DARK
        else:
            btn.BackColor = COLOR_SURFACE
            btn.ForeColor = COLOR_TEXT_PRIMARY
            btn.FlatAppearance.BorderColor = COLOR_BORDER
        return btn

    @staticmethod
    def create_numeric(min_val=0, max_val=100, default=0, decimals=0):
        num = NumericUpDown()
        num.Font = FONT_NORMAL
        num.Minimum = min_val
        num.Maximum = max_val
        num.Value = default
        num.DecimalPlaces = decimals
        set_size(num, DPIScaler.scale(100), DPIScaler.scale(25))
        return num

    @staticmethod
    def create_listview(columns):
        lv = ListView()
        lv.View = WinForms.View.Details  # Use full path to avoid conflict with Revit's View
        lv.FullRowSelect = True
        lv.GridLines = True
        lv.Font = FONT_NORMAL
        lv.BorderStyle = BorderStyle.FixedSingle
        for name, width in columns:
            col = ColumnHeader()
            col.Text = name
            col.Width = DPIScaler.scale(width)
            lv.Columns.Add(col)
        return lv

    @staticmethod
    def create_progressbar(width=200):
        pb = ProgressBar()
        set_size(pb, DPIScaler.scale(width), DPIScaler.scale(20))
        pb.Style = WinForms.ProgressBarStyle.Continuous
        return pb

    @staticmethod
    def create_separator(width=200):
        sep = WinForms.Panel()  # Use full path to avoid conflict with Revit's Panel
        set_size(sep, DPIScaler.scale(width), 1)
        sep.BackColor = COLOR_BORDER
        return sep


def show_info(message, title="Informatie"):
    WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)

def show_warning(message, title="Waarschuwing"):
    WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)

def show_error(message, title="Fout"):
    WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error)

def show_question(message, title="Bevestigen"):
    result = WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Question)
    return result == WinForms.DialogResult.Yes


# =============================================================================
# CONSTANTEN
# =============================================================================

FEET_TO_METERS = 0.3048
METERS_TO_FEET = 1 / FEET_TO_METERS
WMTS_XCORNER = -285401.92
WMTS_YCORNER = 903402.0
WMTS_PIXEL_WIDTH = 256

ZOOMLEVEL_RESOLUTIONS = {
    0: 3440.640, 1: 1720.320, 2: 860.160, 3: 430.080,
    4: 215.040, 5: 107.520, 6: 53.720, 7: 26.880,
    8: 13.440, 9: 6.720, 10: 3.360, 11: 1.680,
    12: 0.840, 13: 0.420, 14: 0.210, 15: 0.105, 16: 0.0575
}

DEFAULT_RD_X = 99628
DEFAULT_RD_Y = 424889
DEFAULT_ADDRESS = "Burgemeester de Raadtsingel 31, Dordrecht"


# =============================================================================
# GIS DATA LAGEN
# =============================================================================

GIS_LAYERS = {
    'luchtfoto_actueel': {'name': 'Luchtfoto Actueel', 'category': 'Rasterkaarten', 'type': 'wmts', 'url': 'https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0', 'layer': 'Actueel_orthoHR', 'zoom': 14, 'sheet_name': 'GIS - Luchtfoto Actueel'},
    'luchtfoto_2023': {'name': 'Luchtfoto 2023', 'category': 'Rasterkaarten', 'type': 'wmts', 'url': 'https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0', 'layer': '2023_orthoHR', 'zoom': 14, 'sheet_name': 'GIS - Luchtfoto 2023'},
    'top10nl': {'name': 'Top10NL Achtergrond', 'category': 'Rasterkaarten', 'type': 'wmts', 'url': 'https://service.pdok.nl/brt/achtergrondkaart/wmts/v2_0', 'layer': 'standaard', 'zoom': 11, 'sheet_name': 'GIS - Top10NL'},
    'bestemmingsplan': {'name': 'Bestemmingsplan', 'category': 'Rasterkaarten', 'type': 'wms', 'url': 'https://service.pdok.nl/kadaster/plu/wms/v1_0', 'layer': 'Bestemmingsplangebied', 'sheet_name': 'GIS - Bestemmingsplan'},
    'bouwvlak_wms': {'name': 'Bouwvlak (kaart)', 'category': 'Rasterkaarten', 'type': 'wms', 'url': 'https://service.pdok.nl/kadaster/plu/wms/v1_0', 'layer': 'Bouwvlak', 'sheet_name': 'GIS - Bouwvlak'},
    'bgt_achtergrond': {'name': 'BGT Achtergrondkaart', 'category': 'Rasterkaarten', 'type': 'wmts', 'url': 'https://service.pdok.nl/lv/bgt/wmts/v1_0', 'layer': 'standaardvisualisatie', 'zoom': 14, 'sheet_name': 'GIS - BGT'},
    'kadaster_kaart': {'name': 'Kadastrale Kaart', 'category': 'Rasterkaarten', 'type': 'wms', 'url': 'https://service.pdok.nl/kadaster/kadastralekaart/wms/v5_0', 'layer': 'Kadastralekaart', 'sheet_name': 'GIS - Kadaster Kaart'},
    'natura2000': {'name': 'Natura2000', 'category': 'Rasterkaarten', 'type': 'wms', 'url': 'https://service.pdok.nl/rvo/natura2000/wms/v1_0', 'layer': 'natura2000', 'sheet_name': 'GIS - Natura2000'},
    'bag_panden_2d': {'name': 'BAG Panden 2D', 'category': '2D Vectordata', 'type': 'wfs', 'url': 'https://service.pdok.nl/lv/bag/wfs/v2_0', 'layer': 'bag:pand', 'sheet_name': 'GIS - BAG 2D'},
    'bgt_wegdelen': {'name': 'BGT Wegdelen', 'category': '2D Vectordata', 'type': 'ogcapi', 'url': 'https://api.pdok.nl/lv/bgt/ogc/v1', 'collection': 'wegdeel', 'sheet_name': 'GIS - BGT Wegdelen'},
    'bgt_waterdelen': {'name': 'BGT Waterdelen', 'category': '2D Vectordata', 'type': 'ogcapi', 'url': 'https://api.pdok.nl/lv/bgt/ogc/v1', 'collection': 'waterdeel', 'sheet_name': 'GIS - BGT Waterdelen'},
    'bgt_panden': {'name': 'BGT Panden', 'category': '2D Vectordata', 'type': 'ogcapi', 'url': 'https://api.pdok.nl/lv/bgt/ogc/v1', 'collection': 'pand', 'sheet_name': 'GIS - BGT Panden'},
    'kadaster_percelen': {'name': 'Kadaster Percelen', 'category': '2D Vectordata', 'type': 'wfs', 'url': 'https://service.pdok.nl/kadaster/kadastralekaart/wfs/v5_0', 'layer': 'kadastralekaart:Perceel', 'sheet_name': 'GIS - Percelen'},
    'bag_3d_cityjson': {'name': '3D BAG LOD2.2 (CityJSON)', 'category': '3D Data', 'type': '3dbag_cityjson', 'url': 'https://api.3dbag.nl', 'sheet_name': 'GIS - 3D BAG'},
}


# =============================================================================
# HELPER FUNCTIES
# =============================================================================

def meters_to_internal(meters):
    return meters * METERS_TO_FEET

def internal_to_meters(internal):
    return internal * FEET_TO_METERS

def web_request(url):
    client = WebClient()
    client.Headers.Add("User-Agent", "Mozilla/5.0 GIS2BIM/PyRevit")
    client.Encoding = Encoding.UTF8
    return client.DownloadString(url)

def download_image(url):
    request = WebRequest.Create(url)
    request.Accept = "image/png,image/*"
    request.UserAgent = "Mozilla/5.0 GIS2BIM/PyRevit"
    request.Timeout = 60000
    response = request.GetResponse()
    bitmap = Bitmap(response.GetResponseStream())
    response.Close()
    return bitmap


# =============================================================================
# REVIT FUNCTIES
# =============================================================================

def get_survey_point(doc):
    collector = FilteredElementCollector(doc)
    base_points = collector.OfClass(BasePoint).ToElements()
    for bp in base_points:
        if bp.IsShared:
            return bp
    return None

def get_project_location(doc):
    try:
        survey_point = get_survey_point(doc)
        if survey_point:
            pos = survey_point.SharedPosition
            return internal_to_meters(pos.X), internal_to_meters(pos.Y)
    except:
        pass
    return None, None

def set_survey_point(doc, rd_x, rd_y):
    survey_point = get_survey_point(doc)
    if survey_point is None:
        raise Exception("Survey Point niet gevonden")
    new_position = XYZ(meters_to_internal(rd_x), meters_to_internal(rd_y), 0)
    move_vector = new_position - survey_point.SharedPosition
    ElementTransformUtils.MoveElement(doc, survey_point.Id, move_vector)

def get_titleblock(doc):
    collector = FilteredElementCollector(doc)
    titleblocks = collector.OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    return titleblocks[0].Id if titleblocks else None

def get_view_family_type(doc, view_family):
    collector = FilteredElementCollector(doc)
    for t in collector.OfClass(ViewFamilyType).ToElements():
        if t.ViewFamily == view_family:
            return t.Id
    return None

def create_drafting_view(doc, name):
    vft_id = get_view_family_type(doc, ViewFamily.Drafting)
    if vft_id is None:
        raise Exception("Geen Drafting View type gevonden")
    view = ViewDrafting.Create(doc, vft_id)
    view.Name = name
    view.Scale = 500
    return view

def create_sheet(doc, sheet_name, sheet_number):
    titleblock_id = get_titleblock(doc)
    sheet = ViewSheet.Create(doc, titleblock_id if titleblock_id else ElementId.InvalidElementId)
    sheet.Name = sheet_name
    sheet.SheetNumber = sheet_number
    return sheet

def place_view_on_sheet(doc, sheet, view):
    try:
        if Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
            return Viewport.Create(doc, sheet.Id, view.Id, XYZ(1.4, 0.5, 0))
    except:
        pass
    return None

def get_or_create_sheet(doc, sheet_name, sheet_number):
    for s in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements():
        if s.SheetNumber == sheet_number:
            return s, False
    return create_sheet(doc, sheet_name, sheet_number), True

def import_image_to_view(doc, view, image_path, width_m):
    options = ImageTypeOptions(image_path, False, ImageTypeSource.Import)
    options.Resolution = 300
    image_type = ImageType.Create(doc, options)
    placement_options = ImagePlacementOptions(XYZ(0, 0, 0), BoxPlacement.Center)
    image_instance = ImageInstance.Create(doc, view, image_type.Id, placement_options)
    try:
        image_instance.Width = meters_to_internal(width_m)
    except:
        pass
    return image_instance


# =============================================================================
# GEOCODING
# =============================================================================

def geocode_address(address):
    encoded = address.replace(' ', '%20').replace(',', '%2C')
    url = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q={}&rows=5".format(encoded)
    data = json.loads(web_request(url))
    results = []
    if data.get('response', {}).get('numFound', 0) > 0:
        for d in data['response']['docs']:
            centroid = d.get('centroide_rd', '')
            if centroid:
                coords = centroid.replace('POINT(', '').replace(')', '').split()
                results.append((float(coords[0]), float(coords[1]), d.get('weergavenaam', address)))
    return results


# =============================================================================
# DATA DOWNLOAD FUNCTIES
# =============================================================================

def download_wmts_tiles(layer_config, rd_x, rd_y, bbox_size):
    zoomlevel = layer_config.get('zoom', 12)
    base_url = layer_config['url']
    layer_name = layer_config['layer']
    resolution = ZOOMLEVEL_RESOLUTIONS.get(zoomlevel, ZOOMLEVEL_RESOLUTIONS[12])
    tile_width_m = WMTS_PIXEL_WIDTH * resolution
    center_col = int((rd_x - WMTS_XCORNER) / tile_width_m)
    center_row = int((WMTS_YCORNER - rd_y) / tile_width_m)
    tiles_needed = int(math.ceil(bbox_size / tile_width_m))
    if tiles_needed % 2 == 1:
        tiles_needed += 1
    half = tiles_needed // 2
    cols = range(center_col - half, center_col + half + 1)
    rows = range(center_row - half, center_row + half + 1)
    tiles = []
    for row in rows:
        row_tiles = []
        for col in cols:
            url = "{base}?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER={layer}&STYLE=default&FORMAT=image/png&TILEMATRIXSET=EPSG:28992&TILEMATRIX={zoom}&TILEROW={row}&TILECOL={col}".format(base=base_url, layer=layer_name, zoom=zoomlevel, row=row, col=col)
            try:
                tile = download_image(url)
            except:
                tile = Bitmap(WMTS_PIXEL_WIDTH, WMTS_PIXEL_WIDTH)
            row_tiles.append(tile)
        tiles.append(row_tiles)
    total_w = len(cols) * WMTS_PIXEL_WIDTH
    total_h = len(rows) * WMTS_PIXEL_WIDTH
    combined = Bitmap(total_w, total_h)
    graphics = Graphics.FromImage(combined)
    for ri, row_tiles in enumerate(tiles):
        for ci, tile in enumerate(row_tiles):
            graphics.DrawImage(tile, ci * WMTS_PIXEL_WIDTH, ri * WMTS_PIXEL_WIDTH)
    graphics.Dispose()
    return combined, len(cols) * tile_width_m, len(rows) * tile_width_m

def download_wms_image(layer_config, rd_x, rd_y, bbox_size, image_size=2048):
    base_url = layer_config['url']
    layer_name = layer_config['layer']
    half = bbox_size / 2
    bbox = "{},{},{},{}".format(rd_x - half, rd_y - half, rd_x + half, rd_y + half)
    url = "{base}?SERVICE=WMS&VERSION=1.3.0&REQUEST=GetMap&LAYERS={layer}&STYLES=&CRS=EPSG:28992&BBOX={bbox}&WIDTH={size}&HEIGHT={size}&FORMAT=image/png&TRANSPARENT=true".format(base=base_url, layer=layer_name, bbox=bbox, size=image_size)
    return download_image(url), bbox_size, bbox_size

def get_wfs_features(layer_config, rd_x, rd_y, bbox_size, max_features=500):
    base_url = layer_config['url']
    layer_name = layer_config['layer']
    half = bbox_size / 2
    bbox = (rd_x - half, rd_y - half, rd_x + half, rd_y + half)
    url = "{base}?SERVICE=WFS&VERSION=2.0.0&REQUEST=GetFeature&TYPENAMES={layer}&BBOX={minx},{miny},{maxx},{maxy},urn:ogc:def:crs:EPSG::28992&COUNT={count}&OUTPUTFORMAT=application/json".format(base=base_url, layer=layer_name, minx=bbox[0], miny=bbox[1], maxx=bbox[2], maxy=bbox[3], count=max_features)
    return json.loads(web_request(url)).get('features', [])

def get_ogcapi_features(layer_config, rd_x, rd_y, bbox_size, max_features=500):
    base_url = layer_config['url']
    collection = layer_config['collection']
    half = bbox_size / 2
    url = "{base}/collections/{collection}/items?bbox={minx},{miny},{maxx},{maxy}&bbox-crs=http://www.opengis.net/def/crs/EPSG/0/28992&crs=http://www.opengis.net/def/crs/EPSG/0/28992&limit={limit}&f=json".format(base=base_url, collection=collection, minx=rd_x-half, miny=rd_y-half, maxx=rd_x+half, maxy=rd_y+half, limit=max_features)
    return json.loads(web_request(url)).get('features', [])

def extract_polygon_rings(geometry):
    geom_type = geometry.get('type', '')
    coords = geometry.get('coordinates', [])
    rings = []
    if geom_type == 'Polygon' and coords and coords[0]:
        rings.append([(c[0], c[1]) for c in coords[0]])
    elif geom_type == 'MultiPolygon' and coords:
        for poly in coords:
            if poly and poly[0]:
                rings.append([(c[0], c[1]) for c in poly[0]])
    return rings

def create_detail_lines_in_view(doc, view, features, rd_x, rd_y):
    lines_created = 0
    offset_value = 0.0001
    for feat_idx, feature in enumerate(features):
        rings = extract_polygon_rings(feature.get('geometry', {}))
        for ring_idx, ring in enumerate(rings):
            if len(ring) < 3:
                continue
            poly_offset = (feat_idx * 100 + ring_idx) * offset_value
            for i in range(len(ring)):
                x1, y1 = ring[i]
                x2, y2 = ring[(i + 1) % len(ring)]
                p1 = XYZ(meters_to_internal(x1 - rd_x) + poly_offset, meters_to_internal(y1 - rd_y) + poly_offset, 0)
                p2 = XYZ(meters_to_internal(x2 - rd_x) + poly_offset, meters_to_internal(y2 - rd_y) + poly_offset, 0)
                if p1.DistanceTo(p2) > 0.01:
                    try:
                        doc.Create.NewDetailCurve(view, Line.CreateBound(p1, p2))
                        lines_created += 1
                    except:
                        pass
    return lines_created


# =============================================================================
# 3D BAG CITYJSON
# =============================================================================

def get_3dbag_cityjson_bbox(rd_x, rd_y, bbox_size):
    half = bbox_size / 2.0
    api_url = "https://api.3dbag.nl/collections/pand/items?bbox={},{},{},{}".format(rd_x-half, rd_y-half, rd_x+half, rd_y+half)
    print("3D BAG API URL: {}".format(api_url))
    try:
        data = json.loads(web_request(api_url))
        metadata = data.get('metadata', {})
        features = data.get('features', [])
        if 'feature' in data and not features:
            features = [data['feature']]
        if not features:
            return None, "Geen gebouwen gevonden"
        print("Gevonden: {} gebouwen".format(len(features)))
        return {'metadata': metadata, 'features': features}, None
    except Exception as e:
        return None, str(e)

def parse_cityjson_lod22_geometry(cityjson_data, rd_x, rd_y):
    metadata = cityjson_data.get('metadata', {})
    features = cityjson_data.get('features', [])
    transform = metadata.get('transform', {})
    scale_raw = transform.get('scale', [1.0, 1.0, 1.0])
    translate_raw = transform.get('translate', [0.0, 0.0, 0.0])
    scale = [float(x) for x in scale_raw.split()] if isinstance(scale_raw, str) else list(scale_raw) if scale_raw else [1.0, 1.0, 1.0]
    translate = [float(x) for x in translate_raw.split()] if isinstance(translate_raw, str) else list(translate_raw) if translate_raw else [0.0, 0.0, 0.0]
    print("Transform - scale: {}, translate: {}".format(scale, translate))
    buildings = []
    for feature in features:
        try:
            city_objects = feature.get('CityObjects', {})
            vertices = feature.get('vertices', [])
            if not vertices:
                continue
            converted_verts = [(v[0]*scale[0]+translate[0]-rd_x, v[1]*scale[1]+translate[1]-rd_y, v[2]*scale[2]+translate[2]) for v in vertices]
            for obj_id, obj in city_objects.items():
                if obj.get('type', '') not in ['BuildingPart', 'Building']:
                    continue
                for geom in obj.get('geometry', []):
                    if str(geom.get('lod', '')) != '2.2':
                        continue
                    geom_type = geom.get('type', '')
                    boundaries = geom.get('boundaries', [])
                    if not boundaries:
                        continue
                    polygon_faces = []
                    if geom_type == 'Solid':
                        for shell in boundaries:
                            for surface in shell:
                                if surface and len(surface) > 0:
                                    outer_ring = surface[0] if isinstance(surface[0], list) else surface
                                    if len(outer_ring) >= 3:
                                        polygon_faces.append(list(outer_ring))
                    elif geom_type == 'MultiSurface':
                        for surface in boundaries:
                            if surface and len(surface) > 0:
                                outer_ring = surface[0] if isinstance(surface[0], list) else surface
                                if len(outer_ring) >= 3:
                                    polygon_faces.append(list(outer_ring))
                    if polygon_faces:
                        buildings.append({'id': obj_id, 'vertices': converted_verts, 'polygon_faces': polygon_faces})
                        print("Building {} - {} faces".format(obj_id, len(polygon_faces)))
        except Exception as e:
            print("Parse error: {}".format(str(e)))
    return buildings

def create_directshapes_from_buildings(doc, buildings):
    if not buildings:
        return 0, "Geen gebouwen"
    count = 0
    last_error = ""
    for building in buildings:
        try:
            vertices = building['vertices']
            polygon_faces = building.get('polygon_faces', [])
            obj_id = building.get('id', 'unknown')
            if not polygon_faces or not vertices:
                continue
            xyz_verts = [XYZ(meters_to_internal(v[0]), meters_to_internal(v[1]), meters_to_internal(v[2])) for v in vertices]
            builder = TessellatedShapeBuilder()
            builder.OpenConnectedFaceSet(False)
            faces_added = 0
            for poly_indices in polygon_faces:
                if len(poly_indices) < 3:
                    continue
                if any(idx >= len(xyz_verts) for idx in poly_indices):
                    continue
                face_verts = List[XYZ]()
                prev_pt = None
                for idx in poly_indices:
                    pt = xyz_verts[idx]
                    if prev_pt is None or pt.DistanceTo(prev_pt) >= 0.0001:
                        face_verts.Add(pt)
                        prev_pt = pt
                if face_verts.Count >= 3:
                    if face_verts[0].DistanceTo(face_verts[face_verts.Count - 1]) < 0.0001:
                        new_verts = List[XYZ]()
                        for i in range(face_verts.Count - 1):
                            new_verts.Add(face_verts[i])
                        face_verts = new_verts
                    if face_verts.Count >= 3:
                        try:
                            builder.AddFace(TessellatedFace(face_verts, ElementId.InvalidElementId))
                            faces_added += 1
                        except:
                            pass
            print("  Faces: {}".format(faces_added))
            if faces_added == 0:
                continue
            builder.CloseConnectedFaceSet()
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry
            builder.Fallback = TessellatedShapeBuilderFallback.Salvage
            builder.Build()
            result = builder.GetBuildResult()
            print("  Outcome: {}".format(result.Outcome))
            if result.Outcome != TessellatedShapeBuilderOutcome.Nothing:
                geom_objects = result.GetGeometricalObjects()
                if geom_objects.Count > 0:
                    ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))
                    ds.SetShape(geom_objects)
                    try:
                        ds.Name = "3DBAG_{}".format(obj_id.replace('NL.IMBAG.Pand.', ''))
                    except:
                        pass
                    count += 1
                    print("DirectShape: {}".format(obj_id))
            else:
                last_error = "TessellatedShapeBuilder returned Nothing"
        except Exception as e:
            last_error = str(e)
            print("Error: {}".format(str(e)))
    return (count, None) if count > 0 else (0, "Geen DirectShapes. Fout: {}".format(last_error))

def import_3dbag_cityjson(doc, rd_x, rd_y, bbox_size):
    print("=" * 50)
    print("3D BAG CityJSON Import - RD {}, {} - {}m".format(int(rd_x), int(rd_y), bbox_size))
    cityjson_data, error = get_3dbag_cityjson_bbox(rd_x, rd_y, bbox_size)
    if error:
        return False, "Download mislukt: {}".format(error)
    buildings = parse_cityjson_lod22_geometry(cityjson_data, rd_x, rd_y)
    if not buildings:
        return False, "Geen LOD 2.2 geometrie gevonden"
    print("Parsed: {} gebouwen".format(len(buildings)))
    count, error = create_directshapes_from_buildings(doc, buildings)
    if error:
        return False, error
    return True, "{} gebouwen geimporteerd".format(count)


# =============================================================================
# MAIN DIALOG - Using composition pattern for IronPython compatibility
# =============================================================================

def create_gis2bim_dialog(doc):
    """Create and show the GIS2BIM dialog."""
    # Initialize UI resources
    init_ui()

    # Create state container
    state = {
        'doc': doc,
        'rd_x': DEFAULT_RD_X,
        'rd_y': DEFAULT_RD_Y,
        'address_results': [],
        'layer_bbox_sizes': {key: 500 for key in GIS_LAYERS},
        'map_image': None,
        'map_bbox_size': 500,
        'map_width_px': 0,
        'map_height_px': 0,
    }

    # Get project location
    proj_x, proj_y = get_project_location(doc)
    if proj_x and proj_y:
        state['rd_x'] = proj_x
        state['rd_y'] = proj_y

    # Create main form using WinForms.Form directly (use full path to avoid IronPython issues)
    form = WinForms.Form()
    form.Text = "GIS2BIM - Nederlandse Geodata v6.4"
    set_size(form, DPIScaler.scale(1100), DPIScaler.scale(750))
    form.StartPosition = FormStartPosition.CenterScreen
    form.BackColor = COLOR_BACKGROUND
    form.Font = FONT_NORMAL
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False

    # Store controls in state
    controls = {}

    scale = DPIScaler.scale
    margin = scale(15)
    map_width = scale(450)
    right_panel_x = map_width + margin * 2
    right_panel_width = form.ClientSize.Width - right_panel_x - margin

    # === MAP PANEL ===
    x, y = margin, margin
    width = map_width
    height = form.ClientSize.Height - margin * 2

    lbl_map = UIFactory.create_label("Locatie Kaart", bold=True, color=COLOR_PRIMARY)
    set_location(lbl_map, x, y)
    form.Controls.Add(lbl_map)

    map_top = y + scale(25)
    map_height = height - scale(160)

    map_picture = PictureBox()
    set_bounds(map_picture, x, map_top, width, map_height)
    map_picture.BorderStyle = BorderStyle.FixedSingle
    map_picture.SizeMode = PictureBoxSizeMode.Zoom
    map_picture.Cursor = Cursors.Cross
    map_picture.BackColor = DrawingColor.FromArgb(240, 240, 240)
    form.Controls.Add(map_picture)
    controls['map_picture'] = map_picture

    state['map_width_px'] = width
    state['map_height_px'] = map_height

    zoom_y = map_top + map_height + scale(5)
    zoom_slider = WinForms.TrackBar()
    set_bounds(zoom_slider, x + scale(20), zoom_y, width - scale(130), scale(20))
    zoom_slider.AutoSize = False
    zoom_slider.Minimum = 500
    zoom_slider.Maximum = 15000
    zoom_slider.Value = 500
    zoom_slider.TickStyle = WinForms.TickStyle.None
    form.Controls.Add(zoom_slider)
    controls['zoom_slider'] = zoom_slider

    lbl_zoom_value = UIFactory.create_label("500m")
    set_location(lbl_zoom_value, x + width - scale(80), zoom_y)
    form.Controls.Add(lbl_zoom_value)
    controls['lbl_zoom_value'] = lbl_zoom_value

    coord_y = zoom_y + scale(25)
    lbl_rdx = UIFactory.create_label("RD X:")
    set_location(lbl_rdx, x, coord_y + scale(5))
    form.Controls.Add(lbl_rdx)

    txt_rd_x = UIFactory.create_textbox()
    set_bounds(txt_rd_x, x + scale(45), coord_y, scale(90), scale(25))
    txt_rd_x.Text = str(int(state['rd_x']))
    form.Controls.Add(txt_rd_x)
    controls['txt_rd_x'] = txt_rd_x

    lbl_rdy = UIFactory.create_label("RD Y:")
    set_location(lbl_rdy, x + scale(145), coord_y + scale(5))
    form.Controls.Add(lbl_rdy)

    txt_rd_y = UIFactory.create_textbox()
    set_bounds(txt_rd_y, x + scale(190), coord_y, scale(90), scale(25))
    txt_rd_y.Text = str(int(state['rd_y']))
    form.Controls.Add(txt_rd_y)
    controls['txt_rd_y'] = txt_rd_y

    btn_goto = UIFactory.create_button("Ga naar", primary=True, width=80)
    set_location(btn_goto, x + scale(290), coord_y)
    form.Controls.Add(btn_goto)
    controls['btn_goto'] = btn_goto

    btn_refresh = UIFactory.create_button("Ververs", primary=False, width=70)
    set_location(btn_refresh, x + scale(380), coord_y)
    form.Controls.Add(btn_refresh)
    controls['btn_refresh'] = btn_refresh

    # === CONTROLS PANEL ===
    x = right_panel_x
    y = margin
    width = right_panel_width

    section1 = UIFactory.create_label("1. Zoek Adres", bold=True, color=COLOR_PRIMARY)
    set_location(section1, x, y)
    form.Controls.Add(section1)
    y += scale(25)

    lbl_location = UIFactory.create_label("Locatie: RD X {} | RD Y {}".format(int(state['rd_x']), int(state['rd_y'])))
    set_location(lbl_location, x, y)
    lbl_location.ForeColor = COLOR_PRIMARY_DARK
    form.Controls.Add(lbl_location)
    controls['lbl_location'] = lbl_location
    y += scale(25)

    txt_address = UIFactory.create_textbox()
    set_bounds(txt_address, x, y, width - scale(90), scale(25))
    txt_address.Text = DEFAULT_ADDRESS
    form.Controls.Add(txt_address)
    controls['txt_address'] = txt_address

    btn_search = UIFactory.create_button("Zoeken", primary=True, width=80)
    set_location(btn_search, x + width - scale(80), y)
    form.Controls.Add(btn_search)
    controls['btn_search'] = btn_search
    y += scale(35)

    lst_address = UIFactory.create_listview([("Adres", 280), ("RD X", 70), ("RD Y", 70)])
    set_bounds(lst_address, x, y, width, scale(80))
    form.Controls.Add(lst_address)
    controls['lst_address'] = lst_address
    y += scale(85)

    btn_apply_loc = UIFactory.create_button("Locatie Toepassen", primary=False, width=150)
    set_location(btn_apply_loc, x, y)
    form.Controls.Add(btn_apply_loc)
    controls['btn_apply_loc'] = btn_apply_loc
    y += scale(40)

    sep1 = UIFactory.create_separator(width=width)
    set_location(sep1, x, y)
    form.Controls.Add(sep1)
    y += scale(15)

    section2 = UIFactory.create_label("2. Selecteer Data", bold=True, color=COLOR_PRIMARY)
    set_location(section2, x, y)
    form.Controls.Add(section2)
    y += scale(25)

    lst_layers = UIFactory.create_listview([("Laag", 160), ("Type", 60), ("Omtrek", 55), ("Status", 55)])
    set_bounds(lst_layers, x, y, width, scale(150))
    lst_layers.CheckBoxes = True
    form.Controls.Add(lst_layers)
    controls['lst_layers'] = lst_layers

    # Populate layers
    for key, layer in GIS_LAYERS.items():
        item = ListViewItem(layer['name'])
        item.SubItems.Add({'wmts': 'Kaart', 'wms': 'Kaart', 'wfs': 'Lijnen', 'ogcapi': 'Lijnen', '3dbag_cityjson': 'LOD2.2'}.get(layer['type'], layer['type']))
        item.SubItems.Add("500m")
        item.SubItems.Add("Gereed")
        item.Tag = key
        lst_layers.Items.Add(item)

    y += scale(155)

    lbl_size = UIFactory.create_label("Omtrek:")
    set_location(lbl_size, x, y + scale(3))
    form.Controls.Add(lbl_size)

    num_layer_bbox = UIFactory.create_numeric(min_val=50, max_val=2000, default=500)
    set_location(num_layer_bbox, x + scale(60), y)
    form.Controls.Add(num_layer_bbox)
    controls['num_layer_bbox'] = num_layer_bbox
    y += scale(30)

    btn_x = x
    btn_all = UIFactory.create_button("Alles", primary=False, width=60)
    set_location(btn_all, btn_x, y)
    form.Controls.Add(btn_all)
    controls['btn_all'] = btn_all
    btn_x += scale(65)

    btn_none = UIFactory.create_button("Geen", primary=False, width=60)
    set_location(btn_none, btn_x, y)
    form.Controls.Add(btn_none)
    controls['btn_none'] = btn_none
    btn_x += scale(65)

    btn_raster = UIFactory.create_button("Kaarten", primary=False, width=60)
    set_location(btn_raster, btn_x, y)
    form.Controls.Add(btn_raster)
    controls['btn_raster'] = btn_raster
    btn_x += scale(65)

    btn_3d = UIFactory.create_button("3D", primary=False, width=60)
    set_location(btn_3d, btn_x, y)
    form.Controls.Add(btn_3d)
    controls['btn_3d'] = btn_3d
    y += scale(40)

    sep2 = UIFactory.create_separator(width=width)
    set_location(sep2, x, y)
    form.Controls.Add(sep2)
    y += scale(15)

    section3 = UIFactory.create_label("3. Uitvoeren", bold=True, color=COLOR_PRIMARY)
    set_location(section3, x, y)
    form.Controls.Add(section3)
    y += scale(25)

    lbl_status = UIFactory.create_label("Gereed")
    set_bounds(lbl_status, x, y, width, scale(20))
    lbl_status.ForeColor = COLOR_TEXT_SECONDARY
    form.Controls.Add(lbl_status)
    controls['lbl_status'] = lbl_status
    y += scale(25)

    progress = UIFactory.create_progressbar(width=width)
    set_location(progress, x, y)
    form.Controls.Add(progress)
    controls['progress'] = progress
    y += scale(35)

    btn_execute = UIFactory.create_button("Start Import", primary=True, width=120)
    set_location(btn_execute, x, y)
    form.Controls.Add(btn_execute)
    controls['btn_execute'] = btn_execute

    btn_close = UIFactory.create_button("Sluiten", primary=False, width=100)
    set_location(btn_close, x + scale(130), y)
    form.Controls.Add(btn_close)
    controls['btn_close'] = btn_close

    # === EVENT HANDLERS ===
    def load_map():
        try:
            half = state['map_bbox_size'] / 2.0
            # Use luchtfoto RGB service (publicly available aerial imagery)
            url = "https://service.pdok.nl/hwh/luchtfotorgb/wms/v1_0?SERVICE=WMS&VERSION=1.3.0&REQUEST=GetMap&LAYERS=Actueel_orthoHR&STYLES=&CRS=EPSG:28992&BBOX={},{},{},{}&WIDTH={}&HEIGHT={}&FORMAT=image/png".format(
                state['rd_x']-half, state['rd_y']-half, state['rd_x']+half, state['rd_y']+half,
                int(state['map_width_px']), int(state['map_height_px']))
            state['map_image'] = download_image(url)
            draw_marker()
        except Exception as e:
            print("Map error: {}".format(str(e)))
            controls['map_picture'].Image = None

    def draw_marker():
        if state['map_image'] is None:
            return
        img = Bitmap(state['map_image'].Width, state['map_image'].Height)
        g = Graphics.FromImage(img)
        g.DrawImage(state['map_image'], 0, 0)
        cx, cy = state['map_image'].Width // 2, state['map_image'].Height // 2
        pen_red = Pen(DrawingColor.Red, 3)
        g.DrawLine(pen_red, cx - 20, cy, cx + 20, cy)
        g.DrawLine(pen_red, cx, cy - 20, cx, cy + 20)
        pen_red.Dispose()
        g.Dispose()
        controls['map_picture'].Image = img

    def on_zoom_changed(sender, args):
        state['map_bbox_size'] = controls['zoom_slider'].Value
        controls['lbl_zoom_value'].Text = "{}m".format(state['map_bbox_size'])
        load_map()

    def on_map_click(sender, args):
        if state['map_image'] is None:
            return
        pb = controls['map_picture']
        img = state['map_image']
        sc = max(float(img.Width) / pb.Width, float(img.Height) / pb.Height)
        offset_x = (pb.Width - img.Width / sc) / 2.0
        offset_y = (pb.Height - img.Height / sc) / 2.0
        img_x = (args.X - offset_x) * sc
        img_y = (args.Y - offset_y) * sc
        half = state['map_bbox_size'] / 2.0
        state['rd_x'] = (state['rd_x'] - half) + (img_x / float(img.Width)) * state['map_bbox_size']
        state['rd_y'] = (state['rd_y'] + half) - (img_y / float(img.Height)) * state['map_bbox_size']
        controls['txt_rd_x'].Text = str(int(state['rd_x']))
        controls['txt_rd_y'].Text = str(int(state['rd_y']))
        controls['lbl_location'].Text = "Locatie: RD X {} | RD Y {}".format(int(state['rd_x']), int(state['rd_y']))
        load_map()

    def on_goto_coords(sender, args):
        try:
            state['rd_x'] = float(controls['txt_rd_x'].Text.replace(',', '.'))
            state['rd_y'] = float(controls['txt_rd_y'].Text.replace(',', '.'))
            controls['lbl_location'].Text = "Locatie: RD X {} | RD Y {}".format(int(state['rd_x']), int(state['rd_y']))
            load_map()
        except:
            show_warning("Ongeldige coordinaten")

    def on_search(sender, args):
        address = controls['txt_address'].Text.strip()
        if not address:
            return
        try:
            controls['lbl_status'].Text = "Zoeken..."
            Application.DoEvents()
            state['address_results'] = geocode_address(address)
            controls['lst_address'].Items.Clear()
            for rx, ry, name in state['address_results']:
                item = ListViewItem(name)
                item.SubItems.Add(str(int(rx)))
                item.SubItems.Add(str(int(ry)))
                controls['lst_address'].Items.Add(item)
            controls['lbl_status'].Text = "{} resultaten".format(len(state['address_results']))
        except Exception as e:
            show_error(str(e))

    def on_address_select(sender, args):
        if controls['lst_address'].SelectedIndices.Count > 0:
            idx = controls['lst_address'].SelectedIndices[0]
            if idx < len(state['address_results']):
                state['rd_x'], state['rd_y'], _ = state['address_results'][idx]
                controls['txt_rd_x'].Text = str(int(state['rd_x']))
                controls['txt_rd_y'].Text = str(int(state['rd_y']))
                controls['lbl_location'].Text = "Locatie: RD X {} | RD Y {}".format(int(state['rd_x']), int(state['rd_y']))
                load_map()

    def on_apply_location(sender, args):
        try:
            with Transaction(state['doc'], "GIS2BIM - Set Location") as t:
                t.Start()
                set_survey_point(state['doc'], state['rd_x'], state['rd_y'])
                t.Commit()
            show_info("Locatie ingesteld!")
        except Exception as e:
            show_error(str(e))

    def on_layer_select(sender, args):
        if controls['lst_layers'].SelectedIndices.Count > 0:
            key = controls['lst_layers'].Items[controls['lst_layers'].SelectedIndices[0]].Tag
            controls['num_layer_bbox'].Value = state['layer_bbox_sizes'].get(key, 500)

    def on_layer_bbox_changed(sender, args):
        if controls['lst_layers'].SelectedIndices.Count > 0:
            item = controls['lst_layers'].Items[controls['lst_layers'].SelectedIndices[0]]
            state['layer_bbox_sizes'][item.Tag] = int(controls['num_layer_bbox'].Value)
            item.SubItems[2].Text = "{}m".format(int(controls['num_layer_bbox'].Value))

    def select_all(sender, args):
        for item in controls['lst_layers'].Items:
            item.Checked = True

    def select_none(sender, args):
        for item in controls['lst_layers'].Items:
            item.Checked = False

    def select_raster(sender, args):
        for item in controls['lst_layers'].Items:
            item.Checked = GIS_LAYERS.get(item.Tag, {}).get('type') in ['wmts', 'wms']

    def select_3d(sender, args):
        for item in controls['lst_layers'].Items:
            item.Checked = GIS_LAYERS.get(item.Tag, {}).get('type') == '3dbag_cityjson'

    def import_layer(key, layer, sheet_num, bbox, output_folder):
        if layer['type'] == '3dbag_cityjson':
            success, msg = import_3dbag_cityjson(state['doc'], state['rd_x'], state['rd_y'], bbox)
            if not success:
                raise Exception(msg)
            return msg
        sheet, _ = get_or_create_sheet(state['doc'], layer['sheet_name'], sheet_num)
        view = None
        for v in FilteredElementCollector(state['doc']).OfClass(ViewDrafting).ToElements():
            if v.Name == "GIS_{}".format(key):
                view = v
                break
        if view is None:
            view = create_drafting_view(state['doc'], "GIS_{}".format(key))
        if layer['type'] == 'wmts':
            bitmap, width_m, _ = download_wmts_tiles(layer, state['rd_x'], state['rd_y'], bbox)
            path = os.path.join(output_folder, "GIS2BIM_{}.png".format(key))
            bitmap.Save(path, ImageFormat.Png)
            import_image_to_view(state['doc'], view, path, width_m)
            place_view_on_sheet(state['doc'], sheet, view)
            return "OK"
        elif layer['type'] == 'wms':
            bitmap, width_m, _ = download_wms_image(layer, state['rd_x'], state['rd_y'], bbox)
            path = os.path.join(output_folder, "GIS2BIM_{}.png".format(key))
            bitmap.Save(path, ImageFormat.Png)
            import_image_to_view(state['doc'], view, path, width_m)
            place_view_on_sheet(state['doc'], sheet, view)
            return "OK"
        elif layer['type'] == 'wfs':
            features = get_wfs_features(layer, state['rd_x'], state['rd_y'], bbox)
            lines = create_detail_lines_in_view(state['doc'], view, features, state['rd_x'], state['rd_y']) if features else 0
            place_view_on_sheet(state['doc'], sheet, view)
            return "{} lijnen".format(lines)
        elif layer['type'] == 'ogcapi':
            features = get_ogcapi_features(layer, state['rd_x'], state['rd_y'], bbox)
            lines = create_detail_lines_in_view(state['doc'], view, features, state['rd_x'], state['rd_y']) if features else 0
            place_view_on_sheet(state['doc'], sheet, view)
            return "{} lijnen".format(lines)
        return "Onbekend"

    def on_execute(sender, args):
        selected = [item.Tag for item in controls['lst_layers'].Items if item.Checked]
        if not selected:
            show_warning("Selecteer minimaal één laag")
            return
        if not show_question("Importeren van {} lagen?".format(len(selected))):
            return
        output_folder = os.path.join(os.path.dirname(state['doc'].PathName), "GIS2BIM") if state['doc'].PathName else tempfile.gettempdir()
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        results = []
        try:
            with Transaction(state['doc'], "GIS2BIM - Import") as t:
                t.Start()
                for i, key in enumerate(selected):
                    layer = GIS_LAYERS[key]
                    bbox = state['layer_bbox_sizes'].get(key, 500)
                    controls['lbl_status'].Text = "Importeren: {}...".format(layer['name'])
                    controls['progress'].Value = int((i / float(len(selected))) * 100)
                    Application.DoEvents()
                    try:
                        result = import_layer(key, layer, "GIS-{:02d}".format(i+1), bbox, output_folder)
                        results.append("{}: {}".format(layer['name'], result))
                    except Exception as e:
                        results.append("{}: FOUT - {}".format(layer['name'], str(e)))
                t.Commit()
            controls['progress'].Value = 100
            controls['lbl_status'].Text = "Voltooid!"
            show_info("Import voltooid!\n\n" + "\n".join(["- " + r for r in results]))
        except Exception as e:
            show_error(str(e))

    def on_close(sender, args):
        form.Close()

    # Connect event handlers
    controls['zoom_slider'].ValueChanged += on_zoom_changed
    controls['map_picture'].MouseClick += on_map_click
    controls['btn_goto'].Click += on_goto_coords
    controls['btn_refresh'].Click += lambda s, e: load_map()
    controls['btn_search'].Click += on_search
    controls['lst_address'].SelectedIndexChanged += on_address_select
    controls['btn_apply_loc'].Click += on_apply_location
    controls['lst_layers'].SelectedIndexChanged += on_layer_select
    controls['num_layer_bbox'].ValueChanged += on_layer_bbox_changed
    controls['btn_all'].Click += select_all
    controls['btn_none'].Click += select_none
    controls['btn_raster'].Click += select_raster
    controls['btn_3d'].Click += select_3d
    controls['btn_execute'].Click += on_execute
    controls['btn_close'].Click += on_close

    # Initial map load
    load_map()

    return form


def main():
    dialog = create_gis2bim_dialog(revit.doc)
    dialog.ShowDialog()


if __name__ == "__main__":
    main()
