# -*- coding: utf-8 -*-
"""
GIS2BIM - Nederlandse Geodata voor Revit
=========================================
Complete workflow met ALLE import types uit de originele Dynamo scripts.
Nu met moderne JMK UI Template styling en interactieve Leaflet kaart.

Auteur: 3BM / JMK
Versie: 6.4 - CityJSON DirectShape methode (geen IFC fallback)
"""

__title__ = "GIS2BIM"
__author__ = "3BM / JMK"
__doc__ = "Nederlandse GIS data workflow - Locatie, Selectie, Import"

# Imports
import clr
import os
import math
import json
import tempfile
import time
import zipfile
import gzip
import shutil

clr.AddReference('System')
clr.AddReference('System.Net')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Net import WebClient, WebRequest
from System.Text import Encoding
from System.Drawing import Bitmap, Graphics, Color as DrawingColor, Pen, SolidBrush, Rectangle, StringFormat, ContentAlignment
from System.Drawing.Drawing2D import SmoothingMode, DashStyle
from System.Drawing.Imaging import ImageFormat
from System.Collections.Generic import List

import System.Windows.Forms as WinForms
from System.Windows.Forms import (
    Application, DialogResult, ListViewItem, PictureBox, 
    PictureBoxSizeMode, Panel, BorderStyle, Cursors, MouseButtons
)

from pyrevit import revit

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import GeometryObject

from Autodesk.Revit.UI import TaskDialog

# UI Template imports
from ui_template import (
    BaseForm, UIFactory, COLORS, FONTS, DPIScaler,
    run_dialog
)

# Import Point en Size vanuit System.Drawing
from System.Drawing import Point, Size


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

# 3D BAG versie (Sept 2025)
DBAG_VERSION = "20250903"

# Default locatie (Dordrecht - Burgemeester de Raadtsingel 31)
DEFAULT_RD_X = 99628
DEFAULT_RD_Y = 424889
DEFAULT_ADDRESS = "Burgemeester de Raadtsingel 31, Dordrecht"


# =============================================================================
# GIS DATA LAGEN (Updated januari 2025)
# =============================================================================

GIS_LAYERS = {
    # === RASTERKAARTEN (WMTS/WMS) ===
    'luchtfoto_actueel': {
        'name': 'Luchtfoto Actueel',
        'category': 'Rasterkaarten',
        'type': 'wmts',
        'url': 'https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0',
        'layer': 'Actueel_orthoHR',
        'zoom': 14,
        'sheet_name': 'GIS - Luchtfoto Actueel'
    },
    'luchtfoto_2023': {
        'name': 'Luchtfoto 2023',
        'category': 'Rasterkaarten',
        'type': 'wmts',
        'url': 'https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0',
        'layer': '2023_orthoHR',
        'zoom': 14,
        'sheet_name': 'GIS - Luchtfoto 2023'
    },
    'top10nl': {
        'name': 'Top10NL Achtergrond',
        'category': 'Rasterkaarten',
        'type': 'wmts',
        'url': 'https://service.pdok.nl/brt/achtergrondkaart/wmts/v2_0',
        'layer': 'standaard',
        'zoom': 11,
        'sheet_name': 'GIS - Top10NL'
    },
    'bestemmingsplan': {
        'name': 'Bestemmingsplan',
        'category': 'Rasterkaarten',
        'type': 'wms',
        'url': 'https://service.pdok.nl/kadaster/plu/wms/v1_0',
        'layer': 'Bestemmingsplangebied',
        'sheet_name': 'GIS - Bestemmingsplan'
    },
    'bouwvlak_wms': {
        'name': 'Bouwvlak (kaart)',
        'category': 'Rasterkaarten',
        'type': 'wms',
        'url': 'https://service.pdok.nl/kadaster/plu/wms/v1_0',
        'layer': 'Bouwvlak',
        'sheet_name': 'GIS - Bouwvlak'
    },
    'bgt_achtergrond': {
        'name': 'BGT Achtergrondkaart',
        'category': 'Rasterkaarten',
        'type': 'wmts',
        'url': 'https://service.pdok.nl/lv/bgt/wmts/v1_0',
        'layer': 'standaardvisualisatie',
        'zoom': 14,
        'sheet_name': 'GIS - BGT'
    },
    'kadaster_kaart': {
        'name': 'Kadastrale Kaart',
        'category': 'Rasterkaarten',
        'type': 'wms',
        'url': 'https://service.pdok.nl/kadaster/kadastralekaart/wms/v5_0',
        'layer': 'Kadastralekaart',
        'sheet_name': 'GIS - Kadaster Kaart'
    },
    'natura2000': {
        'name': 'Natura2000',
        'category': 'Rasterkaarten',
        'type': 'wms',
        'url': 'https://service.pdok.nl/rvo/natura2000/wms/v1_0',
        'layer': 'natura2000',
        'sheet_name': 'GIS - Natura2000'
    },
    
    # === 2D VECTORDATA ===
    'bag_panden_2d': {
        'name': 'BAG Panden 2D',
        'category': '2D Vectordata',
        'type': 'wfs',
        'url': 'https://service.pdok.nl/lv/bag/wfs/v2_0',
        'layer': 'bag:pand',
        'sheet_name': 'GIS - BAG 2D'
    },
    'bgt_wegdelen': {
        'name': 'BGT Wegdelen',
        'category': '2D Vectordata',
        'type': 'ogcapi',
        'url': 'https://api.pdok.nl/lv/bgt/ogc/v1',
        'collection': 'wegdeel',
        'sheet_name': 'GIS - BGT Wegdelen'
    },
    'bgt_waterdelen': {
        'name': 'BGT Waterdelen',
        'category': '2D Vectordata',
        'type': 'ogcapi',
        'url': 'https://api.pdok.nl/lv/bgt/ogc/v1',
        'collection': 'waterdeel',
        'sheet_name': 'GIS - BGT Waterdelen'
    },
    'bgt_panden': {
        'name': 'BGT Panden',
        'category': '2D Vectordata',
        'type': 'ogcapi',
        'url': 'https://api.pdok.nl/lv/bgt/ogc/v1',
        'collection': 'pand',
        'sheet_name': 'GIS - BGT Panden'
    },
    'kadaster_percelen': {
        'name': 'Kadaster Percelen',
        'category': '2D Vectordata',
        'type': 'wfs',
        'url': 'https://service.pdok.nl/kadaster/kadastralekaart/wfs/v5_0',
        'layer': 'kadastralekaart:Perceel',
        'sheet_name': 'GIS - Percelen'
    },
    
    # === 3D DATA ===
    'bag_3d_cityjson': {
        'name': '3D BAG LOD2.2 (CityJSON)',
        'category': '3D Data',
        'type': '3dbag_cityjson',
        'url': 'https://api.3dbag.nl',
        'sheet_name': 'GIS - 3D BAG'
    },
}


# =============================================================================
# COORDINATE TRANSFORMATIE (RD <-> WGS84)
# =============================================================================

def rd_to_wgs84(rd_x, rd_y):
    """Converteer RD (EPSG:28992) naar WGS84 (lat/lon)"""
    x0 = 155000.0
    y0 = 463000.0
    phi0 = 52.15517440
    lam0 = 5.38720621
    
    dx = (rd_x - x0) * 1e-5
    dy = (rd_y - y0) * 1e-5
    
    phi = phi0 + (
        dy * 3235.65389 +
        dx * dx * -32.58297 +
        dy * dy * -0.24750 +
        dx * dx * dy * -0.84978 +
        dy * dy * dy * -0.06550 +
        dx * dx * dy * dy * -0.01709 +
        dx * dy * -0.00738 +
        dx * dx * dx * dx * 0.00530 +
        dx * dx * dx * dx * dy * -0.00039 +
        dx * dx * dx * dx * dy * dy * 0.00033 +
        dy * dy * dy * dy * -0.00012
    ) / 3600.0
    
    lam = lam0 + (
        dx * 5260.52916 +
        dx * dy * 105.94684 +
        dx * dy * dy * 2.45656 +
        dx * dx * dx * -0.81885 +
        dx * dy * dy * dy * 0.05594 +
        dx * dx * dx * dy * -0.05607 +
        dy * 0.01199 +
        dx * dx * dx * dy * dy * -0.00256 +
        dx * dy * dy * dy * dy * 0.00128 +
        dy * dy * 0.00022 +
        dx * dx * dx * dx * dx * -0.00022 +
        dx * dy * dy * dy * dy * dy * 0.00026
    ) / 3600.0
    
    return phi, lam


def wgs84_to_rd(lat, lon):
    """Converteer WGS84 (lat/lon) naar RD (EPSG:28992)"""
    x0 = 155000.0
    y0 = 463000.0
    phi0 = 52.15517440
    lam0 = 5.38720621
    
    dphi = (lat - phi0) * 3600.0
    dlam = (lon - lam0) * 3600.0
    
    rd_x = x0 + (
        dlam * 190094.945 +
        dphi * dlam * -11832.228 +
        dphi * dphi * dlam * -114.221 +
        dlam * dlam * dlam * -32.391 +
        dphi * -0.705 +
        dphi * dphi * dphi * dlam * -2.340 +
        dphi * dlam * dlam * dlam * -0.608 +
        dphi * dphi * dlam * dlam * dlam * -0.008 +
        dlam * dlam * dlam * dlam * dlam * 0.148
    ) / 1e5
    
    rd_y = y0 + (
        dphi * 309056.544 +
        dlam * dlam * 3638.893 +
        dphi * dphi * 73.077 +
        dlam * dlam * dlam * dlam * -157.984 +
        dphi * dlam * dlam * 59.788 +
        dphi * dphi * dphi * 0.433 +
        dlam * dlam * dphi * dphi * -6.439 +
        dphi * dlam * dlam * dlam * dlam * -0.032 +
        dphi * dphi * dlam * dlam * 0.092 +
        dlam * dlam * dlam * dlam * dlam * dlam * -0.054
    ) / 1e5
    
    return rd_x, rd_y


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
            rd_x = internal_to_meters(pos.X)
            rd_y = internal_to_meters(pos.Y)
            return rd_x, rd_y
    except:
        pass
    return None, None

def set_survey_point(doc, rd_x, rd_y):
    survey_point = get_survey_point(doc)
    if survey_point is None:
        raise Exception("Survey Point niet gevonden")
    
    x_internal = meters_to_internal(rd_x)
    y_internal = meters_to_internal(rd_y)
    new_position = XYZ(x_internal, y_internal, 0)
    current_pos = survey_point.SharedPosition
    move_vector = new_position - current_pos
    ElementTransformUtils.MoveElement(doc, survey_point.Id, move_vector)

def get_titleblock(doc):
    collector = FilteredElementCollector(doc)
    titleblocks = collector.OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    if titleblocks:
        return titleblocks[0].Id
    return None

def get_view_family_type(doc, view_family):
    collector = FilteredElementCollector(doc)
    types = collector.OfClass(ViewFamilyType).ToElements()
    for t in types:
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
    if titleblock_id is None:
        sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
    else:
        sheet = ViewSheet.Create(doc, titleblock_id)
    sheet.Name = sheet_name
    sheet.SheetNumber = sheet_number
    return sheet

def place_view_on_sheet(doc, sheet, view):
    try:
        if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
            return None
        center = XYZ(1.4, 0.5, 0)
        viewport = Viewport.Create(doc, sheet.Id, view.Id, center)
        return viewport
    except:
        return None

def get_or_create_sheet(doc, sheet_name, sheet_number):
    collector = FilteredElementCollector(doc)
    sheets = collector.OfClass(ViewSheet).ToElements()
    for s in sheets:
        if s.SheetNumber == sheet_number:
            return s, False
    sheet = create_sheet(doc, sheet_name, sheet_number)
    return sheet, True

def get_or_create_3d_view(doc, view_name):
    collector = FilteredElementCollector(doc)
    views_3d = collector.OfClass(View3D).ToElements()
    for v in views_3d:
        if not v.IsTemplate and v.Name == view_name:
            return v, False
    
    vft_id = get_view_family_type(doc, ViewFamily.ThreeDimensional)
    if vft_id is None:
        raise Exception("Geen 3D View type gevonden")
    
    view_3d = View3D.CreateIsometric(doc, vft_id)
    try:
        view_3d.Name = view_name
    except:
        view_3d.Name = view_name + "_" + str(int(time.time()))
    return view_3d, True

def import_image_to_view(doc, view, image_path, width_m):
    try:
        options = ImageTypeOptions(image_path, False, ImageTypeSource.Import)
        options.Resolution = 300
        image_type = ImageType.Create(doc, options)
        center = XYZ(0, 0, 0)
        placement_options = ImagePlacementOptions(center, BoxPlacement.Center)
        image_instance = ImageInstance.Create(doc, view, image_type.Id, placement_options)
        try:
            target_width = meters_to_internal(width_m)
            image_instance.Width = target_width
        except:
            pass
        return image_instance
    except Exception as e:
        raise Exception("Image import failed: {}".format(str(e)))


# =============================================================================
# GEOCODING
# =============================================================================

def geocode_address(address):
    try:
        encoded = address.replace(' ', '%20').replace(',', '%2C')
        url = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q={}&rows=5".format(encoded)
        response = web_request(url)
        data = json.loads(response)
        
        results = []
        if data.get('response', {}).get('numFound', 0) > 0:
            for d in data['response']['docs']:
                centroid = d.get('centroide_rd', '')
                if centroid:
                    coords = centroid.replace('POINT(', '').replace(')', '').split()
                    rd_x = float(coords[0])
                    rd_y = float(coords[1])
                    display_name = d.get('weergavenaam', address)
                    results.append((rd_x, rd_y, display_name))
        return results
    except Exception as e:
        raise Exception("Geocoding failed: {}".format(str(e)))


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
            url = "{base}?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER={layer}&STYLE=default&FORMAT=image/png&TILEMATRIXSET=EPSG:28992&TILEMATRIX={zoom}&TILEROW={row}&TILECOL={col}".format(
                base=base_url, layer=layer_name, zoom=zoomlevel, row=row, col=col
            )
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
    
    width_m = len(cols) * tile_width_m
    height_m = len(rows) * tile_width_m
    
    return combined, width_m, height_m

def download_wms_image(layer_config, rd_x, rd_y, bbox_size, image_size=2048):
    base_url = layer_config['url']
    layer_name = layer_config['layer']
    
    half = bbox_size / 2
    bbox = "{},{},{},{}".format(rd_x - half, rd_y - half, rd_x + half, rd_y + half)
    
    url = "{base}?SERVICE=WMS&VERSION=1.3.0&REQUEST=GetMap&LAYERS={layer}&STYLES=&CRS=EPSG:28992&BBOX={bbox}&WIDTH={size}&HEIGHT={size}&FORMAT=image/png&TRANSPARENT=true".format(
        base=base_url, layer=layer_name, bbox=bbox, size=image_size
    )
    
    bitmap = download_image(url)
    return bitmap, bbox_size, bbox_size

def get_wfs_features(layer_config, rd_x, rd_y, bbox_size, max_features=500):
    base_url = layer_config['url']
    layer_name = layer_config['layer']
    
    half = bbox_size / 2
    bbox = (rd_x - half, rd_y - half, rd_x + half, rd_y + half)
    
    url = "{base}?SERVICE=WFS&VERSION=2.0.0&REQUEST=GetFeature&TYPENAMES={layer}&BBOX={minx},{miny},{maxx},{maxy},urn:ogc:def:crs:EPSG::28992&COUNT={count}&OUTPUTFORMAT=application/json".format(
        base=base_url, layer=layer_name,
        minx=bbox[0], miny=bbox[1], maxx=bbox[2], maxy=bbox[3],
        count=max_features
    )
    
    response = web_request(url)
    data = json.loads(response)
    return data.get('features', [])

def get_ogcapi_features(layer_config, rd_x, rd_y, bbox_size, max_features=500):
    base_url = layer_config['url']
    collection = layer_config['collection']
    
    half = bbox_size / 2
    minx = rd_x - half
    miny = rd_y - half
    maxx = rd_x + half
    maxy = rd_y + half
    
    url = "{base}/collections/{collection}/items?bbox={minx},{miny},{maxx},{maxy}&bbox-crs=http://www.opengis.net/def/crs/EPSG/0/28992&crs=http://www.opengis.net/def/crs/EPSG/0/28992&limit={limit}&f=json".format(
        base=base_url, collection=collection,
        minx=minx, miny=miny, maxx=maxx, maxy=maxy,
        limit=max_features
    )
    
    response = web_request(url)
    data = json.loads(response)
    return data.get('features', [])

def extract_polygon_rings(geometry):
    geom_type = geometry.get('type', '')
    coords = geometry.get('coordinates', [])
    rings = []
    
    if geom_type == 'Polygon' and coords:
        if coords[0]:
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
        geom = feature.get('geometry', {})
        rings = extract_polygon_rings(geom)
        
        for ring_idx, ring in enumerate(rings):
            if len(ring) < 3:
                continue
            
            poly_offset = (feat_idx * 100 + ring_idx) * offset_value
            
            for i in range(len(ring)):
                x1, y1 = ring[i]
                x2, y2 = ring[(i + 1) % len(ring)]
                
                p1 = XYZ(
                    meters_to_internal(x1 - rd_x) + poly_offset,
                    meters_to_internal(y1 - rd_y) + poly_offset,
                    0
                )
                p2 = XYZ(
                    meters_to_internal(x2 - rd_x) + poly_offset,
                    meters_to_internal(y2 - rd_y) + poly_offset,
                    0
                )
                
                dist = p1.DistanceTo(p2)
                if dist > 0.01:
                    try:
                        line = Line.CreateBound(p1, p2)
                        doc.Create.NewDetailCurve(view, line)
                        lines_created += 1
                    except:
                        pass
    
    return lines_created


# =============================================================================
# 3D BAG IFC DOWNLOAD & IMPORT (v2025.09.03)
# =============================================================================

def find_3dbag_tile_for_location(rd_x, rd_y):
    """Zoek de 3D BAG tile via WFS BAG3D:Tiles layer.
    
    De WFS response bevat direct de ifc_download URL.
    
    Returns: (ifc_url, tile_id) of (None, error_message)
    """
    # Gebruik 3D BAG WFS Tiles layer (hoofdletter T!)
    half = 50  # Kleine bbox rond de locatie
    minx = rd_x - half
    miny = rd_y - half
    maxx = rd_x + half
    maxy = rd_y + half
    
    wfs_url = "https://data.3dbag.nl/api/BAG3D/wfs?SERVICE=WFS&VERSION=2.0.0&REQUEST=GetFeature&TYPENAMES=BAG3D:Tiles&BBOX={},{},{},{}&SRSNAME=EPSG:28992&OUTPUTFORMAT=json".format(
        minx, miny, maxx, maxy
    )
    
    print("WFS URL: {}".format(wfs_url))
    
    try:
        response = web_request(wfs_url)
        data = json.loads(response)
        features = data.get('features', [])
        
        if features:
            # Neem de eerste tile
            f = features[0]
            props = f.get('properties', {})
            tile_id = props.get('tile_id', '')
            ifc_url = props.get('ifc_download', '')
            
            if ifc_url:
                print("Tile gevonden: {} -> {}".format(tile_id, ifc_url))
                return ifc_url, tile_id
            else:
                return None, "Tile {} heeft geen IFC download".format(tile_id)
        else:
            return None, "Geen tiles gevonden voor RD ({}, {})".format(int(rd_x), int(rd_y))
        
    except Exception as e:
        print("Tile lookup fout: {}".format(str(e)))
        return None, str(e)


def download_3dbag_ifc_tile(rd_x, rd_y, output_folder):
    """Download 3D BAG IFC tile voor de opgegeven locatie.
    
    Haalt de IFC URL direct uit de WFS BAG3D:Tiles layer.
    
    Returns: (ifc_path, tile_id) of (None, error_message)
    """
    # Zoek de juiste tile en IFC URL via WFS
    ifc_url, tile_id = find_3dbag_tile_for_location(rd_x, rd_y)
    
    if not ifc_url:
        # tile_id bevat error message
        return None, tile_id
    
    print("IFC URL: {}".format(ifc_url))
    
    # Download het bestand
    client = WebClient()
    client.Headers.Add("User-Agent", "Mozilla/5.0 GIS2BIM/PyRevit")
    
    # Maak bestandsnaam van tile_id (bijv. "10/370/472" -> "10-370-472")
    safe_tile_id = tile_id.replace('/', '-')
    zip_filename = "3DBAG_{}.ifc.zip".format(safe_tile_id)
    zip_path = os.path.join(output_folder, zip_filename)
    
    try:
        print("Downloading IFC tile {}...".format(tile_id))
        client.DownloadFile(ifc_url, zip_path)
        
        if not os.path.exists(zip_path):
            return None, "Download mislukt: bestand niet aangemaakt"
        
        file_size = os.path.getsize(zip_path)
        print("Download voltooid: {} bytes ({} MB)".format(file_size, int(file_size/1024/1024)))
        
        if file_size < 1000:
            os.remove(zip_path)
            return None, "Download mislukt: bestand te klein ({} bytes)".format(file_size)
        
        # Unzip het IFC bestand
        ifc_folder = os.path.join(output_folder, "3DBAG_IFC")
        if not os.path.exists(ifc_folder):
            os.makedirs(ifc_folder)
        
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(ifc_folder)
        
        # Zoek het IFC bestand
        ifc_path = None
        for root, dirs, files in os.walk(ifc_folder):
            for f in files:
                if f.endswith('.ifc'):
                    ifc_path = os.path.join(root, f)
                    break
            if ifc_path:
                break
        
        # Cleanup zip
        try:
            os.remove(zip_path)
        except:
            pass
        
        if ifc_path:
            print("IFC bestand: {}".format(ifc_path))
            return ifc_path, tile_id
        else:
            return None, "Geen IFC bestand gevonden in zip"
        
    except Exception as e:
        try:
            if os.path.exists(zip_path):
                os.remove(zip_path)
        except:
            pass
        return None, "Download fout: {}".format(str(e))


def link_ifc_to_revit(doc, ifc_path, rd_x, rd_y):
    """Link een IFC bestand in Revit.
    
    Probeert meerdere methodes in volgorde van voorkeur:
    1. IFCImportOptions via RevitAPIIFC (beste methode)
    2. RevitLinkType.CreateFromIFC met cache RVT
    3. ExternalResourceReference methode
    
    Returns: (success, message)
    """
    if not os.path.exists(ifc_path):
        return False, "IFC bestand niet gevonden: {}".format(ifc_path)
    
    # Methode 1: IFCImportOptions (officiële IFC import API)
    try:
        print("Poging 1: IFCImportOptions via RevitAPIIFC...")
        
        # Laad de IFC API DLL
        clr.AddReference("RevitAPIIFC")
        from Autodesk.Revit.DB.IFC import IFCImportOptions
        
        # Maak import opties
        options = IFCImportOptions()
        
        # Stel opties in (versie-afhankelijk, wrap in try/except)
        try:
            options.AutoJoin = True
        except:
            pass
        
        # Voer de import uit
        # Let op: doc.Import voor IFC werkt mogelijk alleen in bepaalde Revit versies
        result = doc.Import(ifc_path, options)
        
        if result:
            print("IFC Import via IFCImportOptions succesvol")
            return True, "IFC succesvol geïmporteerd via RevitAPIIFC"
        
    except Exception as e:
        print("Methode 1 mislukt: {}".format(str(e)))
    
    # Methode 2: CreateFromIFC met cache RVT (standaard link workflow)
    try:
        print("Poging 2: RevitLinkType.CreateFromIFC...")
        
        from Autodesk.Revit.DB import (
            RevitLinkOptions, RevitLinkType, RevitLinkInstance,
            ModelPathUtils, SaveAsOptions, UnitSystem
        )
        
        # Pad voor het cache RVT bestand
        rvt_path = ifc_path + ".RVT"
        
        # Maak een nieuw leeg project document voor de IFC cache
        app = doc.Application
        
        print("Aanmaken cache document...")
        ifc_doc = app.NewProjectDocument(UnitSystem.Metric)
        
        # Sla op als RVT
        save_options = SaveAsOptions()
        save_options.OverwriteExistingFile = True
        ifc_doc.SaveAs(rvt_path, save_options)
        ifc_doc.Close(False)
        print("Cache RVT aangemaakt: {}".format(rvt_path))
        
        # Link opties
        link_options = RevitLinkOptions(False)
        
        # CreateFromIFC - link het IFC bestand
        result = RevitLinkType.CreateFromIFC(
            doc,
            ifc_path,
            rvt_path,
            True,  # recreateLink
            link_options
        )
        
        if result and result.ElementId != ElementId.InvalidElementId:
            print("IFC Link Type aangemaakt, ElementId: {}".format(result.ElementId))
            
            try:
                link_instance = RevitLinkInstance.Create(doc, result.ElementId)
                if link_instance:
                    print("IFC Link Instance geplaatst")
                    return True, "IFC succesvol gelinkt via CreateFromIFC"
            except Exception as inst_e:
                print("Instance aanmaken: {}".format(str(inst_e)))
                return True, "IFC Link Type aangemaakt (instance handmatig plaatsen)"
        
    except Exception as e:
        print("Methode 2 mislukt: {}".format(str(e)))
    
    # Methode 3: ExternalResourceReference (alternatieve API voor IFC links)
    try:
        print("Poging 3: ExternalResourceReference.CreateLocalResource...")
        
        from Autodesk.Revit.DB import (
            ExternalResourceReference, ExternalResourceTypes, PathType,
            RevitLinkOptions, RevitLinkType, RevitLinkInstance,
            ModelPathUtils, SaveAsOptions, UnitSystem
        )
        
        # Pad voor het cache RVT bestand
        rvt_path = ifc_path + ".RVT"
        
        # Maak nieuw cache document als het niet bestaat
        if not os.path.exists(rvt_path):
            app = doc.Application
            ifc_doc = app.NewProjectDocument(UnitSystem.Metric)
            save_options = SaveAsOptions()
            save_options.OverwriteExistingFile = True
            ifc_doc.SaveAs(rvt_path, save_options)
            ifc_doc.Close(False)
        
        # Maak ExternalResourceReference voor IFC
        ifc_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(ifc_path)
        
        # IFCLink resource type
        ifc_resource = ExternalResourceReference.CreateLocalResource(
            doc,
            ExternalResourceTypes.BuiltInExternalResourceTypes.IFCLink,
            ifc_model_path,
            PathType.Absolute
        )
        
        link_options = RevitLinkOptions(False)
        
        # CreateFromIFC met ExternalResourceReference
        result = RevitLinkType.CreateFromIFC(
            doc,
            ifc_resource,
            rvt_path,
            True,
            link_options
        )
        
        if result and result.ElementId != ElementId.InvalidElementId:
            print("IFC Link via ExternalResource aangemaakt")
            link_instance = RevitLinkInstance.Create(doc, result.ElementId)
            return True, "IFC succesvol gelinkt via ExternalResourceReference"
            
    except Exception as e:
        print("Methode 3 mislukt: {}".format(str(e)))
    
    # Geen automatische methode werkte - geef duidelijke instructies
    return True, "IFC opgeslagen:\n{}\n\nAutomatische import niet gelukt.\nGebruik in Revit:\n- Insert > Link IFC\n- of File > Open > IFC bestand".format(ifc_path)


def import_ifc_direct(doc, ifc_path):
    """Alternatieve methode: Direct IFC importeren.
    
    Returns: (success, message)
    """
    try:
        # Check of IFCImportOptions beschikbaar is
        try:
            from Autodesk.Revit.DB.IFC import IFCImportOptions
            
            options = IFCImportOptions()
            
            # Probeer opties in te stellen (niet alle bestaan in alle versies)
            try:
                options.AutoJoin = True
            except:
                pass
            
            # IFC import vereist typisch een nieuw document
            # Voor nu geven we het pad terug voor handmatige import
            
        except ImportError:
            pass
        
        # Informeer gebruiker - IFC is gedownload en klaar voor import
        return True, "IFC opgeslagen:\n{}\nIFC import fout: Gebruik File > Open IFC voor handmatige import".format(ifc_path)
        
    except Exception as e:
        return False, "IFC import fout: {}".format(str(e))


# =============================================================================
# 3D BAG CITYJSON DOWNLOAD & DIRECTSHAPE IMPORT (v6.4)
# =============================================================================

def get_3dbag_cityjson_bbox(rd_x, rd_y, bbox_size):
    """Haal 3D BAG gebouwen op via de CityJSON API binnen een bounding box.
    
    De API returnt CityJSONFeature objecten met LOD 1.2, 1.3 en 2.2 geometrie.
    
    Args:
        rd_x, rd_y: Centrum coordinaten in RD (EPSG:28992)
        bbox_size: Grootte van de bounding box in meters
    
    Returns: (cityjson_data, error_message) - data is dict met metadata en features
    """
    half = bbox_size / 2.0
    minx = rd_x - half
    miny = rd_y - half
    maxx = rd_x + half
    maxy = rd_y + half
    
    # 3D BAG API endpoint voor bbox query
    api_url = "https://api.3dbag.nl/collections/pand/items?bbox={},{},{},{}".format(
        minx, miny, maxx, maxy
    )
    
    print("3D BAG API URL: {}".format(api_url))
    
    try:
        response = web_request(api_url)
        data = json.loads(response)
        
        # Check voor metadata en features
        metadata = data.get('metadata', {})
        features = data.get('features', [])
        
        # Ook check voor 'feature' (enkelvoud) bij single building response
        if 'feature' in data and not features:
            features = [data['feature']]
        
        if not features:
            return None, "Geen gebouwen gevonden in bbox rond ({}, {})".format(int(rd_x), int(rd_y))
        
        print("Gevonden: {} gebouwen".format(len(features)))
        
        return {
            'metadata': metadata,
            'features': features
        }, None
        
    except Exception as e:
        print("API fout: {}".format(str(e)))
        return None, str(e)


def parse_cityjson_lod22_geometry(cityjson_data, rd_x, rd_y):
    """Parse CityJSON features en extraheer LOD 2.2 geometrie.
    
    CityJSON gebruikt indexed vertices met een transform voor compacte opslag.
    De transform staat in metadata, niet in individuele features.
    
    Args:
        cityjson_data: Dict met 'metadata' en 'features'
        rd_x, rd_y: Project centrum voor relatieve coordinaten
    
    Returns: List van building dicts met polygon faces (niet triangles)
    """
    metadata = cityjson_data.get('metadata', {})
    features = cityjson_data.get('features', [])
    
    # Haal transform uit metadata (scale en translate)
    # Let op: kunnen strings zijn ("0.001 0.001 0.001") of arrays
    transform = metadata.get('transform', {})
    scale_raw = transform.get('scale', [1.0, 1.0, 1.0])
    translate_raw = transform.get('translate', [0.0, 0.0, 0.0])
    
    # Parse scale en translate (kunnen strings of arrays zijn)
    if isinstance(scale_raw, str):
        scale = [float(x) for x in scale_raw.split()]
    else:
        scale = list(scale_raw) if scale_raw else [1.0, 1.0, 1.0]
    
    if isinstance(translate_raw, str):
        translate = [float(x) for x in translate_raw.split()]
    else:
        translate = list(translate_raw) if translate_raw else [0.0, 0.0, 0.0]
    
    print("Transform - scale: {}, translate: {}".format(scale, translate))
    
    buildings = []
    
    for feature in features:
        try:
            # CityJSON feature structuur
            city_objects = feature.get('CityObjects', {})
            vertices = feature.get('vertices', [])
            
            if not vertices:
                continue
            
            # Converteer vertices met transform (doe dit eenmaal per feature)
            converted_verts = []
            for v in vertices:
                vx, vy, vz = v[0], v[1], v[2]
                # Pas transform toe
                x = vx * scale[0] + translate[0]
                y = vy * scale[1] + translate[1]
                z = vz * scale[2] + translate[2]
                # Relatief t.o.v. project centrum
                converted_verts.append((x - rd_x, y - rd_y, z))
            
            # Zoek naar BuildingPart objecten met LOD 2.2 geometrie
            for obj_id, obj in city_objects.items():
                obj_type = obj.get('type', '')
                
                # BuildingPart bevat de gedetailleerde geometrie
                if obj_type not in ['BuildingPart', 'Building']:
                    continue
                
                geometry_list = obj.get('geometry', [])
                
                for geom in geometry_list:
                    lod = geom.get('lod', '')
                    
                    # We willen specifiek LOD 2.2
                    if str(lod) != '2.2':
                        continue
                    
                    geom_type = geom.get('type', '')
                    boundaries = geom.get('boundaries', [])
                    
                    if not boundaries:
                        continue
                    
                    # Verzamel alle polygon faces (NIET trianguleren)
                    polygon_faces = []
                    
                    if geom_type == 'Solid':
                        # Solid: boundaries[shell][surface][ring]
                        for shell in boundaries:
                            for surface in shell:
                                # surface is een list van rings (outer + holes)
                                # We nemen alleen de outer ring (eerste)
                                if surface and len(surface) > 0:
                                    outer_ring = surface[0] if isinstance(surface[0], list) else surface
                                    if len(outer_ring) >= 3:
                                        # Bewaar originele polygon (geen triangulatie)
                                        polygon_faces.append(list(outer_ring))
                    
                    elif geom_type == 'MultiSurface':
                        # MultiSurface: boundaries[surface][ring]
                        for surface in boundaries:
                            if surface and len(surface) > 0:
                                outer_ring = surface[0] if isinstance(surface[0], list) else surface
                                if len(outer_ring) >= 3:
                                    polygon_faces.append(list(outer_ring))
                    
                    if polygon_faces:
                        buildings.append({
                            'id': obj_id,
                            'vertices': converted_verts,
                            'polygon_faces': polygon_faces,  # Originele polygonen
                            'attributes': obj.get('attributes', {})
                        })
                        
                        print("Building {} - {} polygon faces".format(obj_id, len(polygon_faces)))
        
        except Exception as e:
            print("Feature parse error: {}".format(str(e)))
            continue
    
    return buildings


def triangulate_face(vertex_indices):
    """Trianguleer een face met meer dan 3 vertices (fan triangulation).
    
    Args:
        vertex_indices: List van vertex indices
    
    Returns: List van triangles, elk [(v0, v1, v2), ...]
    """
    if len(vertex_indices) < 3:
        return []
    
    if len(vertex_indices) == 3:
        return [tuple(vertex_indices)]
    
    # Fan triangulation vanuit eerste vertex
    triangles = []
    v0 = vertex_indices[0]
    for i in range(1, len(vertex_indices) - 1):
        v1 = vertex_indices[i]
        v2 = vertex_indices[i + 1]
        triangles.append((v0, v1, v2))
    
    return triangles


def create_directshapes_from_buildings(doc, buildings):
    """Maak Revit DirectShape elementen van building geometrie.
    
    Maakt polygon faces (niet triangles) en gebruikt TessellatedShapeBuilder
    met de originele CityJSON polygon boundaries.
    
    Args:
        doc: Revit document
        buildings: List van building dicts met vertices en polygon faces
    
    Returns: (count_created, error_message)
    """
    if not buildings:
        return 0, "Geen gebouwen om te importeren"
    
    count = 0
    last_error = ""
    
    for building in buildings:
        try:
            vertices = building['vertices']
            polygon_faces = building.get('polygon_faces', [])  # Originele polygonen
            obj_id = building.get('id', 'unknown')
            
            if not polygon_faces:
                print("Building {} - geen polygon faces".format(obj_id))
                continue
            
            if not vertices:
                print("Building {} - geen vertices".format(obj_id))
                continue
            
            print("Building {} - {} vertices, {} polygons".format(obj_id, len(vertices), len(polygon_faces)))
            
            # Converteer alle vertices naar XYZ punten
            xyz_verts = []
            for v in vertices:
                xyz_verts.append(XYZ(
                    meters_to_internal(v[0]), 
                    meters_to_internal(v[1]), 
                    meters_to_internal(v[2])
                ))
            
            # Gebruik TessellatedShapeBuilder met polygon faces (niet triangles)
            try:
                builder = TessellatedShapeBuilder()
                builder.OpenConnectedFaceSet(False)
                
                faces_added = 0
                for poly_indices in polygon_faces:
                    if len(poly_indices) < 3:
                        continue
                    
                    # Valideer indices
                    valid = True
                    for idx in poly_indices:
                        if idx >= len(xyz_verts):
                            valid = False
                            break
                    if not valid:
                        continue
                    
                    # Maak face vertices list met alle polygon punten
                    face_verts = List[XYZ]()
                    prev_pt = None
                    for idx in poly_indices:
                        pt = xyz_verts[idx]
                        # Skip duplicate punten
                        if prev_pt is not None and pt.DistanceTo(prev_pt) < 0.0001:
                            continue
                        face_verts.Add(pt)
                        prev_pt = pt
                    
                    if face_verts.Count < 3:
                        continue
                    
                    # Check of eerste en laatste punt niet hetzelfde zijn
                    if face_verts[0].DistanceTo(face_verts[face_verts.Count - 1]) < 0.0001:
                        # Verwijder laatste punt als het een duplicate is
                        new_face_verts = List[XYZ]()
                        for i in range(face_verts.Count - 1):
                            new_face_verts.Add(face_verts[i])
                        face_verts = new_face_verts
                    
                    if face_verts.Count < 3:
                        continue
                    
                    try:
                        builder.AddFace(TessellatedFace(face_verts, ElementId.InvalidElementId))
                        faces_added += 1
                    except:
                        pass
                
                print("  Faces toegevoegd: {}".format(faces_added))
                
                if faces_added == 0:
                    continue
                
                builder.CloseConnectedFaceSet()
                builder.Target = TessellatedShapeBuilderTarget.AnyGeometry
                builder.Fallback = TessellatedShapeBuilderFallback.Salvage
                
                result_outcome = builder.Build()
                result = builder.GetBuildResult()
                
                print("  Build outcome: {}".format(result.Outcome))
                
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
                        print("DirectShape aangemaakt: {}".format(obj_id))
                else:
                    last_error = "TessellatedShapeBuilder returned Nothing"
                    
            except Exception as build_e:
                print("  Build fout: {}".format(str(build_e)))
                last_error = str(build_e)
            
        except Exception as e:
            last_error = str(e)
            print("DirectShape fout voor {}: {}".format(building.get('id', '?'), str(e)))
            continue
    
    if count > 0:
        return count, None
    else:
        return 0, "Kon geen DirectShape elementen aanmaken. Laatste fout: {}".format(last_error)


def import_3dbag_cityjson(doc, rd_x, rd_y, bbox_size):
    """Hoofdfunctie: Download CityJSON en maak DirectShapes.
    
    Args:
        doc: Revit document
        rd_x, rd_y: Centrum coordinaten
        bbox_size: Bounding box grootte in meters
    
    Returns: (success, message)
    """
    print("=" * 50)
    print("3D BAG CityJSON Import")
    print("Locatie: RD {}, {}".format(int(rd_x), int(rd_y)))
    print("Bbox: {}m".format(bbox_size))
    print("=" * 50)
    
    # Stap 1: Download CityJSON data
    cityjson_data, error = get_3dbag_cityjson_bbox(rd_x, rd_y, bbox_size)
    
    if error:
        return False, "CityJSON download mislukt: {}".format(error)
    
    num_features = len(cityjson_data.get('features', []))
    print("Downloaded: {} features".format(num_features))
    
    # Stap 2: Parse LOD 2.2 geometrie
    buildings = parse_cityjson_lod22_geometry(cityjson_data, rd_x, rd_y)
    
    if not buildings:
        return False, "Geen LOD 2.2 geometrie gevonden in {} features".format(num_features)
    
    print("Parsed: {} gebouwen met LOD 2.2".format(len(buildings)))
    
    # Stap 3: Maak DirectShapes
    count, error = create_directshapes_from_buildings(doc, buildings)
    
    if error:
        return False, error
    
    return True, "{} gebouwen geimporteerd als DirectShape".format(count)


# =============================================================================
# MAIN DIALOG
# =============================================================================

class GIS2BIMDialog(BaseForm):
    """
    Hoofddialoog voor GIS2BIM tool.
    Split layout: Kaart links, Controls rechts.
    """
    
    def __init__(self, doc):
        super(GIS2BIMDialog, self).__init__(
            title="GIS2BIM - Nederlandse Geodata v6.4 (CityJSON)",
            width=1100,
            height=750,
            resizable=False
        )
        
        self.doc = doc
        self.rd_x, self.rd_y = get_project_location(doc)
        
        if not self.rd_x or not self.rd_y:
            self.rd_x = DEFAULT_RD_X
            self.rd_y = DEFAULT_RD_Y
        
        self.address_results = []
        self.bbox_size = 500
        
        self.layer_bbox_sizes = {}
        for key in GIS_LAYERS:
            self.layer_bbox_sizes[key] = 500
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Bouw de UI op met split layout"""
        scale = DPIScaler.scale
        
        margin = scale(15)
        map_width = scale(450)
        right_panel_x = map_width + margin * 2
        right_panel_width = self.ClientSize.Width - right_panel_x - margin
        
        self._setup_map_panel(margin, margin, map_width, self.ClientSize.Height - margin * 2)
        self._setup_controls_panel(right_panel_x, margin, right_panel_width)
    
    def _setup_map_panel(self, x, y, width, height):
        """Setup het kaart paneel met PictureBox"""
        scale = DPIScaler.scale
        
        lbl_map = UIFactory.create_label("Locatie Kaart (klik om marker te verplaatsen)", bold=True)
        lbl_map.Location = Point(x, y)
        lbl_map.Size = Size(width, scale(20))
        lbl_map.ForeColor = COLORS.PRIMARY
        self.Controls.Add(lbl_map)
        
        map_top = y + scale(25)
        map_height = height - scale(160)
        self.map_picture = WinForms.PictureBox()
        self.map_picture.Location = Point(x, map_top)
        self.map_picture.Size = Size(width, map_height)
        self.map_picture.BorderStyle = BorderStyle.FixedSingle
        self.map_picture.SizeMode = PictureBoxSizeMode.Zoom
        self.map_picture.Cursor = Cursors.Cross
        self.map_picture.BackColor = DrawingColor.FromArgb(240, 240, 240)
        self.map_picture.MouseClick += self._on_map_click
        self.Controls.Add(self.map_picture)
        
        self.map_image = None
        self.map_bbox_size = 500
        self.map_width_px = width
        self.map_height_px = map_height
        
        zoom_y = map_top + map_height + scale(5)
        lbl_zoom_out = UIFactory.create_label("-")
        lbl_zoom_out.Location = Point(x, zoom_y)
        lbl_zoom_out.Size = Size(scale(20), scale(20))
        self.Controls.Add(lbl_zoom_out)
        
        self.zoom_slider = WinForms.TrackBar()
        self.zoom_slider.Location = Point(x + scale(20), zoom_y)
        self.zoom_slider.Size = Size(width - scale(130), scale(20))
        self.zoom_slider.AutoSize = False
        self.zoom_slider.Minimum = 500
        self.zoom_slider.Maximum = 15000
        self.zoom_slider.Value = 500
        self.zoom_slider.TickStyle = WinForms.TickStyle.None
        self.zoom_slider.SmallChange = 100
        self.zoom_slider.LargeChange = 500
        self.zoom_slider.ValueChanged += self._on_zoom_changed
        self.Controls.Add(self.zoom_slider)
        
        lbl_zoom_in = UIFactory.create_label("+")
        lbl_zoom_in.Location = Point(x + width - scale(105), zoom_y)
        lbl_zoom_in.Size = Size(scale(20), scale(20))
        self.Controls.Add(lbl_zoom_in)
        
        self.lbl_zoom_value = UIFactory.create_label("500m")
        self.lbl_zoom_value.Location = Point(x + width - scale(80), zoom_y)
        self.lbl_zoom_value.Size = Size(scale(70), scale(20))
        self.Controls.Add(self.lbl_zoom_value)
        
        coord_y = zoom_y + scale(25)
        
        lbl_rdx = UIFactory.create_label("RD X:")
        lbl_rdx.Location = Point(x, coord_y + scale(5))
        self.Controls.Add(lbl_rdx)
        
        self.txt_rd_x = UIFactory.create_textbox()
        self.txt_rd_x.Location = Point(x + scale(45), coord_y)
        self.txt_rd_x.Size = Size(scale(90), scale(25))
        self.txt_rd_x.Text = str(int(self.rd_x))
        self.Controls.Add(self.txt_rd_x)
        
        lbl_rdy = UIFactory.create_label("RD Y:")
        lbl_rdy.Location = Point(x + scale(145), coord_y + scale(5))
        self.Controls.Add(lbl_rdy)
        
        self.txt_rd_y = UIFactory.create_textbox()
        self.txt_rd_y.Location = Point(x + scale(190), coord_y)
        self.txt_rd_y.Size = Size(scale(90), scale(25))
        self.txt_rd_y.Text = str(int(self.rd_y))
        self.Controls.Add(self.txt_rd_y)
        
        self.btn_goto = UIFactory.create_button("Ga naar", primary=True, width=80)
        self.btn_goto.Location = Point(x + scale(290), coord_y)
        self.btn_goto.Click += self._on_goto_coords
        self.Controls.Add(self.btn_goto)
        
        self.btn_refresh_map = UIFactory.create_button("Ververs", primary=False, width=70)
        self.btn_refresh_map.Location = Point(x + scale(380), coord_y)
        self.btn_refresh_map.Click += self._on_refresh_map
        self.Controls.Add(self.btn_refresh_map)
        
        self._load_map()
    
    def _on_zoom_changed(self, sender, args):
        self.map_bbox_size = self.zoom_slider.Value
        self.lbl_zoom_value.Text = "{}m".format(self.map_bbox_size)
        self._load_map()
    
    def _setup_controls_panel(self, x, y, width):
        """Setup het rechter paneel met controls"""
        scale = DPIScaler.scale
        
        section1 = UIFactory.create_label("1. Zoek Adres", bold=True, color=COLORS.PRIMARY)
        section1.Location = Point(x, y)
        self.Controls.Add(section1)
        y += scale(25)
        
        self.lbl_location = UIFactory.create_label(self._get_location_text())
        self.lbl_location.Location = Point(x, y)
        self.lbl_location.Size = Size(width, scale(20))
        self.lbl_location.ForeColor = COLORS.PRIMARY_DARK
        self.Controls.Add(self.lbl_location)
        y += scale(25)
        
        self.txt_address = UIFactory.create_textbox()
        self.txt_address.Location = Point(x, y)
        self.txt_address.Size = Size(width - scale(90), scale(25))
        self.txt_address.Text = DEFAULT_ADDRESS
        self.Controls.Add(self.txt_address)
        
        self.btn_search = UIFactory.create_button("Zoeken", primary=True, width=80)
        self.btn_search.Location = Point(x + width - scale(80), y)
        self.btn_search.Click += self._on_search
        self.Controls.Add(self.btn_search)
        y += scale(35)
        
        self.lst_address = UIFactory.create_listview([
            ("Adres", 280),
            ("RD X", 70),
            ("RD Y", 70)
        ])
        self.lst_address.Location = Point(x, y)
        self.lst_address.Size = Size(width, scale(80))
        self.lst_address.SelectedIndexChanged += self._on_address_select
        self.Controls.Add(self.lst_address)
        y += scale(85)
        
        self.btn_apply_loc = UIFactory.create_button("Locatie Toepassen op Project", primary=False, width=200)
        self.btn_apply_loc.Location = Point(x, y)
        self.btn_apply_loc.Click += self._on_apply_location
        self.Controls.Add(self.btn_apply_loc)
        y += scale(40)
        
        sep1 = UIFactory.create_separator(width=width)
        sep1.Location = Point(x, y)
        self.Controls.Add(sep1)
        y += scale(15)
        
        section2 = UIFactory.create_label("2. Selecteer Data", bold=True, color=COLORS.PRIMARY)
        section2.Location = Point(x, y)
        self.Controls.Add(section2)
        y += scale(25)
        
        self.lst_layers = UIFactory.create_listview([
            ("Laag", 160),
            ("Type", 60),
            ("Omtrek", 55),
            ("Status", 55)
        ])
        self.lst_layers.Location = Point(x, y)
        self.lst_layers.Size = Size(width, scale(150))
        self.lst_layers.CheckBoxes = True
        self.lst_layers.SelectedIndexChanged += self._on_layer_select
        self._populate_layers()
        self.Controls.Add(self.lst_layers)
        y += scale(155)
        
        lbl_layer_size = UIFactory.create_label("Omtrek geselecteerde laag:")
        lbl_layer_size.Location = Point(x, y + scale(3))
        self.Controls.Add(lbl_layer_size)
        
        self.num_layer_bbox = UIFactory.create_numeric(min_val=50, max_val=2000, default=500)
        self.num_layer_bbox.Location = Point(x + scale(160), y)
        self.num_layer_bbox.ValueChanged += self._on_layer_bbox_changed
        self.Controls.Add(self.num_layer_bbox)
        
        lbl_m2 = UIFactory.create_label("m")
        lbl_m2.Location = Point(x + scale(270), y + scale(3))
        self.Controls.Add(lbl_m2)
        y += scale(30)
        
        btn_x = x
        for text, handler in [("Alles", self._select_all), ("Geen", self._select_none), 
                               ("Kaarten", self._select_raster), ("3D", self._select_3d)]:
            btn = UIFactory.create_button(text, primary=False, width=60)
            btn.Location = Point(btn_x, y)
            btn.Click += handler
            self.Controls.Add(btn)
            btn_x += scale(65)
        y += scale(40)
        
        sep2 = UIFactory.create_separator(width=width)
        sep2.Location = Point(x, y)
        self.Controls.Add(sep2)
        y += scale(15)
        
        section3 = UIFactory.create_label("3. Uitvoeren", bold=True, color=COLORS.PRIMARY)
        section3.Location = Point(x, y)
        self.Controls.Add(section3)
        y += scale(25)
        
        self.lbl_status = UIFactory.create_label("Gereed - 3D BAG CityJSON LOD2.2")
        self.lbl_status.Location = Point(x, y)
        self.lbl_status.Size = Size(width, scale(20))
        self.lbl_status.ForeColor = COLORS.TEXT_SECONDARY
        self.Controls.Add(self.lbl_status)
        y += scale(25)
        
        self.progress = UIFactory.create_progressbar(width=width)
        self.progress.Location = Point(x, y)
        self.Controls.Add(self.progress)
        y += scale(35)
        
        self.btn_execute = UIFactory.create_button("Start Import", primary=True, width=120)
        self.btn_execute.Location = Point(x, y)
        self.btn_execute.Click += self._on_execute
        self.Controls.Add(self.btn_execute)
        
        self.btn_close = UIFactory.create_button("Sluiten", primary=False, width=100)
        self.btn_close.Location = Point(x + scale(130), y)
        self.btn_close.Click += self._on_cancel
        self.Controls.Add(self.btn_close)
    
    def _load_map(self):
        try:
            half = self.map_bbox_size / 2.0
            minx = self.rd_x - half
            miny = self.rd_y - half
            maxx = self.rd_x + half
            maxy = self.rd_y + half
            
            url = "https://service.pdok.nl/lv/bgt/wms/v1_0?SERVICE=WMS&VERSION=1.3.0&REQUEST=GetMap&LAYERS=standaardvisualisatie&STYLES=&CRS=EPSG:28992&BBOX={},{},{},{}&WIDTH={}&HEIGHT={}&FORMAT=image/png".format(
                minx, miny, maxx, maxy,
                int(self.map_width_px), int(self.map_height_px)
            )
            
            self.map_image = download_image(url)
            self._draw_marker_on_map()
            
        except Exception as e:
            try:
                half = self.map_bbox_size / 2.0
                minx = self.rd_x - half
                miny = self.rd_y - half
                maxx = self.rd_x + half
                maxy = self.rd_y + half
                
                url = "https://service.pdok.nl/hwh/luchtfotorgb/wms/v1_0?SERVICE=WMS&VERSION=1.3.0&REQUEST=GetMap&LAYERS=Actueel_orthoHR&STYLES=&CRS=EPSG:28992&BBOX={},{},{},{}&WIDTH={}&HEIGHT={}&FORMAT=image/png".format(
                    minx, miny, maxx, maxy,
                    int(self.map_width_px), int(self.map_height_px)
                )
                
                self.map_image = download_image(url)
                self._draw_marker_on_map()
                
            except Exception as e2:
                self.map_picture.Image = None
                print("Kaart laden mislukt: {}".format(str(e2)))
    
    def _draw_marker_on_map(self):
        if self.map_image is None:
            return
        
        img_copy = Bitmap(self.map_image.Width, self.map_image.Height)
        g = Graphics.FromImage(img_copy)
        g.DrawImage(self.map_image, 0, 0)
        
        cx = self.map_image.Width // 2
        cy = self.map_image.Height // 2
        
        pixels_per_meter = float(self.map_image.Width) / self.map_bbox_size
        
        if hasattr(self, 'lst_layers') and self.lst_layers.SelectedIndices.Count > 0:
            idx = self.lst_layers.SelectedIndices[0]
            item = self.lst_layers.Items[idx]
            layer_key = item.Tag
            layer_size = self.layer_bbox_sizes.get(layer_key, 500)
            
            box_size_px = int(layer_size * pixels_per_meter)
            box_half = box_size_px // 2
            
            pen_blue = Pen(DrawingColor.FromArgb(200, 0, 100, 200), 2)
            pen_blue.DashStyle = DashStyle.Dash
            g.DrawRectangle(pen_blue, cx - box_half, cy - box_half, box_size_px, box_size_px)
            pen_blue.Dispose()
            
            brush_blue = SolidBrush(DrawingColor.FromArgb(30, 0, 100, 200))
            g.FillRectangle(brush_blue, cx - box_half, cy - box_half, box_size_px, box_size_px)
            brush_blue.Dispose()
        
        marker_size = 20
        pen_red = Pen(DrawingColor.Red, 3)
        pen_white = Pen(DrawingColor.White, 1)
        brush_red = SolidBrush(DrawingColor.FromArgb(150, 255, 0, 0))
        
        g.FillEllipse(brush_red, cx - marker_size//2, cy - marker_size//2, marker_size, marker_size)
        g.DrawLine(pen_red, cx - marker_size, cy, cx + marker_size, cy)
        g.DrawLine(pen_red, cx, cy - marker_size, cx, cy + marker_size)
        g.DrawLine(pen_white, cx - marker_size + 1, cy, cx + marker_size - 1, cy)
        g.DrawLine(pen_white, cx, cy - marker_size + 1, cx, cy + marker_size - 1)
        
        pen_red.Dispose()
        pen_white.Dispose()
        brush_red.Dispose()
        g.Dispose()
        
        self.map_picture.Image = img_copy
    
    def _on_map_click(self, sender, args):
        if self.map_image is None:
            return
        
        pb = self.map_picture
        img = self.map_image
        
        scale_x = float(img.Width) / pb.Width
        scale_y = float(img.Height) / pb.Height
        scale = max(scale_x, scale_y)
        
        scaled_w = img.Width / scale
        scaled_h = img.Height / scale
        offset_x = (pb.Width - scaled_w) / 2.0
        offset_y = (pb.Height - scaled_h) / 2.0
        
        img_x = (args.X - offset_x) * scale
        img_y = (args.Y - offset_y) * scale
        
        half = self.map_bbox_size / 2.0
        frac_x = img_x / float(img.Width)
        frac_y = img_y / float(img.Height)
        
        new_rd_x = (self.rd_x - half) + (frac_x * self.map_bbox_size)
        new_rd_y = (self.rd_y + half) - (frac_y * self.map_bbox_size)
        
        self.rd_x = new_rd_x
        self.rd_y = new_rd_y
        self.txt_rd_x.Text = str(int(new_rd_x))
        self.txt_rd_y.Text = str(int(new_rd_y))
        self.lbl_location.Text = self._get_location_text()
        
        self._load_map()
    
    def _on_refresh_map(self, sender, args):
        self._load_map()
    
    def _on_goto_coords(self, sender, args):
        try:
            new_x = float(self.txt_rd_x.Text.replace(',', '.'))
            new_y = float(self.txt_rd_y.Text.replace(',', '.'))
            
            if new_x < 0 or new_x > 300000 or new_y < 300000 or new_y > 650000:
                self.show_warning("Coordinaten buiten Nederland.\nRD X: 0-300000, RD Y: 300000-650000")
                return
            
            self.rd_x = new_x
            self.rd_y = new_y
            self.lbl_location.Text = self._get_location_text()
            self._load_map()
            
        except ValueError:
            self.show_warning("Voer geldige numerieke coordinaten in.")
    
    def _get_location_text(self):
        if self.rd_x and self.rd_y:
            return "Locatie: RD X {} | RD Y {}".format(int(self.rd_x), int(self.rd_y))
        return "Locatie: Niet ingesteld"
    
    def _populate_layers(self):
        for key, layer in GIS_LAYERS.items():
            item = ListViewItem(layer['name'])
            type_display = {
                'wmts': 'Kaart', 'wms': 'Kaart',
                'wfs': 'Lijnen', 'ogcapi': 'Lijnen',
                '3dbag_ifc': 'IFC 3D',
                '3dbag_cityjson': 'LOD2.2'
            }.get(layer['type'], layer['type'])
            item.SubItems.Add(type_display)
            item.SubItems.Add(str(self.layer_bbox_sizes.get(key, 500)) + "m")
            item.SubItems.Add("Gereed")
            item.Tag = key
            item.Checked = False
            self.lst_layers.Items.Add(item)
    
    def _on_layer_select(self, sender, args):
        if self.lst_layers.SelectedIndices.Count > 0:
            idx = self.lst_layers.SelectedIndices[0]
            item = self.lst_layers.Items[idx]
            layer_key = item.Tag
            bbox_size = self.layer_bbox_sizes.get(layer_key, 500)
            self.num_layer_bbox.Value = bbox_size
            self._draw_marker_on_map()
    
    def _on_layer_bbox_changed(self, sender, args):
        if self.lst_layers.SelectedIndices.Count > 0:
            idx = self.lst_layers.SelectedIndices[0]
            item = self.lst_layers.Items[idx]
            layer_key = item.Tag
            new_size = int(self.num_layer_bbox.Value)
            self.layer_bbox_sizes[layer_key] = new_size
            item.SubItems[2].Text = str(new_size) + "m"
            self._draw_marker_on_map()
    
    def _on_search(self, sender, args):
        address = self.txt_address.Text.strip()
        if not address:
            self.show_warning("Voer een adres in om te zoeken.")
            return
        
        try:
            self.lbl_status.Text = "Zoeken..."
            Application.DoEvents()
            
            results = geocode_address(address)
            self.lst_address.Items.Clear()
            self.address_results = results
            
            for rx, ry, name in results:
                item = ListViewItem(name)
                item.SubItems.Add(str(int(rx)))
                item.SubItems.Add(str(int(ry)))
                self.lst_address.Items.Add(item)
            
            self.lbl_status.Text = "{} resultaten".format(len(results))
            
            if not results:
                self.show_info("Geen resultaten gevonden.")
                
        except Exception as e:
            self.show_error("Zoekfout: {}".format(str(e)))
            self.lbl_status.Text = "Fout"
    
    def _on_address_select(self, sender, args):
        if self.lst_address.SelectedIndices.Count > 0:
            idx = self.lst_address.SelectedIndices[0]
            if idx < len(self.address_results):
                rx, ry, _ = self.address_results[idx]
                self.rd_x = rx
                self.rd_y = ry
                self.txt_rd_x.Text = str(int(rx))
                self.txt_rd_y.Text = str(int(ry))
                self.lbl_location.Text = self._get_location_text()
                self._load_map()
    
    def _on_apply_location(self, sender, args):
        if not self.rd_x or not self.rd_y:
            self.show_warning("Selecteer eerst een locatie.")
            return
        
        try:
            with Transaction(self.doc, "GIS2BIM - Set Location") as t:
                t.Start()
                set_survey_point(self.doc, self.rd_x, self.rd_y)
                t.Commit()
            
            self.show_info("Locatie ingesteld!\n\nRD X: {}\nRD Y: {}".format(
                int(self.rd_x), int(self.rd_y)
            ))
            self.lbl_status.Text = "Locatie ingesteld"
            
        except Exception as e:
            self.show_error("Fout: {}".format(str(e)))
    
    def _select_all(self, sender, args):
        for item in self.lst_layers.Items:
            item.Checked = True
    
    def _select_none(self, sender, args):
        for item in self.lst_layers.Items:
            item.Checked = False
    
    def _select_raster(self, sender, args):
        for item in self.lst_layers.Items:
            layer = GIS_LAYERS.get(item.Tag, {})
            item.Checked = layer.get('type') in ['wmts', 'wms']
    
    def _select_3d(self, sender, args):
        for item in self.lst_layers.Items:
            layer = GIS_LAYERS.get(item.Tag, {})
            item.Checked = layer.get('type') in ['3dbag_ifc', '3dbag_cityjson']
    
    def _on_execute(self, sender, args):
        if not self.rd_x or not self.rd_y:
            self.show_warning("Stel eerst een locatie in.")
            return
        
        selected = [item.Tag for item in self.lst_layers.Items if item.Checked]
        
        if not selected:
            self.show_warning("Selecteer minimaal één laag.")
            return
        
        if not self.show_question("Wilt u {} lagen importeren?".format(len(selected))):
            return
        
        if self.doc.PathName:
            output_folder = os.path.join(os.path.dirname(self.doc.PathName), "GIS2BIM")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
        else:
            output_folder = tempfile.gettempdir()
        
        results_summary = []
        total_steps = len(selected)
        
        try:
            with Transaction(self.doc, "GIS2BIM - Import Data") as t:
                t.Start()
                
                for i, layer_key in enumerate(selected):
                    layer = GIS_LAYERS[layer_key]
                    sheet_name = layer['sheet_name']
                    sheet_number = "GIS-{:02d}".format(i + 1)
                    bbox_size = self.layer_bbox_sizes.get(layer_key, 500)
                    
                    self.lbl_status.Text = "Importeren: {} ({}m)...".format(layer['name'], bbox_size)
                    self.progress.Value = int((i / float(total_steps)) * 100)
                    Application.DoEvents()
                    
                    for item in self.lst_layers.Items:
                        if item.Tag == layer_key:
                            item.SubItems[3].Text = "Bezig..."
                    Application.DoEvents()
                    
                    try:
                        result_text = self._import_layer(
                            layer_key, layer, sheet_name, sheet_number,
                            self.rd_x, self.rd_y, bbox_size, output_folder
                        )
                        results_summary.append("{}: {}".format(layer['name'], result_text))
                        
                        for item in self.lst_layers.Items:
                            if item.Tag == layer_key:
                                item.SubItems[3].Text = "OK"
                        
                    except Exception as e:
                        results_summary.append("{}: Fout - {}".format(layer['name'], str(e)))
                        for item in self.lst_layers.Items:
                            if item.Tag == layer_key:
                                item.SubItems[3].Text = "Fout"
                    
                    Application.DoEvents()
                
                t.Commit()
            
            self.progress.Value = 100
            self.lbl_status.Text = "Voltooid!"
            
            summary = "Import voltooid!\n\n"
            for r in results_summary:
                summary += "- {}\n".format(r)
            
            self.show_info(summary, "Voltooid")
            
        except Exception as e:
            self.show_error("Import fout: {}".format(str(e)))
            self.lbl_status.Text = "Fout"
    
    def _import_layer(self, layer_key, layer, sheet_name, sheet_number, rd_x, rd_y, bbox_size, output_folder):
        """Import een enkele laag"""
        doc = self.doc
        sheet = None
        view = None
        
        if layer['type'] not in ['3dbag_ifc', '3dbag_cityjson']:
            sheet, _ = get_or_create_sheet(doc, sheet_name, sheet_number)
            
            view_name = "GIS_{}".format(layer_key)
            collector = FilteredElementCollector(doc)
            views = collector.OfClass(ViewDrafting).ToElements()
            for v in views:
                if v.Name == view_name:
                    view = v
                    break
            
            if view is None:
                view = create_drafting_view(doc, view_name)
        
        if layer['type'] == 'wmts':
            bitmap, width_m, _ = download_wmts_tiles(layer, rd_x, rd_y, bbox_size)
            save_path = os.path.join(output_folder, "GIS2BIM_{}.png".format(layer_key))
            bitmap.Save(save_path, ImageFormat.Png)
            import_image_to_view(doc, view, save_path, width_m)
            if sheet and view:
                place_view_on_sheet(doc, sheet, view)
            return "OK"
        
        elif layer['type'] == 'wms':
            bitmap, width_m, _ = download_wms_image(layer, rd_x, rd_y, bbox_size)
            save_path = os.path.join(output_folder, "GIS2BIM_{}.png".format(layer_key))
            bitmap.Save(save_path, ImageFormat.Png)
            import_image_to_view(doc, view, save_path, width_m)
            if sheet and view:
                place_view_on_sheet(doc, sheet, view)
            return "OK"
        
        elif layer['type'] == 'wfs':
            features = get_wfs_features(layer, rd_x, rd_y, bbox_size)
            if features and view:
                lines = create_detail_lines_in_view(doc, view, features, rd_x, rd_y)
                if sheet and view:
                    place_view_on_sheet(doc, sheet, view)
                return "{} lijnen".format(lines)
            return "Geen data"
        
        elif layer['type'] == 'ogcapi':
            features = get_ogcapi_features(layer, rd_x, rd_y, bbox_size)
            if features and view:
                lines = create_detail_lines_in_view(doc, view, features, rd_x, rd_y)
                if sheet and view:
                    place_view_on_sheet(doc, sheet, view)
                return "{} lijnen".format(lines)
            return "Geen data"
        
        elif layer['type'] == '3dbag_ifc':
            # Download 3D BAG IFC tile
            self.lbl_status.Text = "Downloading 3D BAG IFC tile..."
            Application.DoEvents()
            
            ifc_path, tile_id = download_3dbag_ifc_tile(rd_x, rd_y, output_folder)
            
            if not ifc_path:
                # tile_id bevat error message
                raise Exception(tile_id)
            
            self.lbl_status.Text = "IFC tile gedownload: {}".format(tile_id)
            Application.DoEvents()
            
            # Probeer IFC te linken/importeren
            success, message = link_ifc_to_revit(doc, ifc_path, rd_x, rd_y)
            
            if success:
                return "IFC tile {} gelinkt".format(tile_id)
            else:
                # Als automatische import niet lukt, informeer gebruiker
                return "IFC opgeslagen: {}\n{}".format(ifc_path, message)
        
        elif layer['type'] == '3dbag_cityjson':
            # CityJSON DirectShape import
            self.lbl_status.Text = "Downloading 3D BAG CityJSON..."
            Application.DoEvents()
            
            success, message = import_3dbag_cityjson(doc, rd_x, rd_y, bbox_size)
            
            if success:
                return message
            else:
                raise Exception(message)
        
        return "Onbekend"
    
    def _on_cancel(self, sender, args):
        self.close_cancel()


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main entry point"""
    doc = revit.doc
    
    result, dialog = run_dialog(GIS2BIMDialog, doc)
    
    if result == DialogResult.OK:
        TaskDialog.Show("GIS2BIM", "Import succesvol afgerond!")


if __name__ == "__main__":
    main()
