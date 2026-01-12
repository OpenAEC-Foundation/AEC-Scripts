# -*- coding: utf-8 -*-
"""
CityJSON Parser voor 3D BAG - Aparte module voor betere onderhoudbaarheid
BELANGRIJK: 3D BAG API returnt vertices als strings "x y z", niet als arrays!
"""


def parse_cityjson_vertices(raw_vertices):
    """Parse raw vertices die als strings of arrays kunnen komen.
    
    Args:
        raw_vertices: List van vertices, kunnen strings ("x y z") of arrays ([x,y,z]) zijn
    
    Returns:
        List van [x, y, z] integer arrays
    """
    vertices = []
    for v in raw_vertices:
        if isinstance(v, str):
            # String format: "12158 4854 485"
            parts = v.split()
            if len(parts) >= 3:
                vertices.append([int(parts[0]), int(parts[1]), int(parts[2])])
        elif isinstance(v, (list, tuple)) and len(v) >= 3:
            vertices.append([v[0], v[1], v[2]])
    return vertices


def triangulate_polygon(vertex_indices):
    """Trianguleer een polygon met fan triangulation.
    
    Args:
        vertex_indices: List van vertex indices
    
    Returns:
        List van triangles [(v0, v1, v2), ...]
    """
    if len(vertex_indices) < 3:
        return []
    
    if len(vertex_indices) == 3:
        return [tuple(vertex_indices)]
    
    # Fan triangulation vanuit eerste vertex
    triangles = []
    v0 = vertex_indices[0]
    for i in range(1, len(vertex_indices) - 1):
        triangles.append((v0, vertex_indices[i], vertex_indices[i + 1]))
    
    return triangles


def extract_lod22_faces(geometry):
    """Extraheer faces uit CityJSON geometry object.
    
    Args:
        geometry: CityJSON geometry dict met type en boundaries
    
    Returns:
        List van triangulated faces
    """
    geom_type = geometry.get('type', '')
    boundaries = geometry.get('boundaries', [])
    
    if not boundaries:
        return []
    
    faces = []
    
    if geom_type == 'Solid':
        # Solid: boundaries[shell][surface][rings]
        # Elke surface heeft 1+ rings (outer + optional inner)
        for shell in boundaries:
            for surface in shell:
                # surface is een lijst van rings, neem outer ring (eerste)
                if surface and len(surface) > 0:
                    ring = surface[0] if isinstance(surface[0], list) else surface
                    if len(ring) >= 3:
                        faces.extend(triangulate_polygon(ring))
    
    elif geom_type == 'MultiSurface':
        # MultiSurface: boundaries[surface][rings]
        for surface in boundaries:
            if surface and len(surface) > 0:
                ring = surface[0] if isinstance(surface[0], list) else surface
                if len(ring) >= 3:
                    faces.extend(triangulate_polygon(ring))
    
    return faces


def transform_vertices(vertices, scale, translate, rd_x, rd_y):
    """Pas CityJSON transform toe en maak relatief t.o.v. project centrum.
    
    Args:
        vertices: List van [x, y, z] integer vertices
        scale: [sx, sy, sz] scale factors
        translate: [tx, ty, tz] translate values
        rd_x, rd_y: Project centrum coordinaten
    
    Returns:
        List van (x, y, z) tuples in relatieve meters
    """
    converted = []
    for vx, vy, vz in vertices:
        x = float(vx) * scale[0] + translate[0]
        y = float(vy) * scale[1] + translate[1]
        z = float(vz) * scale[2] + translate[2]
        converted.append((x - rd_x, y - rd_y, z))
    return converted
