# Pointcloud Slice Tool

## Beschrijving
Deze tool extraheert punten uit een pointcloud binnen een gedefinieerde "slice" en genereert automatisch een Revit element (vloer of wand) op basis van die punten.

## Gebruik

### Stap 1: Pointcloud selecteren
- Als er een pointcloud geselecteerd is, wordt deze automatisch gebruikt
- Anders toont de tool een lijst van beschikbare pointclouds in het project

### Stap 2: Slice configureren

**Slice type:**
- **Horizontaal (Vloer/Plafond)**: Snijdt horizontaal door de pointcloud op een bepaalde Z-hoogte
- **Verticaal X (Wand)**: Snijdt verticaal langs de X-as
- **Verticaal Y (Wand)**: Snijdt verticaal langs de Y-as

**Parameters:**
- **Slice positie (mm)**: De positie van het snijvlak (hoogte voor horizontaal, X/Y positie voor verticaal)
- **Slice dikte (mm)**: Hoe dik de slice is (punten binnen dit bereik worden meegenomen)
- **Max punten**: Maximum aantal punten om te verwerken (meer = nauwkeuriger maar langzamer)

### Stap 3: Element configureren
- **Level**: Selecteer het level waaraan het element gekoppeld wordt
- **Vloer/Wand type**: Selecteer het type element dat gegenereerd wordt

### Stap 4: Preview of Genereren
- **Preview Punten**: Toont hoeveel punten gevonden zijn en hun bounds
- **Genereer Element**: Maakt het Revit element aan

## Tips

### Horizontale slices (vloeren/plafonds)
- Gebruik een dikte van 50-200mm voor beste resultaten
- De tool maakt een rechthoekige vloer op basis van de bounding box van de punten
- Voor complexere vormen zijn meerdere slices nodig

### Verticale slices (wanden)
- Positioneer de slice precies op de wand
- Gebruik een dikte die iets groter is dan de wanddikte
- De tool maakt een rechte wand van begin tot eind

## Beperkingen

1. **Bounding box geometrie**: De tool genereert rechthoekige elementen gebaseerd op de bounding box van gevonden punten. Complexe vormen worden niet automatisch gevolgd.

2. **Performance**: Grote pointclouds met veel punten kunnen langzaam zijn. Begin met minder max punten en verhoog indien nodig.

3. **Nauwkeurigheid**: De resultaten zijn afhankelijk van de kwaliteit en dichtheid van de pointcloud.

## Technische details

### Revit API
- Gebruikt `PointCloudInstance.GetPoints()` met `PointCloudFilter`
- Floor wordt gemaakt via `Floor.Create()` met CurveLoop
- Wall wordt gemaakt via `Wall.Create()` met Line

### Dependencies
- pyRevit
- ui_template.py (JMK lib)

## Versie
1.0 - Initiële release

## Auteur
JMK
