# Installatie & Troubleshooting Guide

## Inhoudsopgave
1. [Installatie pyRevit](#installatie-pyrevit)
2. [Script Installeren](#script-installeren)
3. [Eerste Gebruik](#eerste-gebruik)
4. [Veelvoorkomende Problemen](#veelvoorkomende-problemen)
5. [Advanced Troubleshooting](#advanced-troubleshooting)

---

## Installatie pyRevit

### Stap 1: Download pyRevit

1. Ga naar [https://github.com/pyrevitlabs/pyRevit/releases](https://github.com/pyrevitlabs/pyRevit/releases)
2. Download de laatste **pyRevit_[versie]_signed.exe**
3. Of gebruik direct link: [pyRevit Installer](https://github.com/pyrevitlabs/pyRevit/releases/latest)

### Stap 2: Installeer pyRevit

```
1. Run pyRevit_xxx_signed.exe als Administrator
2. Volg de installatie wizard
3. Kies installatie locatie (standaard: C:\pyRevit)
4. Selecteer Revit versie(s) waarvoor je het wilt installeren
5. Klik "Install"
6. Wacht tot installatie compleet is
```

### Stap 3: Verifieer Installatie

```
1. Start Revit
2. Kijk of "pyRevit" tab verschijnt in de ribbon
3. Klik op pyRevit tab → Settings → Check/Verify
4. Als alles groen is: installatie succesvol!
```

### Alternatief: CLI Installatie

Voor gevorderde gebruikers via command line:

```powershell
# Download installer
Invoke-WebRequest -Uri "https://github.com/pyrevitlabs/pyRevit/releases/latest/download/pyRevit_xxx_signed.exe" -OutFile "pyRevit_installer.exe"

# Installeer silent mode
.\pyRevit_installer.exe /SILENT /ALLUSERS
```

---

## Script Installeren

### Methode 1: Direct in pyRevit Extensions Folder

```
1. Zoek je pyRevit extensions folder:
   - pyRevit tab → Settings → Custom Extension Directories
   - Standaard: C:\Users\[USER]\AppData\Roaming\pyRevit\Extensions
   
2. Maak nieuwe folder: 
   C:\...\Extensions\3BM.extension\
   
3. Binnen deze folder, maak structuur:
   3BM.extension\
   └── 3BM.tab\
       └── Kozijnen.panel\
           └── Kozijn Creator.pushbutton\
               ├── icon.png (optioneel)
               └── script.py (= kozijn_family_creator.py)
   
4. Herlaad pyRevit: pyRevit tab → Reload
```

### Methode 2: Via pyRevit Custom Script Location

```
1. Kopieer kozijn_family_creator.py naar een eigen folder, bijv:
   C:\RevitScripts\3BM\
   
2. In pyRevit: Settings → Custom Extension Directories
3. Voeg je folder toe: C:\RevitScripts\3BM\
4. Reload pyRevit
```

### Methode 3: Direct Run (Quick & Dirty)

```
1. Open RevitPythonShell (Add-ins → RevitPythonShell)
2. File → Open → Selecteer kozijn_family_creator.py
3. Run (F5)
```

---

## Eerste Gebruik

### Checklist voor Eerste Run:

- [ ] **Revit versie**: 2019 of hoger
- [ ] **pyRevit geïnstalleerd**: Check pyRevit tab in ribbon
- [ ] **Script geplaatst**: In juiste folder structuur
- [ ] **Revit herstart**: Na script installatie
- [ ] **Project open**: Heb een project open waar je kozijn in wilt
- [ ] **Schrijfrechten**: Voor opslaan .rfa bestanden

### Test Run:

```
1. Klik op "Kozijn Creator" button in pyRevit
2. Dialog moet openen met parameters
3. Test met standaard waarden:
   - Vakken: 2 x 1
   - Breedte: 1600 mm
   - Hoogte: 1200 mm
4. Klik "Family Aanmaken"
5. Kies locatie en bestandsnaam
6. Sla op
7. Laad in project? → Ja
8. Check of family beschikbaar is in Project Browser
```

### Eerste Family Plaatsen:

```
1. Architecture tab → Component
2. Load Family (als nog niet geladen)
3. Selecteer je nieuwe kozijn
4. Klik in view om te plaatsen
5. Check in Properties panel of alle parameters zichtbaar zijn
```

---

## Veelvoorkomende Problemen

### Probleem 1: Script start niet

**Symptomen:**
- Klikken op button doet niets
- Geen error message
- Script verschijnt niet in pyRevit

**Oplossingen:**

```python
# Check 1: pyRevit correct geïnstalleerd?
pyRevit tab → Settings → About
# Moet versie nummer tonen

# Check 2: Script in juiste folder?
# Correcte structuur:
3BM.extension\
└── 3BM.tab\
    └── Kozijnen.panel\
        └── Kozijn Creator.pushbutton\
            └── script.py

# Check 3: Script naam correct?
# Moet heten: script.py (NIET kozijn_family_creator.py in deze folder)
# Of rename kozijn_family_creator.py → __init__.py

# Check 4: Herlaad pyRevit
pyRevit tab → Reload
# Of herstart Revit

# Check 5: Check pyRevit output window
pyRevit tab → Settings → Output Window
# Kijk voor error messages
```

### Probleem 2: Import Errors

**Error Message:**
```
ImportError: No module named 'Autodesk.Revit.DB'
```

**Oplossing:**
```python
# pyRevit mist Revit API references
# Fix:
1. Herinstalleer pyRevit
2. Zorg dat je Revit EERST installeert, DAN pyRevit
3. Run pyRevit installer opnieuw en selecteer je Revit versie

# Als blijft falen:
# Check Revit API DLLs locatie:
C:\Program Files\Autodesk\Revit 20XX\RevitAPI.dll
C:\Program Files\Autodesk\Revit 20XX\RevitAPIUI.dll
```

### Probleem 3: Family Template niet gevonden

**Error:**
```
Error: Kon geen Generic Model template vinden
```

**Oplossing:**
```python
# Check template locatie:
# Revit 2024:
C:\ProgramData\Autodesk\RVT 2024\Family Templates\English\Metric Generic Model.rft

# Als niet standaard locatie:
# Pas script aan, regel 221-230:
template_path = "C:\\JouwPad\\Templates\\Metric Generic Model.rft"

# Of kopieer template naar standaard locatie
```

### Probleem 4: Kan family niet opslaan

**Error:**
```
Access denied / Geen schrijfrechten
```

**Oplossing:**
```powershell
# Check folder permissions:
# Right-click op folder → Properties → Security
# Zorg dat je account "Full Control" heeft

# Alternatief: Sla op naar andere locatie
# Bijv: C:\Temp\ of je Documents folder
```

### Probleem 5: Family niet zichtbaar na laden

**Symptomen:**
- Family geladen maar niet zichtbaar in view
- Geen error, maar component ontbreekt

**Oplossing:**
```
1. Check Visibility Graphics (VG of VV)
2. Generic Models → Check moet aan staan
3. View Range → Check of elementen binnen range
4. View Filter → Zorg dat Generic Models niet gefilterd zijn
5. Temporary Hide/Isolate → Reset (blauwe rand)
```

### Probleem 6: Parameters niet aanpasbaar

**Symptomen:**
- Parameters grijs/disabled
- Waarden niet te wijzigen

**Oplossing:**
```
# Parameters zijn Type parameters, niet Instance
1. Selecteer family instance
2. Properties panel → Edit Type button
3. Pas parameters aan in Type Properties dialog
4. OK → Apply

# Of:
# Wijzig in Family Editor (.rfa file) en reload
```

### Probleem 7: Geometrie ziet er vreemd uit

**Symptomen:**
- Kozijn onderdelen door elkaar
- Glas buiten kozijn
- Rare extrusions

**Diagnose & Fix:**
```python
# Check parameters:
1. Totale afmetingen realistisch? (min 400mm)
2. Randstijl niet te breed voor totale breedte?
3. Aantal vakken past bij afmetingen?

# Bereken minimale breedte:
min_breedte = (2 × randstijl_breedte) + (aantal_vakken - 1) × tussenstijl + (aantal_vakken × 200mm)

# Als geometrie issue blijft:
# Open .rfa in Family Editor
# View → Visibility/Graphics → Check layers
# 3D view → Visual Style → Shaded om overlaps te zien
```

### Probleem 8: Script traag / hangt

**Symptomen:**
- Script reageert niet
- Revit "Not Responding"
- Lange wachttijd

**Oplossing:**
```python
# Voor grote aantallen vakken (bijv 10x5):
# Verwachte performance:
# - 1-2 vakken: < 5 sec
# - 3-4 vakken: 5-15 sec  
# - 5-10 vakken: 15-30 sec
# - 10+ vakken: 30-60 sec

# Als langer:
1. Check Revit versie (2024 sneller dan 2019)
2. Sluit andere Revit documenten
3. Sluit zware apps (Chrome, etc)
4. Herstart Revit en probeer opnieuw

# Voor batch processing:
# Wacht tot huidige family klaar is voor volgende
```

---

## Advanced Troubleshooting

### Python Console Debugging

```python
# In script, voeg toe voor debugging:
import sys
print("Debug: Script gestart")
print("Python versie:", sys.version)
print("Revit versie:", app.VersionNumber)

# Check parameters:
print("Totale breedte:", self.params.totale_breedte)
print("Aantal vakken H:", self.params.aantal_vakken_h)

# Output verschijnt in pyRevit Output Window
```

### Log Files

```
# pyRevit logs locatie:
C:\Users\[USER]\AppData\Roaming\pyRevit\

# Check:
pyRevit_run.log         # Script execution logs
pyRevit_routes.log      # Extension loading
pyRevit_usage.log       # Usage statistics

# Open in Notepad++ of VS Code voor analyse
```

### API Version Compatibility

```python
# Check Revit API versie:
print(app.VersionNumber)  # 2024, 2023, etc
print(app.VersionBuild)   # Build number

# Known compatibility:
# Revit 2019: API versie 19.0
# Revit 2020: API versie 20.0
# Revit 2021: API versie 21.0
# Revit 2022: API versie 22.0
# Revit 2023: API versie 23.0
# Revit 2024: API versie 24.0
# Revit 2025: API versie 25.0

# Script should work 2019+
```

### Performance Optimization

```python
# Voor snellere family creatie:
# 1. Disable views regeneration tijdens creatie
# 2. Use transactions efficiently  
# 3. Batch geometry creation

# In script, optimalisatie (advanced):
from Autodesk.Revit.DB import TransactionStatus

trans.Start()
# Set regeneration option
opt = TransactionOptions()
opt.RegenerationOfViews = False
# ... geometry creation ...
trans.Commit()
```

### Custom Error Handling

```python
# Voeg toe aan script voor betere errors:
import traceback

try:
    # Your code here
    creator = KozijnFamilyCreator(params)
    creator.create_family()
except Exception as e:
    error_msg = "FOUT bij kozijn creatie:\n\n"
    error_msg += str(e) + "\n\n"
    error_msg += "Traceback:\n"
    error_msg += traceback.format_exc()
    TaskDialog.Show("Detailed Error", error_msg)
```

---

## Support Checklist

Als je support nodig hebt, verzamel deze info:

```
□ Revit versie: _______
□ pyRevit versie: _______
□ Windows versie: _______
□ Script versie: _______
□ Error message (exact): _______
□ Screenshot van error
□ Gebruikte parameters:
  - Vakken: __ x __
  - Afmetingen: __ x __ mm
  - Materiaal: _______
□ Log files (pyRevit_run.log)
□ Stappen om te reproduceren
```

---

## Handige Commands

```python
# In RevitPythonShell, test Revit API:

# Check active document:
print(doc.Title)

# Check template path:
print(app.FamilyTemplateLocation)

# List all families in project:
collector = FilteredElementCollector(doc).OfClass(Family)
for fam in collector:
    print(fam.Name)

# Check current view:
print(doc.ActiveView.Name)
```

---

## Contact & Feedback

### Voor hulp:
1. Check deze troubleshooting guide
2. Check pyRevit forums: [https://discourse.pyrevitlabs.io/](https://discourse.pyrevitlabs.io/)
3. Check Revit API docs: [https://www.revitapidocs.com/](https://www.revitapidocs.com/)
4. Open issue in repository met bovenstaande support checklist

### Voor feedback & feature requests:
- Beschrijf gewenste functionaliteit
- Geef use case voorbeelden
- Include screenshots indien mogelijk

---

**Versie**: 1.0  
**Laatste update**: Januari 2026

Veel succes! 🛠️
