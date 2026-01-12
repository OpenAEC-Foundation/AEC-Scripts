# 3BM Kozijn Family Creator voor Revit

Een professionele tool voor het parametrisch genereren van kozijn families in Revit, gebaseerd op 3BM specificaties.

## 🎯 Features

- ✅ **Vrij aantal vakken**: 1-10 horizontaal, 1-5 verticaal
- ✅ **3BM standaard afmetingen**: 80x90mm, 86x120mm kozijnen
- ✅ **Geen wall hosting**: Vrij plaatsbaar in 3D ruimte
- ✅ **Spouwlat & PE-folie**: Prefab bouwfase bescherming
- ✅ **Materiaal keuze**: Meranti, Grenen, Kunststof
- ✅ **HR++ en HR+++ glas**: Triple beglazing standaard
- ✅ **Nederlandse bouwpraktijk**: Conform NEN normen

## 🚀 Quick Start

### Installatie (pyRevit)
```bash
1. Installeer pyRevit: https://www.pyrevitlabs.io/
2. Kopieer kozijn_family_creator.py naar pyRevit folder
3. Herstart Revit
4. Klik op script in pyRevit toolbar
```

### Gebruik in 3 stappen
```
1. Run script → Configuratie scherm opent
2. Stel parameters in (vakken, afmetingen, materiaal)
3. Klik "Family Aanmaken" → Kies locatie → Klaar!
```

## 📐 Standaard 3BM Configuraties

### Standaard Raam 2-luiks
- **Afmetingen**: 1600 x 1200 mm
- **Kozijn**: 80 x 90 mm (meranti)
- **Vakken**: 2 horizontaal
- **Glas**: HR+++ Triple (4-14-4-15-4)

### Voordeur met bovenlicht
- **Afmetingen**: 1000 x 2400 mm
- **Kozijn**: 80 x 90 mm
- **Vakken**: 1H x 2V
- **Offset vloer**: 0 mm (peil)

### Puiraam 3-luiks
- **Afmetingen**: 2400 x 1500 mm
- **Kozijn**: 86 x 120 mm (thermo)
- **Vakken**: 3 horizontaal
- **Glas**: HR+++ Triple

## 🔧 Parameters Overzicht

### Vakindeling
| Parameter | Range | Standaard |
|-----------|-------|-----------|
| Vakken horizontaal | 1-10 | 2 |
| Vakken verticaal | 1-5 | 1 |

### Kozijnhout (mm)
| Parameter | Range | 3BM Standaard |
|-----------|-------|----------------|
| Randstijl breedte | 60-120 | 80 of 86 |
| Randstijl dikte | 60-160 | 90 of 120 |
| Tussenstijl breedte | 50-100 | 65 |

### Positie & Details
- **Offset vloer**: 0-500mm (standaard 100mm)
- **Sponning**: Binnen/Buiten (standaard Binnen)
- **Spouwlat**: 18 x 50 mm gelamineerd grenen
- **PE-folie**: 120 Mu bouwfasebescherming

## 📂 Bestanden

```
kozijn_family_creator.py          # Hoofdscript (run dit in Revit)
HANDLEIDING_Kozijn_Creator.md     # Uitgebreide gebruikersdocumentatie
README.md                          # Dit bestand
```

## 🎓 Uitgebreide Documentatie

Voor gedetailleerde informatie, zie **HANDLEIDING_Kozijn_Creator.md**:
- Complete stap-voor-stap instructies
- Technische details en berekeningen
- Veelgestelde vragen (FAQ)
- Voorbeelden en use cases
- Troubleshooting tips

## 💡 Gebruik Voorbeelden

### Voorbeeld 1: Standaard Woonkamer Raam
```python
Vakken: 3 x 1
Afmetingen: 2400 x 1500 mm
Kozijn: 80 x 90 mm
Materiaal: Meranti
Output: 3 vakken van 733 x 1340 mm (netto glas)
```

### Voorbeeld 2: Slaapkamer Raam 2-luiks
```python
Vakken: 2 x 1
Afmetingen: 1600 x 1200 mm
Kozijn: 80 x 90 mm
Materiaal: Kunststof
Output: 2 vakken van 720 x 1040 mm (netto glas)
```

### Voorbeeld 3: Achterdeur met zijlicht
```python
Vakken: 2 x 1
Afmetingen: 1600 x 2400 mm
Kozijn: 80 x 90 mm
Vak 1: Deur (900mm breed)
Vak 2: Glas (620mm breed)
```

## 🏗️ Workflow in Revit Project

### Stap 1: Family Aanmaken
```
1. Open Revit project
2. Run kozijn_family_creator.py
3. Configureer parameters
4. Sla .rfa bestand op
5. Laad in project (optioneel direct)
```

### Stap 2: Family Plaatsen
```
1. Architecture tab → Component
2. Selecteer je kozijn family
3. Plaats op gewenste locatie
4. Pas instance parameters aan indien nodig
```

### Stap 3: Parameters Aanpassen
```
Properties Panel:
- Totale_Breedte: Aanpasbaar
- Totale_Hoogte: Aanpasbaar  
- Offset_Vloer: Aanpasbaar
- Alle andere parameters instelbaar
```

## 🔍 Technische Specificaties

### Coordinate System
- **Origin**: Linker benedenhoek kozijn
- **X-as**: Breedte (horizontaal)
- **Y-as**: Hoogte (verticaal)
- **Z-as**: Diepte (uit de muur)

### Materiaal Mapping
- **Kozijnhout**: Generic Model - Wood
- **Glas**: Generic Model - Glass
- **Spouwlat**: Generic Model - Wood

### Glas Specificaties
- **HR++ Dubbel**: 5-15-4 mm (24mm totaal)
- **HR+++ Triple**: 4-14-4-15-4 mm (41mm totaal)
- **Sponning**: 15mm per zijde (standaard)

## 🛠️ Systeemvereisten

- **Revit**: 2019 of hoger aanbevolen
- **Python**: Via pyRevit of RevitPythonShell
- **OS**: Windows (Revit requirement)
- **Rechten**: Schrijftoegang voor .rfa bestanden

## 📋 Checklist voor Gebruik

- [ ] pyRevit geïnstalleerd
- [ ] Script gekopieerd naar juiste folder
- [ ] Revit herstart na installatie
- [ ] Project geopend waar family nodig is
- [ ] Besloten over:
  - [ ] Aantal vakken (H x V)
  - [ ] Totale afmetingen (B x H)
  - [ ] Kozijnhout afmetingen
  - [ ] Materiaal keuze
  - [ ] Spouwlat ja/nee
  - [ ] Offset vanaf vloer

## ⚠️ Belangrijke Opmerkingen

1. **Niet wall-hosted**: Deze families zijn Generic Models en hosten niet aan wanden. Dit is conform 3BM prefab aanpak.

2. **Family Editor**: Voor geavanceerde aanpassingen (deuren, draai/kiep, beslag) open de .rfa in Family Editor.

3. **Materialen**: Wijs materialen toe in project via Material overrides.

4. **Schedules**: Families verschijnen in Generic Model schedules. Pas Type Mark toe voor filtering.

5. **IFC Export**: Family's zijn IFC-compatible als Generic Building Element.

## 🐛 Bekende Issues & Workarounds

### Issue: Script start niet
**Oplossing**: Controleer of pyRevit correct is geïnstalleerd en Revit is herstart.

### Issue: Family niet zichtbaar na laden
**Oplossing**: Check visibility graphics (VG) → Generic Models moet aan staan.

### Issue: Parameters niet aanpasbaar
**Oplossing**: Parameters zijn Type parameters. Pas aan via Edit Type.

## 🚧 Toekomstige Ontwikkelingen

### Geplande Features (v2.0)
- [ ] Per-vak vulling configuratie UI
- [ ] Draai/kiep parameters per vak
- [ ] Beslag keuze bibliotheek
- [ ] Automatische U-waarde berekening
- [ ] Screen/hor integratie
- [ ] Detail sponning profielen
- [ ] Dorpel varianten (binnen/buiten/alu)
- [ ] Material presets (RAL kleuren)

### Onder Overweging
- [ ] Batch processing (meerdere kozijnen)
- [ ] Template library (veelgebruikte types)
- [ ] Integration met 3BM online configurator
- [ ] BIM object library koppeling
- [ ] Revit 2025 Extended Reality support

## 📞 Support & Feedback

### Bug Reports
Open een issue met:
- Revit versie
- Python/pyRevit versie
- Screenshot van error
- .rfa bestand (indien mogelijk)

### Feature Requests
Beschrijf gewenste functionaliteit en use case.

### Bijdragen
Pull requests zijn welkom! Check eerst open issues.

## 📄 Licentie

Dit script is ontwikkeld voor intern gebruik en kan vrij worden aangepast binnen je organisatie.

## 🙏 Credits

- **Gebaseerd op**: 3BM Kozijnen specificaties
- **Documentatie**: 3BM verwerkingsvoorschriften
- **Ontwikkeld voor**: Nederlandse bouwpraktijk
- **Compatible met**: NEN normen

## 📚 Referenties

- [3BM Kozijnen](https://www.3bm.nl/)
- [pyRevit Documentation](https://www.notion.so/pyRevit-bd907d6292ed4ce997c46e84b6ef67a0)
- [Revit API Documentation](https://www.revitapidocs.com/)

---

**Versie**: 1.0  
**Datum**: Januari 2026  
**Status**: Production Ready

**Veel succes met je 3BM kozijnen!** 🪟✨
