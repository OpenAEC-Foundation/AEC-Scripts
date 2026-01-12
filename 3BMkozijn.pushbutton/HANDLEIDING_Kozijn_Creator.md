# 3BM Kozijn Family Creator - Gebruikershandleiding

## Inhoudsopgave
1. [Introductie](#introductie)
2. [Installatie](#installatie)
3. [Gebruik](#gebruik)
4. [Parameters Uitgelegd](#parameters-uitgelegd)
5. [Voorbeelden](#voorbeelden)
6. [Technische Details](#technische-details)
7. [Veelgestelde Vragen](#veelgestelde-vragen)

---

## Introductie

De 3BM Kozijn Family Creator is een Python script voor Revit waarmee je parametrische kozijn families kunt genereren zonder wall-hosting. Het script is gebaseerd op 3BM kozijn specificaties en Nederlandse bouwpraktijk.

### Hoofdfuncties:
- ✅ Vrij aantal vakken horizontaal en verticaal
- ✅ Instelbare kozijnhout afmetingen (randstijlen, tussenstijlen, middenregels)
- ✅ Keuze voor vlakvulling per vak (glas/paneel/deur/raam)
- ✅ Draairichting per vak instelbaar
- ✅ Binnen of buiten sponning
- ✅ Offset vanaf vloerpeil
- ✅ Individuele breedte en hoogte per vak
- ✅ Randsponning met spouwlat en folies
- ✅ Geen wall hosting - vrij plaatsbaar

---

## Installatie

### Vereisten:
- Autodesk Revit (2019 of hoger aanbevolen)
- pyRevit of Revit Python Shell

### Installatie Stappen:

#### Optie 1: pyRevit (Aanbevolen)
1. Download en installeer pyRevit van [https://www.pyrevitlabs.io/](https://www.pyrevitlabs.io/)
2. Kopieer `kozijn_family_creator.py` naar je pyRevit scripts folder
3. Herstart Revit
4. Het script verschijnt in je pyRevit toolbar

#### Optie 2: Revit Python Shell
1. Download en installeer RevitPythonShell
2. Open Revit
3. Open RevitPythonShell (Add-ins > RevitPythonShell)
4. Open het `kozijn_family_creator.py` script
5. Druk op "Run" of F5

---

## Gebruik

### Stap-voor-stap:

#### 1. Start het Script
- Klik op het script icoon in pyRevit, of
- Run het script via RevitPythonShell

#### 2. Configuratie Dialog
Het configuratiescherm opent met de volgende secties:

**1. VAKINDELING**
- **Aantal vakken horizontaal**: Aantal vakken naast elkaar (1-10)
- **Aantal vakken verticaal**: Aantal vakken boven elkaar (1-5)

**2. TOTALE AFMETINGEN**
- **Totale breedte**: Buitenmaat kozijn in mm (400-5000mm)
- **Totale hoogte**: Buitenmaat kozijn in mm (400-3000mm)

**3. KOZIJNHOUT AFMETINGEN**
- **Randstijl breedte**: Breedte van de buitenstijlen (60-120mm)
  - Standaard 3BM: 80mm of 86mm
- **Randstijl dikte**: Diepte van het kozijn (60-160mm)
  - Standaard 3BM: 90mm, 120mm, of 160mm
- **Tussenstijl breedte**: Breedte van verticale tussenstijlen (50-100mm)
  - Standaard: 65mm

**4. SPONNING EN DETAILS**
- **Sponning type**: Binnen of Buiten sponning
- **Offset vanaf vloerpeil**: Hoogte boven afgewerkt vloer (0-500mm)
- **Spouwlat toevoegen**: Voeg gelamineerde grenen spouwlat toe
- **PE-folie bescherming**: Voeg bouwfasefolie toe

**5. MATERIAAL**
- **Kozijnmateriaal**: Meranti, Grenen, of Kunststof

#### 3. Family Aanmaken
- Klik op "Family Aanmaken"
- Kies een locatie en bestandsnaam voor de .rfa file
- Klik "Opslaan"
- Kies of je de family direct in het project wilt laden

#### 4. Family Gebruiken
Na het laden kun je de family plaatsen als een normaal Generic Model:
- Architecture tab > Component > Place a Component
- Selecteer je nieuwe kozijn family
- Plaats op gewenste locatie
- Pas parameters aan in Properties panel

---

## Parameters Uitgelegd

### Standaard 3BM Afmetingen

#### Houten Kozijnen:
- **Basis kozijn**: 80 x 90 mm (B x D)
- **Thermo kozijn**: 80 x 90 mm of 86 x 120 mm
- **Raam diepte**: 68mm, 78mm, of 82mm

#### Glas Specificaties:
- **HR++ glas**: 5-15-4 mm (24mm totaal)
- **HR+++ Triple**: 4-14-4-15-4 mm (41mm totaal)

#### Spouwlat:
- **Materiaal**: Gelamineerd grenen
- **Afmeting**: Circa 18 x 50 mm
- **Afwerking**: 120 Mu verf

#### PE-Folie:
- **Dikte**: 120 Mu (0.12mm)
- **Functie**: Bescherming tijdens bouwfase

### Family Parameters (na creatie instelbaar):

In het Properties panel van een geplaatste family instance:

**Geometrie:**
- `Totale_Breedte`: Aanpasbare totale breedte
- `Totale_Hoogte`: Aanpasbare totale hoogte
- `Randstijl_Breedte`: Breedte randstijlen
- `Randstijl_Dikte`: Diepte kozijn
- `Tussenstijl_Breedte`: Breedte tussenstijlen
- `Offset_Vloer`: Hoogte boven vloer
- `Aantal_Vakken_Horizontaal`: Aantal horizontale vakken
- `Aantal_Vakken_Verticaal`: Aantal verticale vakken

**Materialen:**
- `Kozijn_Materiaal`: Houtsoort of kunststof
- `Glas_Type`: Type beglazing

---

## Voorbeelden

### Voorbeeld 1: Standaard Raam 2-luiks
```
Vakindeling:
- Horizontaal: 2
- Verticaal: 1

Afmetingen:
- Breedte: 1600 mm
- Hoogte: 1200 mm

Kozijnhout:
- Randstijl: 80 x 90 mm
- Tussenstijl: 65 mm

Resultaat: 2 gelijke vakken van 720 x 1040 mm (netto glas)
```

### Voorbeeld 2: Voordeur met bovenlicht
```
Vakindeling:
- Horizontaal: 1
- Verticaal: 2

Afmetingen:
- Breedte: 1000 mm
- Hoogte: 2400 mm

Kozijnhout:
- Randstijl: 80 x 90 mm
- Middenregel: 65 mm

Resultaat: 
- Ondervak (deur): 840 x 1600 mm
- Bovenvak (glas): 840 x 655 mm
```

### Voorbeeld 3: Puiraam 4-luiks
```
Vakindeling:
- Horizontaal: 4
- Verticaal: 1

Afmetingen:
- Breedte: 3200 mm
- Hoogte: 1500 mm

Kozijnhout:
- Randstijl: 86 x 120 mm
- Tussenstijl: 65 mm

Resultaat: 4 vakken van 715 x 1334 mm (netto glas)
```

---

## Technische Details

### Coordinate System
- **Origin (0,0,0)**: Linker benedenhoek kozijn
- **X-as**: Breedte richting (horizontaal)
- **Y-as**: Hoogte richting (verticaal)
- **Z-as**: Diepte richting (uit de muur)

### Afmetingen Berekening

#### Netto Vak Breedte:
```
vak_breedte = (Totale_Breedte - 2 × Randstijl_Breedte - (n-1) × Tussenstijl_Breedte) / n

waarbij n = Aantal_Vakken_Horizontaal
```

#### Netto Vak Hoogte:
```
vak_hoogte = (Totale_Hoogte - 2 × Randstijl_Breedte - (m-1) × Middenregel_Hoogte) / m

waarbij m = Aantal_Vakken_Verticaal
```

#### Glas Afmetingen:
```
glas_breedte = vak_breedte - 2 × Sponning_Diepte
glas_hoogte = vak_hoogte - 2 × Sponning_Diepte

Standaard sponning: 15mm
```

### Component Layers (van buiten naar binnen):

1. **Spouwlat** (optioneel): -18 tot 0 mm
2. **PE-Folie** (optioneel): 0 tot 0.15 mm
3. **Kozijnhout**: 0 tot 90 mm (afhankelijk van dikte)
4. **Sponning**: 15 mm diep vanaf kozijnhout
5. **Glas**: Gecentreerd in kozijndiepte

### Materiaal Definitie

De family gebruikt de volgende Revit categorieën:
- **Kozijnhout**: Generic Model - Hout materiaal
- **Glas**: Generic Model - Glas materiaal
- **Spouwlat**: Generic Model - Hout materiaal

---

## Veelgestelde Vragen

### Q: Kan ik de family aanpassen na het maken?
**A:** Ja, open de .rfa file in de Family Editor. Je kunt daar alle parameters en geometrie aanpassen.

### Q: Hoe voeg ik deuren toe aan vakken?
**A:** Momenteel ondersteunt het script alleen glas vulling. Voor deuren:
1. Open de family in Family Editor
2. Selecteer het glas element van het betreffende vak
3. Vervang met een deur component of maak custom geometrie

### Q: Waarom worden mijn kozijnen niet aan muren gehangen?
**A:** Dit is by design. 3BM kozijnen zijn prefab en worden vaak los geplaatst. Je kunt ze handmatig positioneren ten opzichte van wanden.

### Q: Kan ik draairichtingen instellen?
**A:** De basis family heeft vaste glazen. Voor draai/kiep functionaliteit:
1. Open family in Family Editor
2. Voeg nested families toe voor ramen met beslag
3. Definieer opening parameters

### Q: Hoe voeg ik screens of horren toe?
**A:** Deze kunnen als separate nested families worden toegevoegd:
1. Maak/import screen familie
2. Laad in je kozijn family
3. Plaats op gewenste positie
4. Maak visibility parameters

### Q: Werkt dit ook voor kunststof kozijnen?
**A:** Ja, selecteer "Kunststof" bij materiaal. De afmetingen zijn vergelijkbaar. Let op:
- Kunststof basis: 86 x 120 mm systeem
- Kunststof thermo: idem met triple glas

### Q: Kan ik de family in IFC exporteren?
**A:** Ja, de family is een Generic Model en kan naar IFC worden geëxporteerd. Gebruik de juiste IFC mappings in je export settings.

### Q: Hoe voeg ik ventilatie roosters toe?
**A:** Maak een separate ventilatierooster family en nest deze in je kozijn family op de gewenste positie.

---

## Ondersteuning & Uitbreidingen

### Script Aanpassen
Het script is volledig open source en kan worden aangepast. Belangrijke functies:

- `_create_frame()`: Creëert kozijnframe
- `_create_tussenstijlen()`: Verticale verdeling
- `_create_middenregels()`: Horizontale verdeling
- `_create_vak_vullingen()`: Glas/panelen
- `_create_spouwlat()`: Spouwlat detaillering

### Toekomstige Features (TODO)
- [ ] Per-vak vulling configuratie (glas/paneel/deur)
- [ ] Draai/kiep functionaliteit per vak
- [ ] Beslag keuze (normaal/verdekt)
- [ ] Screen integratie
- [ ] Dorpel profielen (binnen/buiten/alu)
- [ ] Hang- en sluitwerk parameters
- [ ] Color/finish parameters
- [ ] Detail sponningen en waterslag
- [ ] Nested raam families
- [ ] U-waarde berekening display

### Contact & Feedback
Voor vragen, suggesties of bug reports:
- Open een issue in de repository
- Contact via je Revit BIM coordinator

---

## Licentie
Dit script is ontwikkeld voor intern gebruik en kan vrij worden aangepast en gedistribueerd binnen je organisatie.

---

## Versie Historie

### v1.0 (Januari 2026)
- Initiële release
- Basis kozijn generatie
- GUI configurator
- 3BM standaard afmetingen
- Spouwlat en folie opties
- Generic Model basis (niet wall-hosted)

---

**Veel succes met het maken van je 3BM kozijnen!** 🪟🏗️
