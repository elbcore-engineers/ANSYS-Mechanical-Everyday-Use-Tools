# Reaction Master

## Inhalt
- reaction_master_v02.py
- reaction_master_v01.py
- reaction_master_image.png

## Beschreibung
Dieses Python-Snippet lässt sich als Button in Mechanical importieren, um Reaktionskräfte und Schraubenasulastungen automatisiert als CSV abzuspeichern.

## Installation
1. Pyhon-Datei und Bilddatei herunterladen.
2. Static-Structural in Workbench öffnen.
3. Im oberen Reiter auf "Automation/Manage" klicken - Es öffnet sich der Button Editor.
4. Oben auf "Click to import a button script".
5. reaction_master_v02.py im Explorer auswählen und auf "öffnen" drücken.
6. Links auf das Bildsymbol im Button Editor drücken und im "reaction_master_image.png" im Explorer auswählen.
7. Oben Links auf "Click to save script for the button".

## Anwendung
### Allgemein
1. Static Structural-Modell lösen.
2. Probe Results für die gewünschten Randbedingungen/Connections setzen.
3. "Evaluate all results" nicht vergessen ;).
4. Auf das Button-Smbol im Automation-Reiter drücken.
5. Die CSV-Datein finden sich nun unter: ".../Pfad-zum-Workbench-Projekt/user_files/".

### Schraubenasuwertung
1. In Connection: Beam oder Joint-Objekt erstellen.
2. Im Namen des Objekts muss irgendwo der Nenndurchmesser mit dem Präfix "M" auftauchen, z. B. "M16 Schraube links", um die Connection als Schraube mit dem Durchmesser 16 zu erkennen.
3. Standardmäßig wird von einer Festigkeit von 640 MPa ausgegangen.
4. (Optional) Im Namen irgendwo die Festigkeitsklasse ergänzen, z. B. "4.6", umd eine Streckgrenze von 240 MPa zu berücksichtigen.
5. Die Schraubenauslastung findet sich in ".../Pfad-zum-Workbench-Projekt/user_files/10_Schraubenauslastungen.csv"

Hinweis: 
- Die Annahme ist, dass die lokale x-Richtung der Joints stets axial ausgerichtet ist (Normalfall),
- Normalspannungen und Biegespannungen werden betragsmäßig überlagert (es kommt zu keiner Abminderung bei unterschiedlichen Vorzeichen oder Entlastungen bei Druckspannungen)
   
## Was ist in der neuesten Version inkludiert?
Kräfte für:
- Forces,
- Remote Forces,
- Acceleration,
- Gravitation,
- Fixed Support,
- Displacement,
- Remote Displacement,
- Joints,
- Beams,
- Bonded Contacts.

Momente für:
- Joints,
- Beams,
- Bonded Contacts.

Sonstiges:
- Schraubenauswertung
- Kräftebilanz,
- Einheitensystem wird berücksichtigt

## Was fehlt?
- Reaktionslasten von Federn,..., usw.
- Kräfte in Folge von Drücken als Lastvektor,
- Momente an globalen Randbedingungen.

