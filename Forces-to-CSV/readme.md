# Forces to CSV Button

## Inhalt
- Mechanical_Button_Compare_Loads_v01.py
- Ansys_Mechanical_Lagerreaktionen_Image.png

## Beschreibung
Dieses Python-Snippet lässt sich als Button in Mechanical importieren, um Reaktionskräfte automatisiert als CSV abzuspeichern.

## Installation
1. Pyhon-Datei und Bilddatei herunterladen.
2. Static-Structural in Workbench öffnen.
3. Im oberen Reiter auf "Automation/Manage" klicken - Es öffnet sich der Button Editor.
4. Oben auf "Click to import a button script".
5. Mechanical_Button_Compare_Loads_v01.py im Explorer auswählen und auf "öffnen" drücken.
6. Links auf das Bildsymbol im Button Editor drücken und im "Ansys_Mechanical_Lagerreaktionen_Image.png" im Explorer auswählen.
7. Oben Links auf "Click to save script for the button".

## Anwendung
1. Static Structural-Modell lösen.
2. Probe Results für die gewünschten Randbedingungen/Connections setzen.
3. "Evaluate all results" nicht vergessen ;).
4. Auf das Button-Smbol im Automation-Reiter drücken.
5. Die CSV-Datein finden sich nun unter: ".../Pfad-zum-Workbench-Projekt/user_files/".

## Was ist inkludiert?
Kräfte für:
- Forces,
- Remote Forces,
- Acceleration,
- Gravitation,
- Fixed Support,
- Displacement,
- Remote Displacement,
- Joints,
- Bonded Contacts.

Momente für:
- Joints,
- Bonded Contacts.

Sonstiges:
- Kräftebilanz,
- Einheitensystem wird berücksichtigt

## Was fehlt?
- Reaktionslasten von Balken, Federn, usw.
- Kräfte in Folge von Drücken als Lastvektor,
- Momente an Randbedingungen.

