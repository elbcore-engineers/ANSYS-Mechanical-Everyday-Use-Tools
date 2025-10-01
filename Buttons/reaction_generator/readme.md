# Reaction Generator

## Inhalt
- reaction_generator_v0-1.py (aktuell)
- reaction_generator.png

## Beschreibung
Dieses Python-Snippet lässt sich als Button in Mechanical importieren, um Lastreaktionen (Kräfte und Momente) für Connections (Contacts, Joints, Beams) und Boundary Conditions (Displacements, Remote Displacements, Fixed Supports) zu erzeugen.
Es werden keine Dopplungen erzeugt.

## Installation
1. Pyhon-Datei und Bilddatei herunterladen.
2. Static-Structural in Workbench öffnen.
3. Im oberen Reiter auf "Automation/Manage" klicken - Es öffnet sich der Button Editor.
4. Oben auf "Click to import a button script".
5. reaction_generator_v0-1.py im Explorer auswählen und auf "öffnen" drücken.
6. Links auf das Bildsymbol im Button Editor drücken und "reaction_generator.png" im Explorer auswählen.
7. Oben Links auf "Click to save script for the button".

## Anwendung
1. Modell lösen.
3. Auf das entsprechende Button-Symbol im Automation-Reiter drücken.
4. Auswahl:
   - "Vorhandene Reaktionen überschreiben?": Vorhandene Reaktionen werden namentlich an die neu generierten Reaktionen angepasst.
   - "Generierte Reaktionen gruppieren?": Reaktionen werden nach ihrem Typ gruppiert.

