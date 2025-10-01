# Script
# Einheitensystem: s,t,mm,N

# ----------------
# Importiere Module
# ----------------
import os
import sys
import wbjn
import context_menu
import subprocess
import csv
import math
from datetime import datetime

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Form, Button, Label, DialogResult, FolderBrowserDialog,
    FormStartPosition, TextBox
)
from System.Drawing import Point, Size

# ----------------
# Einstellungen
# ----------------
seperator = ","
nameSumLoad = "Gesamtlast"
nameSumBc = "Gesamtreaktion"
nameSumDif = "Differenz"
nameSumJointForce = ""
nameSumJointMoment = ""
nameSumContactForce = ""
nameSumContactMoment = ""
nameSumBeamForce = ""
nameSumBeamMoment = ""

# ----------------
# Funktionen
# ----------------

def makeFolder(path):
    """Check and make folder"""
    if not os.path.exists(path):
        os.makedirs(path)
    return(path)

def check_for_stop(directory):
    """
    Checks if a file named "stop", "stop.txt","kill","kill.txt","exit" or "exit.txt" exists in the given directory. 
    If found, the program exits with a message.
    """
    stop_files = ["stop", "stop.txt","kill","kill.txt","exit","exit.txt"]
    for stop_file in stop_files:
        stop_path = os.path.join(directory, stop_file)
        if os.path.isfile(stop_path):
            print(str(stop_file)+" found in directory" + str(directory)+". Stopping execution.")
            sys.exit(0)

def extractValue(entry, seperator):
    """
    Extract the value "100" from an input with the format "100 [MPa]"
    """
    entryStr = str(entry)
    if " [" in entryStr:
        return entryStr.split(" [")[0].replace(".", seperator)
    return ""

def extractUnit(entry):
    """
    Extract the unit "[MPA]" from an input with the format "100 [MPa]"
    """
    entryStr = str(entry)
    bracketIndex = entryStr.find('[')
    if bracketIndex != -1:
        return entryStr[bracketIndex:]
    return ""

def isEmpty(input):
    """Check if input is empty"""
    if input in [[],[[]], None, False]:
        return True
    else:
        return False


def changeSeperator(entry, decimalSeperator):
    """Change seperator"""
    entry = str(entry)
    if "." in entry and decimalSeperator == ",":
        return entry.replace(".", decimalSeperator)
    elif "," in entry and decimalSeperator == ".":
        return entry.replace(",", decimalSeperator)
    else:
        return(entry)

def printTable(table):
    for line in table:
        print(line)

def adjustName(text):
    text = text.replace("=",":")
    text = text.replace("   "," ")
    text = text.replace("  "," ")
    text = text.replace("ä","ae")
    text = text.replace("ö","oe")
    text = text.replace("ü","ue")
    text = text.replace("Ä","AE")
    text = text.replace("Ö","OE")
    text = text.replace("Ü","UE")
    text = text.replace("ß","ss")
    text = text.replace("²","^2")
    text = text.replace("³","^3")
    text = text.replace("·","*")
    text = text.replace("°","Degrees")
    return(text)

def adjustPath(text):
    text = text.replace("=","")
    text = text.replace(": ","_")
    text = text.replace("   "," ")
    text = text.replace("  "," ")
    text = text.replace(" ","-")
    text = text.replace("/","-")
    text = text.replace(":","-")
    text = text.replace("ä","ae")
    text = text.replace("ö","oe")
    text = text.replace("ü","ue")
    text = text.replace("Ä","AE")
    text = text.replace("Ö","OE")
    text = text.replace("Ü","UE")
    text = text.replace("ß","ss")
    text = text.replace("²","^2")
    text = text.replace("³","^3")
    text = text.replace("·","*")
    text = text.replace("°","Degrees")
    return(text)
    
    
def readTabularData(object_to_read, seperator):
    """Read tabular data of an object and return a 2D list"""
    try:
        object_to_read.Activate()
        Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
        tabularData = Pane.ControlUnknown
        
        table_values = []
        
        for row in range(1, tabularData.RowsCount + 1):
            line = []
            for column in range(2, tabularData.ColumnsCount + 1):
                cell = tabularData.cell(row, column)
                if cell is None or cell.Text is None:
                    line.append("")
                else:
                    line.append(changeSeperator(adjustName(cell.Text), seperator))
            table_values.append(line)

        return(table_values)
    except:
        print("Unsuccessfully read tabular data")


def createAccelerationForce(mass, accelerationTable):
    """
    Create a force table (F = m * a) from mass and acceleration table.
    """
    try:
        mass = float(mass)
    except:
        print("Invalid mass value.")
        return []

    forceTable = []
    
    header = accelerationTable[0]

    # Manuell Indexe finden
    time_index = -1
    x_index = -1
    y_index = -1
    z_index = -1
    for i in range(len(header)):
        if header[i].startswith("Time"):
            time_index = i
        elif header[i].startswith("X"):
            x_index = i
        elif header[i].startswith("Y"):
            y_index = i
        elif header[i].startswith("Z"):
            z_index = i

    if time_index == -1 or x_index == -1 or y_index == -1 or z_index == -1:
        print("Could not find required columns in header.")
        return []

    forceTable.append(["Time [s]", "X [N]", "Y [N]", "Z [N]"])

    for row in accelerationTable[1:]:
        if len(row) <= max(x_index, y_index, z_index):
            continue
        try:
            time = float(row[time_index].replace(",", "."))
            ax = float(adjustName(row[x_index].replace(",", "."))) * 0.001
            ay = float(adjustName(row[y_index].replace(",", "."))) * 0.001
            az = float(adjustName(row[z_index].replace(",", "."))) * 0.001

            fx = mass * ax
            fy = mass * ay
            fz = mass * az

            forceTable.append([str(time), str(fx), str(fy), str(fz)])
        except:
            continue

    return forceTable


#def isEmpty(obj):
#    return not obj or len(obj) == 0

def combineTables(listOfTables, nameList, nameSum, data_type):
    """
    Kombiniert Tabellen nebeneinander:
    - Erste Spalte: Zeit (aus erster Tabelle)
    - Dann: X/Y/Z- oder Beam-Werte je Tabelle
    - Optional: Summenspalten (wenn nameSum gesetzt ist)

    Unterstützte data_type-Werte:
        "force"       → ["X [N]", "Y [N]", "Z [N]"]
        "moment"      → ["X [Nmm]", "Y [Nmm]", "Z [Nmm]"]
        "beamForce"   → ["Axial [N]", "Quer I [N]", "Quer J [N]"]
        "beamMoment"  → ["Torsion [Nmm]", "Biegung I [Nmm]", "Biegung J [Nmm]"]
    """

    if len(listOfTables) != len(nameList):
        print("Fehler: Tabellen- und Namensanzahl stimmen nicht überein.")
        return []

    if isEmpty(nameList):
        return []

    # Achsen-/Spaltenüberschriften nach Typ definieren
    if data_type == "force":
        headers = ["X [N]", "Y [N]", "Z [N]"]
    elif data_type == "moment":
        headers = ["X [Nmm]", "Y [Nmm]", "Z [Nmm]"]
    elif data_type == "beamForce":
        headers = ["Axial [N]", "Quer I [N]", "Quer J [N]"]
    elif data_type == "beamMoment":
        headers = ["Torsion [Nmm]", "Biegung I [Nmm]", "Biegung J [Nmm]"]
    else:
        print("Unbekannter data_type: " + str(data_type))
        return []

    combinedTable = []
    maxLen = 0
    for table in listOfTables:
        if len(table) > maxLen:
            maxLen = len(table)

    # Tabellen auf gleiche Länge bringen
    for i in range(len(listOfTables)):
        while len(listOfTables[i]) < maxLen:
            listOfTables[i].append([""] * (len(headers) + 1))

    # Kopfzeile 1: Gruppennamen
    headerLine1 = ["", ""]
    for name in nameList:
        headerLine1 += [name] + [""] * (len(headers) - 1)
    if nameSum:
        headerLine1 += [nameSum] + [""] * (len(headers) - 1)
    combinedTable.append(headerLine1)

    # Kopfzeile 2: Achsen/Bezeichnungen
    headerLine2 = ["Time [s]"]
    for _ in nameList:
        headerLine2 += headers
    if nameSum:
        headerLine2 += headers
    combinedTable.append(headerLine2)

    # Datenzeilen aufbauen
    for rowIndex in range(1, maxLen):
        row = []

        # Zeitwert aus erster Tabelle
        try:
            time = listOfTables[0][rowIndex][0]
        except:
            time = ""
        row.append(time)

        # Summe initialisieren
        sums = [0.0] * len(headers)

        for table in listOfTables:
            values = []
            for i in range(len(headers)):
                try:
                    valStr = table[rowIndex][i + 1]
                    val = float(str(valStr).replace(",", "."))
                except:
                    val = 0.0
                values.append(val)
                sums[i] += val

            # Werte zur Zeile hinzufügen
            for v in values:
                row.append(str(v))

        # Summenzeile anhängen
        if nameSum:
            for s in sums:
                row.append(str(s))

        combinedTable.append(row)

    return combinedTable


def cleanBCTable(rawTable, data_type):
    """
    Wandelt ein Raw-Table mit möglichen X, Y, Z und Total-Spalten in ein sauberes Format um.
    Gibt eine Tabelle im Format zurück:
        - "force"  → ["Time [s]", "X [N]", "Y [N]", "Z [N]"]
        - "moment" → ["Time [s]", "X [Nmm]", "Y [Nmm]", "Z [Nmm]"]

    Alle fehlenden Richtungen werden mit 0 ersetzt.
    Die echte Summenspalte (z.B. ".. (Total) [N]") wird ignoriert.
    """

    if not rawTable or len(rawTable) < 2:
        return []

    # Einheit je nach Datentyp
    unit = "Nmm" if str(data_type).lower() == "moment" else "N"

    header = rawTable[0]

    time_index = -1
    x_index = -1
    y_index = -1
    z_index = -1

    for i, col in enumerate(header):
        col_low = str(col).lower().strip()

        # Echte Summenspalten erkennen → ignorieren
        if col_low == "total" or col_low.endswith("(total) [n]") or col_low.endswith("(total) [nmm]"):
            continue

        # Zeitspalte
        if "time" in col_low and time_index == -1:
            time_index = i

        # X, Y, Z erkennen (auch wenn "total" im Namen vorkommt → nicht ignorieren)
        elif "x" in col_low and x_index == -1:
            x_index = i
        elif "y" in col_low and y_index == -1:
            y_index = i
        elif "z" in col_low and z_index == -1:
            z_index = i

    # Neue Tabelle initialisieren
    cleanedTable = []
    cleanedTable.append(["Time [s]", "X [%s]" % unit, "Y [%s]" % unit, "Z [%s]" % unit])

    for row in rawTable[1:]:
        newRow = []

        # Zeit
        try:
            timeVal = row[time_index] if time_index != -1 else ""
            newRow.append(float(str(timeVal).replace(",", ".").replace("s", "").strip()))
        except:
            newRow.append(0.0)

        # X, Y, Z – mit Fallback 0.0
        for idx in [x_index, y_index, z_index]:
            try:
                if idx != -1:
                    val = str(row[idx]).replace(",", ".")
                    newRow.append(float(val))
                else:
                    newRow.append(0.0)
            except:
                newRow.append(0.0)

        cleanedTable.append(newRow)

    return cleanedTable



def splitBeamTable(rawTable):
    """
    Zerlegt eine Beam-Ergebnistabelle in zwei Tabellen:
    - Kräfte:    ["Time [s]", "Axial [N]", "Quer I [N]", "Quer J [N]"]
    - Momente:   ["Time [s]", "Torsion [Nmm]", "Biegung I [Nmm]", "Biegung J [Nmm]"]

    Funktioniert mit typischen Beam Probe-Ergebnissen aus Ansys Mechanical.
    """
    if not rawTable or len(rawTable) < 2:
        return [], []

    header = rawTable[0]
    
    # Indexe für relevante Spalten
    time_index = -1
    axial_index = -1
    shearI_index = -1
    shearJ_index = -1
    torque_index = -1
    momentI_index = -1
    momentJ_index = -1

    for i in range(len(header)):
        col = str(header[i]).lower()
        if "time" in col:
            time_index = i
        elif "axial" in col:
            axial_index = i
        elif "shear" in col and "at i" in col:
            shearI_index = i
        elif "shear" in col and "at j" in col:
            shearJ_index = i
        elif "torque" in col:
            torque_index = i
        elif "moment" in col and "at i" in col:
            momentI_index = i
        elif "moment" in col and "at j" in col:
            momentJ_index = i

    # Tabellenköpfe definieren
    forceTable = []
    momentTable = []

    forceTable.append(["Time [s]", "Axial [N]", "Quer I [N]", "Quer J [N]"])
    momentTable.append(["Time [s]", "Torsion [Nmm]", "Biegung I [Nmm]", "Biegung J [Nmm]"])

    for i in range(1, len(rawTable)):
        row = rawTable[i]

        # Zeitwert
        try:
            time = row[time_index] if time_index != -1 else ""
        except:
            time = ""

        # Kraftwerte
        try:
            axial = float(str(row[axial_index]).replace(",", ".")) if axial_index != -1 else 0.0
        except:
            axial = 0.0

        try:
            shearI = float(str(row[shearI_index]).replace(",", ".")) if shearI_index != -1 else 0.0
        except:
            shearI = 0.0

        try:
            shearJ = float(str(row[shearJ_index]).replace(",", ".")) if shearJ_index != -1 else 0.0
        except:
            shearJ = 0.0

        forceTable.append([time, str(axial), str(shearI), str(shearJ)])

        # Momentenwerte
        try:
            torque = float(str(row[torque_index]).replace(",", ".")) if torque_index != -1 else 0.0
        except:
            torque = 0.0

        try:
            momentI = float(str(row[momentI_index]).replace(",", ".")) if momentI_index != -1 else 0.0
        except:
            momentI = 0.0

        try:
            momentJ = float(str(row[momentJ_index]).replace(",", ".")) if momentJ_index != -1 else 0.0
        except:
            momentJ = 0.0

        momentTable.append([time, str(torque), str(momentI), str(momentJ)])

    return forceTable, momentTable


def convertLoadForceTable(rawTable):
    """
    Problem: Lasten beinhalten noch Spalte mit Steps. Das macht Probleme bei der Kombination von Tabellen.
    Funktion: Konvertiert eine Tabelle der Form:
    ['Steps', 'Time [s]', 'X [N]', 'Y [N]', 'Z [N]']
    in das Format:
    ['Time [s]', 'X [N]', 'Y [N]', 'Z [N]']
    """
    if not rawTable or len(rawTable) < 2:
        return []

    newTable = []
    header = ["Time [s]", "X [N]", "Y [N]", "Z [N]"]
    newTable.append(header)

    for row in rawTable[1:]:  # Datenzeilen
        try:
            time = row[1] if len(row) > 1 else ""
            fx = row[2] if len(row) > 2 else "0.0"
            fy = row[3] if len(row) > 3 else "0.0"
            fz = row[4] if len(row) > 4 else "0.0"
        except:
            time, fx, fy, fz = "", "0.0", "0.0", "0.0"

        newTable.append([time, fx, fy, fz])

    return newTable


def combineResultTable(loadTable, bcTable, nameloads="Gesamtlast", nameBCs="Gesamtreaktion", nameDiff="Differenz"):
    """
    Erstellt eine Tabelle mit Zeit, Summen der Lasten, Reaktionen und Differenz (Last - Reaktion).
    
    - Annahme: Letzte 3 Spalten jeder Tabelle sind X/Y/Z-Gesamtsummen
    - Annahme: Erste Spalte ist "Time"
    """

    resultTable = []

    # Kopfzeile mit "verschmolzenen" Überschriften: nur einmal Name, dann 2 leere Zellen
    header = [
        "","",nameloads, "","",nameBCs, "","",nameDiff, "", 
    ]
    # Darunter eine zweite Headerzeile für X, Y, Z-Spalten
    subheader = [
        "Time [S]",  # Zeitspalte leer
        "X [N]", "Y [N]", "Z [N]",
        "X [N]", "Y [N]", "Z [N]",
        "X [N]", "Y [N]", "Z [N]"
    ]

    resultTable.append(header)
    resultTable.append(subheader)

    # Berechne Zeilenanzahl anhand kleinster Tabelle (ab drittem Datensatz bei Lasten)
    numRows = min(len(loadTable) - 3, len(bcTable) - 2)

    for i in range(numRows):
        row = []
        loadRow = loadTable[i + 3]  # Skip 0er-Zeitwert (Header + Zeile 0)
        bcRow = bcTable[i + 2]      # Header + echte Daten

        try:
            time = loadRow[0]
            row.append(time)

            # Load-Werte
            fx_load = float(loadRow[-3].replace(",", "."))
            fy_load = float(loadRow[-2].replace(",", "."))
            fz_load = float(loadRow[-1].replace(",", "."))

            # Reaction-Werte
            fx_bc = float(bcRow[-3].replace(",", "."))
            fy_bc = float(bcRow[-2].replace(",", "."))
            fz_bc = float(bcRow[-1].replace(",", "."))

            # Differenz
            fx_diff = fx_load + fx_bc
            fy_diff = fy_load + fy_bc
            fz_diff = fz_load + fz_bc

            row += [
                str(fx_load), str(fy_load), str(fz_load),
                str(fx_bc), str(fy_bc), str(fz_bc),
                str(fx_diff), str(fy_diff), str(fz_diff),
            ]
        except:
            row += [""] * 9

        resultTable.append(row)

    return resultTable


def collectAllConnections(container, contactObjectList, jointObjectList, beamObjectList, connectionGroupTypes):
    """
    Sammelt rekursiv alle Connections aus einer beliebig
    verschachtelten Containerstruktur.

    container      : aktuelles Connection-Objekt oder -Gruppe
    connectionList : Liste zum Auffüllen
    """
    # Manche Container haben Children, andere sind selbst schon Connections
    print("Read connection: " + container.Name)

    if hasattr(container, "Children"):

        for item in container.Children:
            itemType = item.GetType()
            print("\t Type:" + str(itemType))
            if itemType in connectionGroupTypes:
                # rekursiv in Untergruppe
                collectAllConnections(item, contactObjectList, jointObjectList, beamObjectList, connectionGroupTypes)

            elif itemType == Ansys.ACT.Automation.Mechanical.Connections.ContactRegion:
                print("\t -> assigned to contact type")
                contactObjectList.append(item)

            elif itemType == Ansys.ACT.Automation.Mechanical.Connections.Joint:
                print("\t -> assigned to joint type")
                jointObjectList.append(item)

            elif itemType == Ansys.ACT.Automation.Mechanical.Connections.Beam:
                print("\t -> assigned to beam type")
                beamObjectList.append(item)

    else:
        # kein Container, sondern direkte Connection

        if itemType == Ansys.ACT.Automation.Mechanical.Connections.ContactRegion:
            print("\t -> ssigned to contact type")
            contactObjectList.append(item)

        elif itemType == Ansys.ACT.Automation.Mechanical.Connections.Joint:
            print("\t -> assigned to joint type")
            jointObjectList.append(item)

        elif itemType == Ansys.ACT.Automation.Mechanical.Connections.Beam:
            print("\t -> assigned to beam type")
            beamObjectList.append(item)

def collectAllBoundaries(container, bouboundaryObjectList, boundaryTypes):
    # Manche Container haben Children, andere sind selbst schon boundaries
    print("Read boundary: " + container.Name)

    if item.GetType() in boundaryTypes:
        bouboundaryObjectList.append(item)


def getBoltDataFromName(text): 
    """
    Extrahiert Nenndurchmesser [mm] und Streckgrenze [MPa] 
    aus dem Namen einer Schrauben-Connection.

    Input:
      text : string (z.B. 'Bolt M12 8.8')

    Output:
      (d, A_sp, fp, fm, av) : Durchmesser [mm], Spannungsquerscnitt [mm²], Streckgrenze [MPa], Zugfestigkeit [MPa], Abscherkoeefizient
      False  : wenn kein passender Eintrag gefunden
    """
    
    text = adjustName(text)

    d = 0.1  # Standardwert Nenndurchmesser
    A_sp = 0.1 # Standardwert Spannungsquerschnitt
    fp = 640  # Standardwert Streckgrenze
    fm = 800  # Standardwert Zugfestigkeit
    av = 0.5 # Standardwert Abscherkoeffizient

    bolt_sizes = {
        "M64": 64.0, "M56": 56.0, "M48": 48.0, "M42": 42.0, "M36": 36.0,
        "M30": 30.0, "M27": 27.0, "M24": 24.0, "M22": 22.0, "M20": 20.0,
        "M18": 18.0, "M16": 16.0, "M14": 14.0, "M12": 12.0, "M10": 10.0,
        "M8": 8.0, "M6": 6.0, "M5": 5.0, "M4": 4.0, "M3": 3.0
    }

    stress_areas = {
        "M64": 2675.97, "M56": 2030.02, "M48": 1473.15, "M42": 1120.91,
        "M36": 816.72, "M30": 560.59, "M27": 459.41, "M24": 352.50,
        "M22": 303.40, "M20": 244.79, "M18": 192.47, "M16": 156.67,
        "M14": 115.44, "M12": 84.27, "M10": 57.99, "M8": 36.61,
        "M6": 20.12, "M5": 14.18, "M4": 8.78, "M3": 5.03
    }

    yield_strengths = {
        "12.9": 940, "10.9": 720, "9.8": 660, "8.8": 640, "6.8": 480,
        "5.8": 420, "5.6": 300, "4.8": 340, "4.6": 240, "4.4": 150, "3.6": 180
    }

    tensile_strenghts = {
        "12.9": 1200, "10.9": 1000, "9.8": 900, "8.8": 800, "6.8": 600,
        "5.8": 500, "5.6": 500, "4.8": 400, "4.6": 400, "4.4": 400, "3.6": 300
    }

    shear_coefficient = {
        "12.9": 0.5, "10.9": 0.5, "9.8": 0.5, "8.8": 0.6, "6.8": 0.5,
        "5.8": 0.5, "5.6": 0.6, "4.8": 0.6, "4.6": 0.6, "4.4": 0.5, "3.6": 0.5
    }

    for key in bolt_sizes:
        if key in text:
            d = bolt_sizes[key]
            break

    for key in stress_areas:
        if key in text:
            A_sp = stress_areas[key]
            break

    for key in yield_strengths:
        if key in text:
            fp = yield_strengths[key]
            break

    for key in tensile_strenghts:
        if key in text:
            fm = tensile_strenghts[key]
            break

    for key in shear_coefficient:
        if key in text:
            av = shear_coefficient[key]
            break

    if d == 0.1:
        return False
    else:
        return d, A_sp, fp, fm, av
    

def tryGetData(name, table):
    """
    Liest eine Spalte aus einer Tabellendatenstruktur.
    Falls Werte nicht konvertierbar sind -> 999999999 einsetzen.
    
    name  : Spaltenname (string)
    table : 2D-Liste (erste Zeile = Header)

    Output: Liste mit floats (eine pro Zeile)
    """

    for i, col in enumerate(table[0]):
        if col.strip() == name:
            values = []
            for row in table[1:]:
                try:
                    values.append(float(row[i]))
                except:
                    values.append(999999999)
            return values
    return [0.0] * (len(table) - 1)


def getArea(d):
    """
    Querschnittsfläche einer Schraube [mm²].
    d : Nenndurchmesser [mm]
    """

    return math.pi * (d ** 2) / 4.0

def getWb(d):
    """
    Widerstandsmoment für Biegung [mm³].
    d : Nenndurchmesser [mm]
    """

    return math.pi * (d ** 3) / 32.0


def getWt(d):
    """
    Widerstandsmoment für Torsion [mm³].
    d : Nenndurchmesser [mm]
    """

    return math.pi * (d ** 3) / 16.0


def computeStresses(Fx, Fy, Fz, Mx, My, Mz, d, A_sp, fp, fm, av):
    """
    Berechnet Spannungen und Auslastung einer Schraube (konservativ).
    - Axial- und Biegespannung werden betragsmäßig überlagert.
    - Vergleichsspannung nach von-Mises-Kriterium.

    Input:
      Fx,Fy,Fz : Kräfte [N]
      Mx,My,Mz : Momente [Nmm]
      d        : Schraubendurchmesser [mm]
      f        : Streckgrenze [MPa]

    Output:
      (sigN, tau, sigB, sigmaV, utilization)
    """

    Q = math.sqrt(Fy**2 + Fz**2) # Querkraft

    # ------
    # Als Balken und nach Mises
    # ------
    k_t = 0.5 # Verringerungsfaktor der Torsionsbeanspruchung

    #A = getArea(d) # Nennquerschnitt
    A = A_sp
    d = math.sqrt(4*A/math.pi) # Durchmesser aus Spannungsquerschnitt

    Wb = getWb(d)
    Wt = getWt(d)

    sigN = Fx / A
    tau = Q / A + abs(Mx) * k_t / Wt
    sigB = math.sqrt(My**2 + Mz**2) / Wb

    sigmaV = math.sqrt((abs(sigN) + abs(sigB))**2 + 3 * tau**2)
    utilizationMises = sigmaV / fp

    # ------
    # Nach Eurocode
    # ------
    y_M2 = 1.25 # Beanspruchbarkeit von Schrauben
    K2 = 0.9 # Reduzierungskoeffizient für Zug (Schraube ohne Senkkopf)

    utilizitationShear =  Q/(fm * A_sp * av / y_M2) # Auslastung Abscherung
    utilizitationTension = Fx/(fm * A_sp * K2 / y_M2) # Auslastung Zug
    utilizationEuroCode = utilizitationShear + utilizitationTension/1.4

    return sigN, tau, sigB, sigmaV, utilizationMises*100, utilizitationShear*100, utilizitationTension*100, utilizationEuroCode*100


def write_CSV(path, file_name, data, seperator="."):
    """CSV-Datei schreiben – mit Ersetzen von Dezimalpunkten für IronPython (ACT)"""
    
    if not isEmpty(data):
        try:
            output_data = []

            for row in data:
                new_row = []
                for cell in row:
                    cell_str = str(cell)
                    if isinstance(cell, int) or isinstance(cell, float): # eventuell wieder ruasnehmen. Sorgt dafür, dass Header "." enthalten kann
                        if "." in cell_str and seperator != ".":
                            cell_str = cell_str.replace(".", seperator)
                    new_row.append(cell_str)
                output_data.append(new_row)

            # Dateipfad manuell zusammensetzen (kein os.path.join in IronPython)
            full_path = path + "\\" + str(file_name) + ".csv"

            with open(full_path, "wb") as f:
                writer = csv.writer(f, delimiter=';', lineterminator='\n')
                writer.writerows(output_data)

            print("---> Successfully written data: " + file_name + ".csv")

        except Exception as e:
            print("---> Unsuccessfully written data: " + file_name + ".csv")
            print(str(e))


# --- Form-Klasse ---
class PathSelectionForm(Form):
    def __init__(self, default_path):
        self.Text = "Datei-Ablage auswählen"
        self.Size = Size(600,180)
        self.StartPosition = FormStartPosition.CenterScreen

        self.selected_path = None

        # Label
        self.label = Label()
        self.label.Text = "Bitte Pfad angeben oder auswählen:"
        self.label.Location = Point(10,10)
        self.label.Size = Size(560,20)
        self.Controls.Add(self.label)

        # Eingabebox für Pfad (editierbar!)
        self.pathBox = TextBox()
        self.pathBox.Location = Point(10,35)
        self.pathBox.Size = Size(460,20)
        self.pathBox.Text = default_path
        self.Controls.Add(self.pathBox)

        # Durchsuchen-Button
        self.browseButton = Button()
        self.browseButton.Text = "Durchsuchen..."
        self.browseButton.Location = Point(480,33)
        self.browseButton.Size = Size(100,25)
        self.browseButton.Click += self.browsePath
        self.Controls.Add(self.browseButton)

        # OK-Button
        self.okButton = Button()
        self.okButton.Text = "OK"
        self.okButton.Location = Point(180,90)
        self.okButton.Size = Size(100,30)
        self.okButton.Click += self.usePath
        self.Controls.Add(self.okButton)

        # Abbrechen-Button
        self.cancelButton = Button()
        self.cancelButton.Text = "Abbrechen"
        self.cancelButton.Location = Point(320,90)
        self.cancelButton.Size = Size(100,30)
        self.cancelButton.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.cancelButton)

        self.AcceptButton = self.okButton
        self.CancelButton = self.cancelButton

    def browsePath(self, sender, args):
        dlg = FolderBrowserDialog()
        dlg.Description = "Bitte Zielordner auswählen"
        dlg.SelectedPath = self.pathBox.Text
        if dlg.ShowDialog() == DialogResult.OK:
            self.pathBox.Text = dlg.SelectedPath

    def usePath(self, sender, args):
        text = self.pathBox.Text.strip()
        if text:
            self.selected_path = text
            self.DialogResult = DialogResult.OK
            self.Close()


# --------------------
# --------------------
# MAIN
# --------------------
# --------------------

evaluateBolts = True


# --------------------
# Datenpfade 
# --------------------
#user_DIR = wbjn.ExecuteCommand(ExtAPI, "returnValue(GetUserFilesDirectory())")
projectDIR = ExtAPI.DataModel.Project.ProjectDirectory
user_DIR = os.path.join(projectDIR, "exports")

date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
default_model_DIR = makeFolder(os.path.join(user_DIR,date+"_Reactions"))    

# Formular anzeigen
form = PathSelectionForm(default_model_DIR)
result = form.ShowDialog()

if result == DialogResult.OK and form.selected_path:
    model_DIR = makeFolder(form.selected_path)
    try:
        os.startfile(model_DIR)
    except Exception as e:
        print("Explorer konnte nicht geöffnet werden. Pfad zum Kopieren:")
        print(model_DIR)
else:
    print("Skript abgebrochen, da kein Pfad ausgewählt wurde.")
    sys.exit()



# --------------------
# Speichere Aktuelles Einheitensystem und wechsle dann auf t mm s
# --------------------
initialUnitSystem  = ExtAPI.Application.ActiveUnitSystem
ExtAPI.Application.ActiveUnitSystem = MechanicalUnitSystem.StandardNMMton

# --------------------
# Typendefinitionen
# --------------------
accType = Ansys.ACT.Automation.Mechanical.BoundaryConditions.Acceleration
gravType = Ansys.ACT.Automation.Mechanical.BoundaryConditions.EarthGravity
loadTypesList = [
    Ansys.ACT.Automation.Mechanical.BoundaryConditions.RemoteForce,
    Ansys.ACT.Automation.Mechanical.BoundaryConditions.Force
    ]
resultForceType = Ansys.ACT.Automation.Mechanical.Results.ProbeResults.ForceReaction
resultMomentType = Ansys.ACT.Automation.Mechanical.Results.ProbeResults.MomentReaction
bcType = LocationDefinitionMethod.BoundaryCondition
contactType = LocationDefinitionMethod.ContactRegion
jointType = Ansys.ACT.Automation.Mechanical.Results.ProbeResults.JointProbe
beamType = Ansys.ACT.Automation.Mechanical.Results.ProbeResults.BeamProbe
probeForceType = ProbeResultType.ForceReaction
probeMomentType = ProbeResultType.MomentReaction

boltTypes = [resultForceType, resultMomentType, beamType, jointType]
boltContactExclusion = contactType

connectionGroupTypes = [
    Ansys.ACT.Automation.Mechanical.Connections.ConnectionGroup,
    Ansys.ACT.Automation.Mechanical.TreeGroupingFolder
]

boundaryTypes = [Ansys.ACT.Automation.Mechanical.BoundaryConditions.RemoteDisplacement,
                Ansys.ACT.Automation.Mechanical.BoundaryConditions.Displacement,
                Ansys.ACT.Automation.Mechanical.BoundaryConditions.FixedSupport]

# --------------------
# Mass
# --------------------
print("Get model mass")
modelMass = float(extractValue(ExtAPI.DataModel.Project.Model.Geometry.Mass,"."))*1000 # von t in kg umwandeln

# --------------------
# Alle connections sammeln
# --------------------
contactObjectList = []
jointObjectList = []
beamObjectList = []

try:
    for item in Model.Connections.Children:
        collectAllConnections(item, contactObjectList, jointObjectList, beamObjectList, connectionGroupTypes)
except:
    print("SOMETHING WENT WRONG DURING GATHERING OF CONNECTION OBJECTS")


for analysisIndex, analysis in enumerate(ExtAPI.DataModel.AnalysisList):
    print("Read analysis: " +analysis.Name)
    analysis_DIR = makeFolder(os.path.join(model_DIR,"A" + str(analysisIndex+1) + "_" + adjustPath(analysis.Name)))

    # Tabellen fuer csv
    resultRows = []
    boltResultRows = []

    # Zur Laufzeitoptimierung bereits benutze Reaktion herausnehmen
    usedResults = []

    # Contacts
    for contactObject in contactObjectList:
        print("Read contact: " + contactObject.Name)
        check_for_stop(user_DIR)

        # initialize lists and flags
        Fx_list, Fy_list, Fz_list = [], [], []
        Mx_list, My_list, Mz_list = [], [], []
        time_list = []
        hasForce = False
        hasMoment = False

        for item in analysis.Solution.Children:
            try:
                if item.ContactRegionSelection == contactObject and item not in usedResults:
                    # Force
                    if item.GetType() == resultForceType and not hasForce:
                        hasForce = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        contactForceTable = cleanBCTable(rawTable, "force")
                        Fx_list = tryGetData("X [N]", contactForceTable)
                        Fy_list = tryGetData("Y [N]", contactForceTable)
                        Fz_list = tryGetData("Z [N]", contactForceTable)
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", contactForceTable)

                    # Moment
                    elif item.GetType() == resultMomentType and not hasMoment:
                        hasMoment = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        contactMomentTable = cleanBCTable(rawTable, "moment")
                        Mx_list = tryGetData("X [Nmm]", contactMomentTable)
                        My_list = tryGetData("Y [Nmm]", contactMomentTable)
                        Mz_list = tryGetData("Z [Nmm]", contactMomentTable)
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", contactMomentTable)
            except:
                pass

            if hasForce and hasMoment:
                break

        if not isEmpty(time_list):
            print("\t-> " + item.Name + " will be evaluated")

            for i in range(len(time_list)):
                Fx = Fx_list[i] if i < len(Fx_list) else 0.0
                Fy = Fy_list[i] if i < len(Fy_list) else 0.0
                Fz = Fz_list[i] if i < len(Fz_list) else 0.0
                Mx = Mx_list[i] if i < len(Mx_list) else 0.0
                My = My_list[i] if i < len(My_list) else 0.0
                Mz = Mz_list[i] if i < len(Mz_list) else 0.0

                resultRow = [
                    adjustName(contactObject.Name),
                    "Contact",
                    round(time_list[i], 3),
                    round(Fx, 0),
                    round(Fy, 0),
                    round(Fz, 0),
                    round(Mx, 0),
                    round(My, 0),
                    round(Mz, 0),
                ]

                resultRows.append(resultRow)
                printTable(resultRow)


    # Joints
    for jointObject in jointObjectList:
        print("Read joint: " + jointObject.Name)
        check_for_stop(user_DIR)

        # initialize lists
        hasForce = False
        hasMoment = False
        Fx_list, Fy_list, Fz_list = [], [], []
        Mx_list, My_list, Mz_list = [], [], []
        time_list = []

        for item in analysis.Solution.Children:
            try:
                if item.BoundaryConditionSelection == jointObject and item not in usedResults:
                    
                    if item.ResultType == probeForceType and not hasForce:
                        hasForce = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        jointForceTable = cleanBCTable(rawTable, "force")
                        Fx_list = tryGetData("X [N]", jointForceTable)
                        Fy_list = tryGetData("Y [N]", jointForceTable)
                        Fz_list = tryGetData("Z [N]", jointForceTable)
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", jointForceTable)

                    elif item.ResultType == probeMomentType and not hasMoment:
                        hasMoment = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        jointMomentTable = cleanBCTable(rawTable, "moment")
                        Mx_list = tryGetData("X [Nmm]", jointMomentTable)
                        My_list = tryGetData("Y [Nmm]", jointMomentTable)
                        Mz_list = tryGetData("Z [Nmm]", jointMomentTable)
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", jointMomentTable)

            except:
                pass

            if hasForce and hasMoment:
                break
        if not isEmpty(time_list):
            print("\t-> " + item.Name + " will be evaluated")
            for i in range(len(time_list)):
                Fx = Fx_list[i] if i < len(Fx_list) else 0.0
                Fy = Fy_list[i] if i < len(Fy_list) else 0.0
                Fz = Fz_list[i] if i < len(Fz_list) else 0.0
                Mx = Mx_list[i] if i < len(Mx_list) else 0.0
                My = My_list[i] if i < len(My_list) else 0.0
                Mz = Mz_list[i] if i < len(Mz_list) else 0.0

                resultRow = [
                    adjustName(jointObject.Name),
                    "Joint",
                    round(time_list[i], 3),
                    round(Fx, 0),
                    round(Fy, 0),
                    round(Fz, 0),
                    round(Mx, 0),
                    round(My, 0),
                    round(Mz, 0),
                ]

                resultRows.append(resultRow)
                printTable(resultRow)

                if evaluateBolts:
                    boltData = getBoltDataFromName(jointObject.Name)
                    if boltData:
                        d, A_sp, fp, fm, av = boltData
                        sigN, tau, sigB, sigmaV, utilizationMises, utilizitationShear, utilizitationTension, utilizationEuroCode = computeStresses(Fz, Fx, Fy, Mz, Mx, My, d, A_sp, fp, fm, av)
                        boltResultRow = [
                            adjustName(jointObject.Name),
                            "Joint",
                            int(d),
                            round(A_sp, 1),
                            int(fp),
                            int(fm),
                            round(time_list[i], 3),
                            round(Fx, 0),
                            round(Fy, 0),
                            round(Fz, 0),
                            round(Mx, 0),
                            round(My, 0),
                            round(Mz, 0),
                            round(sigN, 0),
                            round(tau, 0),
                            round(sigB, 0),
                            round(sigmaV, 0),
                            round(utilizationMises, 0),
                            round(utilizitationTension, 0),
                            round(utilizitationShear, 0),
                            round(utilizationEuroCode, 0)
                        ]

                        boltResultRows.append(boltResultRow)


    # Beams
    for beamObject in beamObjectList:
        print("Read beam: " + beamObject.Name)
        check_for_stop(user_DIR)

        # initialize lists and flag
        Fx_list, Fy_list, Fz_list = [], [], []
        Mx_list, My_list, Mz_list = [], [], []
        time_list = []
        hasBeenProcessed = False

        for item in analysis.Solution.Children:
            try:
                if item.BoundaryConditionSelection == beamObject and item not in usedResults and not hasBeenProcessed:
                    hasBeenProcessed = True
                    usedResults.append(item)
                    print("\t-> " + item.Name + " is evaluated")

                    rawTable = readTabularData(item, ".")
                    beamForceTable, beamMomentTable = splitBeamTable(rawTable)

                    Fx_list = tryGetData("Axial [N]", beamForceTable)
                    Fy_list = tryGetData("Quer I [N]", beamForceTable)
                    Fz_list = tryGetData("Quer J [N]", beamForceTable)

                    Mx_list = tryGetData("Torsion [Nmm]", beamMomentTable)
                    My_list = tryGetData("Biegung I [Nmm]", beamMomentTable)
                    Mz_list = tryGetData("Biegung J [Nmm]", beamMomentTable)

                    time_list = tryGetData("Time [s]", beamForceTable)

            except:
                pass

            if hasBeenProcessed:
                break

        print(time_list)
        if not isEmpty(time_list):
            print("\t-> " + item.Name + " will be evaluated")

            for i in range(len(time_list)):
                Fx = Fx_list[i] if i < len(Fx_list) else 0.0
                Fy = Fy_list[i] if i < len(Fy_list) else 0.0
                Fz = Fz_list[i] if i < len(Fz_list) else 0.0
                Mx = Mx_list[i] if i < len(Mx_list) else 0.0
                My = My_list[i] if i < len(My_list) else 0.0
                Mz = Mz_list[i] if i < len(Mz_list) else 0.0

                resultRow = [
                    adjustName(beamObject.Name),
                    "Beam",
                    round(time_list[i], 3),
                    round(Fx, 0),
                    round(Fy, 0),
                    round(Fz, 0),
                    round(Mx, 0),
                    round(My, 0),
                    round(Mz, 0),
                ]

                resultRows.append(resultRow)
                printTable(resultRow)

                if evaluateBolts:
                    boltData = getBoltDataFromName(beamObject.Name)
                    if not boltData:
                        d, A_sp, fp, fm, av = boltData
                        sigN, tau, sigB, sigmaV, utilizationMises, utilizitationShear, utilizitationTension, utilizationEuroCode = computeStresses(Fx, Fy, Fz, Mx, My, Mz, d, A_sp, fp, fm, av)
                        boltResultRow = [
                            adjustName(beamObject.Name),
                            "Beam",
                            int(d),
                            round(A_sp, 1),
                            int(fp),
                            int(fm),
                            round(time_list[i], 3),
                            round(Fx, 0),
                            round(Fy, 0),
                            round(Fz, 0),
                            round(Mx, 0),
                            round(My, 0),
                            round(Mz, 0),
                            round(sigN, 0),
                            round(tau, 0),
                            round(sigB, 0),
                            round(sigmaV, 0),
                            round(utilizationMises, 0),
                            round(utilizitationTension, 0),
                            round(utilizitationShear, 0),
                            round(utilizationEuroCode, 0)
                        ]

                        boltResultRows.append(boltResultRow)


    header = [
        "Name", "Typ", "Time [s]",
        "Fx [N]", "Fy [N]", "Fz [N]", 
        "Mx [Nmm]", "My [Nmm]", "Mz [Nmm]"
    ]


    boltHeader = [
        "Name", "Typ", "d [mm]", "A_sp [mm^2]","Rp [MPa]", "Rm [MPa]", "Time [s]",
        "Fx [N]", "Fy [N]", "Fz [N]", "Mx [Nmm]", "My [Nmm]", "Mz [Nmm]",
        "(VDI) sig_Zug [MPa]", "(VDI) tau [MPa]", "(VDI) sig_b [MPa]",
        "(VDI) Mises [MPa]", "(VDI) Ausl. Mises [%]",
        "(EC3) Ausl. Zug [%]", "(EC3) Ausl. Schub [%]", "(EC3) Ausl. Komb [%]"
    ]

    # --------------------
    # Alle Boundary Conditions sammeln
    # --------------------
    boundaryObjectList = []

    try:
        for item in analysis.Children:
            collectAllBoundaries(item, boundaryObjectList, boundaryTypes)
    except:
        print("SOMETHING WENT WRONG DURING GATHERING OF BOUNDARY OBJECTS")


    for bcObject in boundaryObjectList:
        print("Read Boundary: " + bcObject.Name)
        check_for_stop(user_DIR)

        # initialize lists and flags
        Fx_list, Fy_list, Fz_list = [], [], []
        Mx_list, My_list, Mz_list = [], [], []
        time_list = []
        hasForce = False
        hasMoment = False

        for item in analysis.Solution.Children:
            try:
                if item.BoundaryConditionSelection == bcObject and item not in usedResults:

                    # Force
                    if item.GetType() == resultForceType and not hasForce:
                        hasForce = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        bcForceTable = cleanBCTable(rawTable, "force")
                        Fx_list = tryGetData("X [N]", bcForceTable)
                        Fy_list = tryGetData("Y [N]", bcForceTable)
                        Fz_list = tryGetData("Z [N]", bcForceTable)
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", bcForceTable)

                    # Moment
                    elif item.GetType() == resultMomentType and not hasMoment:
                        hasMoment = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        bcMomentTable = cleanBCTable(rawTable, "moment")
                        Mx_list = tryGetData("X [Nmm]", bcMomentTable)
                        My_list = tryGetData("Y [Nmm]", bcMomentTable)
                        Mz_list = tryGetData("Z [Nmm]", bcMomentTable)
                        # Falls die Zeitliste noch leer ist (z.B. kein Force-Item gefunden), füllen wir sie hier
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", bcMomentTable)

            except:
                pass

            if hasForce and hasMoment:
                break

        if not isEmpty(time_list):
            print("\t-> " + item.Name + " will be evaluated")

            for i in range(len(time_list)):
                Fx = Fx_list[i] if i < len(Fx_list) else 0.0
                Fy = Fy_list[i] if i < len(Fy_list) else 0.0
                Fz = Fz_list[i] if i < len(Fz_list) else 0.0
                Mx = Mx_list[i] if i < len(Mx_list) else 0.0
                My = My_list[i] if i < len(My_list) else 0.0
                Mz = Mz_list[i] if i < len(Mz_list) else 0.0

                resultRow = [
                    adjustName(bcObject.Name),
                    "Boundary Condition",
                    round(time_list[i], 3),
                    round(Fx, 0),
                    round(Fy, 0),
                    round(Fz, 0),
                    round(Mx, 0),
                    round(My, 0),
                    round(Mz, 0),
                ]

                resultRows.append(resultRow)
                printTable(resultRow)

    connectionTable = []
    connectionTable.append(header)
    connectionTable.extend(resultRows)

    if evaluateBolts:
        boltsTable = []
        boltsTable.append(boltHeader)
        boltsTable.extend(boltResultRows)

    printTable(connectionTable)
    printTable(boltsTable)

    # -------
    # CSV-Datein herausschreiben
    # -------

    if not isEmpty(resultRows):
        write_CSV(analysis_DIR, "Reactions", connectionTable, seperator)

    if not isEmpty(boltResultRows): 
        write_CSV(analysis_DIR, "Bolts", boltsTable, seperator)

# Stelle initiales Einheitensstem wieder her
ExtAPI.Application.ActiveUnitSystem = initialUnitSystem
