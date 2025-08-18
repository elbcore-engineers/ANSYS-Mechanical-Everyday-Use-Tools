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

user_DIR = wbjn.ExecuteCommand(ExtAPI, "returnValue(GetUserFilesDirectory())")

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
    Checks if a file named 'stop' or 'stop.txt' exists in the given directory. 
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
    text = text.replace("= ","")
    text = text.replace("   "," ")
    text = text.replace("  "," ")
    text = text.replace(" "," ")
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
    text = text.replace("°","Degree")
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


def isEmpty(obj):
    return not obj or len(obj) == 0

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
    Die Total-Spalte wird ignoriert.
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

    for i in range(len(header)):
        col = str(header[i]).lower()

        # Zeitspalte
        if "time" in col:
            time_index = i

        # X, Y, Z erkennen, aber "total" ignorieren
        elif "x" in col and "total" not in col:
            x_index = i
        elif "y" in col and "total" not in col:
            y_index = i
        elif "z" in col and "total" not in col:
            z_index = i

        # Fallback: Akzeptiere "total moment/force x" nur wenn nichts anderes gefunden wurde
        elif "total" in col and "x" in col and x_index == -1:
            x_index = i
        elif "total" in col and "y" in col and y_index == -1:
            y_index = i
        elif "total" in col and "z" in col and z_index == -1:
            z_index = i

    # Neue Tabelle initialisieren
    cleanedTable = []
    cleanedTable.append(["Time [s]", "X [%s]" % unit, "Y [%s]" % unit, "Z [%s]" % unit])

    for i in range(1, len(rawTable)):
        row = rawTable[i]
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
                val = str(row[idx]).replace(",", ".")
                if "e+" in val or "e-" in val:
                    val = float(val)
                else:
                    val = float(val)
                newRow.append(val)
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


def collectAllConnections(container, connectionList):
    """
    Sammelt rekursiv alle Connections aus einer beliebig
    verschachtelten Containerstruktur.

    container      : aktuelles Connection-Objekt oder -Gruppe
    connectionList : Liste zum Auffüllen
    """
    # Manche Container haben Children, andere sind selbst schon Connections
    if hasattr(container, "Children"):
        for item in container.Children:
            itemType = item.GetType()
            if itemType in connectionGroupTypes:
                # rekursiv in Untergruppe
                collectAllConnections(item, connectionList)
            else:
                # direkte Connection gefunden
                connectionList.append(item)
    else:
        # kein Container, sondern direkte Connection
        connectionList.append(container)
            

def getBoltDataFromName(text): 
    """
    Extrahiert Nenndurchmesser [mm] und Streckgrenze [MPa] 
    aus dem Namen einer Schrauben-Connection.

    Eingabe:
      text : string (z.B. 'Bolt M12 8.8')

    Rückgabe:
      (d, f) : Durchmesser [mm], Streckgrenze [MPa]
      False  : wenn kein passender Eintrag gefunden
    """
    
    text = str(text)

    d = 0.1  # Standardwert Nenndurchmesser
    f = 640  # Standardwert Streckgrenze

    bolt_sizes = {
        "M64": 64.0, "M56": 56.0, "M48": 48.0, "M42": 42.0, "M36": 36.0,
        "M30": 30.0, "M27": 27.0, "M24": 24.0, "M22": 22.0, "M20": 20.0,
        "M18": 18.0, "M16": 16.0, "M14": 14.0, "M12": 12.0, "M10": 10.0,
        "M8": 8.0, "M6": 6.0, "M5": 5.0, "M4": 4.0, "M3": 3.0
    }

    yield_strengths = {
        "12.9": 940, "10.9": 720, "9.8": 660, "8.8": 640, "6.8": 480,
        "5.8": 420, "5.6": 300, "4.8": 340, "4.6": 240, "4.4": 150, "3.6": 180
    }

    for key in bolt_sizes:
        if key in text:
            d = bolt_sizes[key]
            break

    for key in yield_strengths:
        if key in text:
            f = yield_strengths[key]
            break

    if d == 0.1:
        return False
    else:
        return d, f
    

def tryGetData(name, table):
    """
    Liest eine Spalte aus einer Tabellendatenstruktur.
    Falls Werte nicht konvertierbar sind -> 999999999 einsetzen.
    
    name  : Spaltenname (string)
    table : 2D-Liste (erste Zeile = Header)

    Rückgabe: Liste mit floats (eine pro Zeile)
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


def computeStresses(Fx, Fy, Fz, Mx, My, Mz, d, f):
    """
    Berechnet Spannungen und Auslastung einer Schraube (konservativ).
    - Axial- und Biegespannung werden betragsmäßig überlagert.
    - Vergleichsspannung nach von-Mises-Kriterium.

    Eingaben:
      Fx,Fy,Fz : Kräfte [N]
      Mx,My,Mz : Momente [Nmm]
      d        : Schraubendurchmesser [mm]
      f        : Streckgrenze [MPa]

    Rückgabe:
      (sigN, tau, sigB, sigmaV, utilization)
    """

    A = getArea(d)
    Wb = getWb(d)
    Wt = getWt(d)

    sigN = Fx / A
    tau = math.sqrt(Fy**2 + Fz**2) / A + abs(Mx) / Wt
    sigB = math.sqrt(My**2 + Mz**2) / Wb

    sigmaV = math.sqrt((abs(sigN) + abs(sigB))**2 + 3 * tau**2)
    utilization = 100*sigmaV / f

    return sigN, tau, sigB, sigmaV, utilization


def write_CSV(path, file_name, data, seperator="."):
    """CSV-Datei schreiben – mit Ersetzen von Dezimalpunkten für IronPython (ACT)"""

    if not isEmpty(data):
        try:
            output_data = []

            for row in data:
                new_row = []
                for cell in row:
                    cell_str = str(cell)
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


# --------------------
# --------------------
# MAIN
# --------------------
# --------------------

# Datenpfade 
date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
model_DIR = makeFolder(os.path.join(user_DIR,date+"_Reaktionen"))

# Speichere Aktuelles Einheitensystem und wechsle dann auf t mm s
initialUnitSystem  = ExtAPI.Application.ActiveUnitSystem
ExtAPI.Application.ActiveUnitSystem = MechanicalUnitSystem.StandardNMMton


# Typendefinitionen
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

# Modell-Informationen
modelMass = float(extractValue(ExtAPI.DataModel.Project.Model.Geometry.Mass,"."))*1000 # von t in kg umwandeln


for analysisIndex, analysis in enumerate(ExtAPI.DataModel.AnalysisList):
    analysis_DIR = makeFolder(os.path.join(model_DIR,"A" + str(analysisIndex+1) + "_" + analysis.Name))

    # -------
    # [1] Lasten einlesen
    # -------

    nameListLoads = []
    tableListLoads = []

    # Lasten aus Beschleunigungen
    for item in analysis.Children:
        if item.GetType() == accType:
            nameListLoads.append(adjustName(item.Name))
            
            # wandel die Bescleunigungstabelle in eine Krafttabelle um
            accTable = readTabularData(item,".")
            accForceTable = createAccelerationForce(-modelMass, accTable)
            tableListLoads.append(accForceTable)
        elif item.GetType() == gravType:
            nameListLoads.append(adjustName(item.Name))
            
            # wandel die Bescleunigungstabelle in eine Krafttabelle um
            accTable = readTabularData(item,".")
            accForceTable = createAccelerationForce(modelMass, accTable)
            tableListLoads.append(accForceTable)
        
    # Lasten aus Einzellasten
    for item in analysis.Children:
        if item.GetType() in loadTypesList:
            nameListLoads.append(adjustName(item.Name))

            forceTable = convertLoadForceTable(readTabularData(item,"."))
            tableListLoads.append(forceTable)

    combinedLoadTable = combineTables(tableListLoads, nameListLoads, nameSumLoad, "force")


    # -------
    # [2] Lagerreaktionen einlesen
    # -------
    nameListBc = []
    tableListBc = []

    for item in analysis.Solution.Children:
        if item.GetType() == resultForceType:
            if item.LocationMethod == bcType:
                nameListBc.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCTable(rawTable,"force")
                tableListBc.append(cleanedTable)

    combinedBcTable = combineTables(tableListBc, nameListBc, nameSumBc, "force")


    # -------
    # [3] Lasten bilanzieren
    # -------

    combinedResultTable = combineResultTable(combinedLoadTable, combinedBcTable, nameloads=nameSumLoad, nameBCs=nameSumBc ,nameDiff=nameSumDif)


    # -------
    # [4] Gelenkreaktionen einlesen
    # -------

    # Kraft
    nameListJointForce = []
    tableListJointForce  = []

    for item in analysis.Solution.Children:
        if item.GetType() == jointType:
            if item.ResultType == probeForceType:
                nameListJointForce.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCTable(rawTable,"force")
                tableListJointForce.append(cleanedTable)

    combinedJointForceTable = combineTables(tableListJointForce, nameListJointForce, nameSumJointForce, "force")

    # Moment
    nameListJointMoment = []
    tableListJointMoment  = []

    for item in analysis.Solution.Children:
        if item.GetType() == jointType:
            if item.ResultType == probeMomentType:
                nameListJointMoment.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCTable(rawTable,"moment")
                tableListJointMoment.append(cleanedTable)

    combinedJointMomentTable = combineTables(tableListJointMoment, nameListJointMoment, nameSumJointMoment, "moment")


    # -------
    # [5] Kontaktreaktionen einlesen
    # -------

    # Kraft
    nameListContactForce = []
    tableListContactForce  = []

    for item in analysis.Solution.Children:
        if item.GetType() == resultForceType:
            if item.LocationMethod == contactType:
                nameListContactForce.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCTable(rawTable,"force")
                tableListContactForce.append(cleanedTable)

    combinedContactForceTable = combineTables(tableListContactForce, nameListContactForce, nameSumContactForce, "force")

    # Moment
    nameListContactMoment = []
    tableListContactMoment  = []

    for item in analysis.Solution.Children:
        if item.GetType() == resultMomentType:
            if item.LocationMethod == contactType:
                nameListContactMoment.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCTable(rawTable,"moment")
                tableListContactMoment.append(cleanedTable)

    combinedContactMomentTable = combineTables(tableListContactMoment, nameListContactMoment, nameSumContactMoment, "force")


    # -------
    # [6] Balken einlesen
    # -------
    nameListBeams = []
    tableListBeamsForce = []
    tableListBeamsMoment = []

    for item in analysis.Solution.Children:
        if item.GetType() == beamType:
            nameListBeams.append(adjustName(item.Name))
            rawTable = readTabularData(item, ".")
            beamForceTable, beamMomenttable = splitBeamTable(rawTable)

            tableListBeamsForce.append(beamForceTable)
            tableListBeamsMoment.append(beamMomenttable)      

    combinedBeamForceTable = combineTables(tableListBeamsForce, nameListBeams, nameSumBeamForce, "beamForce")
    combinedBeamMomentTable = combineTables(tableListBeamsMoment, nameListBeams, nameSumBeamMoment, "beamMoment")


    # -------
    # [7] Schrauben auswerten
    # -------

    # Hauptliste für alle Connections
   # Hauptliste für alle Connections
    allConnections = []

    # Start: oberste Ebene durchlaufen
    for item in Model.Connections.Children:
        collectAllConnections(item, allConnections)

    # Ausgabe testen
    for item in allConnections:
        print(item.Name)
    

    boltConnections = []
    boltConnectionsUsedForce = []
    boltConnectionsUsedMoments = []

    for connection in allConnections:
        if getBoltDataFromName(connection.Name):
            boltConnections.append(connection)

    resultRows = []

    for boltConnection in boltConnections:
        boltData = getBoltDataFromName(boltConnection.Name)
        if not boltData:
            continue
        d, f = boltData

        Fx_list = []
        Fy_list = []
        Fz_list = []
        Mx_list = []
        My_list = []
        Mz_list = []
        time_list = []

        for item in analysis.Solution.Children:

            # ---- Joints ----
            if item.GetType() == jointType and item.BoundaryConditionSelection == boltConnection:

                if item.ResultType == probeForceType and item.BoundaryConditionSelection not in boltConnectionsUsedForce:
                    boltConnectionsUsedForce.append(boltConnection)
                    rawTable = readTabularData(item, ".")
                    jointForceTable = cleanBCTable(rawTable,"force")
                    Fx_list = tryGetData("X [N]", jointForceTable)
                    Fy_list = tryGetData("Y [N]", jointForceTable)
                    Fz_list = tryGetData("Z [N]", jointForceTable)
                    time_list = tryGetData("Time [s]", jointForceTable)

                elif item.ResultType == probeMomentType and item.BoundaryConditionSelection not in boltConnectionsUsedMoments:
                    boltConnectionsUsedMoments.append(boltConnection)
                    rawTable = readTabularData(item, ".")
                    jointMomentTable = cleanBCTable(rawTable,"moment")

                    Mx_list = tryGetData("X [Nmm]", jointMomentTable)
                    My_list = tryGetData("Y [Nmm]", jointMomentTable)
                    Mz_list = tryGetData("Z [Nmm]", jointMomentTable)
                    time_list = tryGetData("Time [s]", jointMomentTable)

            # ---- Beams ----
            elif item.GetType() == beamType and item.BoundaryConditionSelection == boltConnection:
                rawTable = readTabularData(item, ".")
                beamForceTable, beamMomentTable = splitBeamTable(rawTable)

                Fx_list = tryGetData("Axial [N]", beamForceTable)
                Fy_list = tryGetData("Quer I [N]", beamForceTable)
                Fz_list = tryGetData("Quer J [N]", beamForceTable)

                Mx_list = tryGetData("Torsion [Nmm]", beamMomentTable)
                My_list = tryGetData("Biegung I [Nmm]", beamMomentTable)
                Mz_list = tryGetData("Biegung J [Nmm]", beamMomentTable)

                time_list = tryGetData("Time [s]", beamForceTable)

        if not time_list:
            continue

        for i in range(len(time_list)):
            Fx = Fx_list[i] if i < len(Fx_list) else 0.0
            Fy = Fy_list[i] if i < len(Fy_list) else 0.0
            Fz = Fz_list[i] if i < len(Fz_list) else 0.0
            Mx = Mx_list[i] if i < len(Mx_list) else 0.0
            My = My_list[i] if i < len(My_list) else 0.0
            Mz = Mz_list[i] if i < len(Mz_list) else 0.0

            sigN, tau, sigB, sigmaV, utilization = computeStresses(Fx, Fy, Fz, Mx, My, Mz, d, f)

            resultRow = [
                boltConnection.Name,
                int(d),
                int(f),
                round(time_list[i], 3),
                round(Fx, 2),
                round(Fy, 2),
                round(Fz, 2),
                round(Mx, 2),
                round(My, 2),
                round(Mz, 2),
                round(sigN, 2),
                round(tau, 2),
                round(sigB, 2),
                round(sigmaV, 2),
                round(utilization, 1)
            ]

            resultRows.append(resultRow)

    header = [
        "Name", "d [mm]", "f [MPa]", "Time [s]",
        "Fx [N]", "Fy [N]", "Fz [N]",
        "Mx [Nmm]", "My [Nmm]", "Mz [Nmm]",
        "Druck/Zug [MPa]", "Schub [MPa]", "Biegung [MPa]",
        "Mises [MPa]", "Auslastung [%]"
    ]

    boltComprehensiveTable = []
    boltComprehensiveTable.append(header)
    boltComprehensiveTable.extend(resultRows)




    # -------
    # CSV-Datein herausschreiben
    # -------

    if not isEmpty(nameListLoads):
        write_CSV(analysis_DIR, "01_Lasten", combinedLoadTable, seperator)
        
    if not isEmpty(nameListBc):
        write_CSV(analysis_DIR, "02_Randbedingungen", combinedBcTable, seperator)

    if not isEmpty(nameListLoads):
        write_CSV(analysis_DIR, "03_Kraftbilanz", combinedResultTable, seperator)

    if not isEmpty(nameListJointForce):
        write_CSV(analysis_DIR, "04_Gelenkkraft", combinedJointForceTable, seperator)

    if not isEmpty(nameListJointMoment):
        write_CSV(analysis_DIR, "05_Gelenkmoment", combinedJointMomentTable, seperator)

    if not isEmpty(nameListContactForce):
        write_CSV(analysis_DIR, "06_Kontaktkraft", combinedContactForceTable, seperator)

    if not isEmpty(nameListContactMoment):
        write_CSV(analysis_DIR, "07_Kontaktmoment", combinedContactMomentTable, seperator)

    if not isEmpty(nameListBeams):
        write_CSV(analysis_DIR, "08_Balkenkraft", combinedBeamForceTable, seperator)

    if not isEmpty(nameListBeams):
        write_CSV(analysis_DIR, "09_Balkenmoment", combinedBeamMomentTable, seperator)

    if not isEmpty(resultRows): 
        write_CSV(analysis_DIR, "10_Schraubenauslastungen", boltComprehensiveTable, seperator)


# Stelle initiales Einheitensstem wieder her
ExtAPI.Application.ActiveUnitSystem = initialUnitSystem
