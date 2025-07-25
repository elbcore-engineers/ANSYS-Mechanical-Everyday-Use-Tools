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


def combineForceTables(listOfTables, nameList, nameSum):
    """
    Kombiniert Tabellen nebeneinander:
    - Erste Spalte: Zeit (nur aus erster Tabelle)
    - Dann: X/Y/Z-Kräfte je Eintrag in listOfTables (ohne deren Zeitspalten)
    - Optional: Abschließend Summenspalten (nur wenn nameSum gesetzt ist)
    """
    if len(listOfTables) != len(nameList):
        print("Fehler: Tabellen- und Namensanzahl stimmen nicht überein.")
        return []
    
    if isEmpty(nameList):
        return []

    combinedTable = []
    maxLen = max(len(table) for table in listOfTables)

    # Tabellen auf gleiche Länge bringen
    for i in range(len(listOfTables)):
        while len(listOfTables[i]) < maxLen:
            listOfTables[i].append([""] * 4)

    # Erste Zeile: Gruppennamen
    headerLine1 = ["", ""]
    for name in nameList:
        headerLine1 += [name, "", ""]
    if nameSum:  # Summen-Header nur, wenn nameSum gesetzt ist
        headerLine1 += [nameSum, "", ""]
    combinedTable.append(headerLine1)

    # Zweite Zeile: Fixe Achsen
    headerLine2 = ["Time [s]"]
    for _ in nameList:
        headerLine2 += ["X [N]", "Y [N]", "Z [N]"]
    if nameSum:
        headerLine2 += ["X [N]", "Y [N]", "Z [N]"]
    combinedTable.append(headerLine2)

    # Datenzeilen
    for rowIndex in range(1, maxLen):
        row = []

        # Zeit aus erster Tabelle
        try:
            time = listOfTables[0][rowIndex][0]
        except:
            time = ""
        row.append(time)

        fx_total = 0.0
        fy_total = 0.0
        fz_total = 0.0

        for table in listOfTables:
            try:
                fx = float(table[rowIndex][1].replace(",", ".")) if len(table[rowIndex]) > 1 else 0.0
                fy = float(table[rowIndex][2].replace(",", ".")) if len(table[rowIndex]) > 2 else 0.0
                fz = float(table[rowIndex][3].replace(",", ".")) if len(table[rowIndex]) > 3 else 0.0
            except:
                fx, fy, fz = 0.0, 0.0, 0.0

            row += [str(fx), str(fy), str(fz)]

            fx_total += fx
            fy_total += fy
            fz_total += fz

        # Summenspalten nur anhängen, wenn nameSum gesetzt ist
        if nameSum:
            row += [str(fx_total), str(fy_total), str(fz_total)]

        combinedTable.append(row)

    return combinedTable


def combineMomentTables(listOfTables, nameList, nameSum):
    """
    Kombiniert Tabellen nebeneinander:
    - Erste Spalte: Zeit (nur aus erster Tabelle)
    - Dann: X/Y/Z-Kräfte je Eintrag in listOfTables (ohne deren Zeitspalten)
    - Optional: Abschließend Summenspalten (nur wenn nameSum gesetzt ist)
    """
    if len(listOfTables) != len(nameList):
        print("Fehler: Tabellen- und Namensanzahl stimmen nicht überein.")
        return []
    
    if isEmpty(nameList):
        return []

    combinedTable = []
    maxLen = max(len(table) for table in listOfTables)

    # Tabellen auf gleiche Länge bringen
    for i in range(len(listOfTables)):
        while len(listOfTables[i]) < maxLen:
            listOfTables[i].append([""] * 4)

    # Erste Zeile: Gruppennamen
    headerLine1 = ["", ""]
    for name in nameList:
        headerLine1 += [name, "", ""]
    if nameSum:  # Summen-Header nur, wenn nameSum gesetzt ist
        headerLine1 += [nameSum, "", ""]
    combinedTable.append(headerLine1)

    # Zweite Zeile: Fixe Achsen
    headerLine2 = ["Time [s]"]
    for _ in nameList:
        headerLine2 += ["X [Nmm]", "Y [Nmm]", "Z [Nmm]"]
    if nameSum:
        headerLine2 += ["X [Nmm]", "Y [Nmm]", "Z [Nmm]"]
    combinedTable.append(headerLine2)

    # Datenzeilen
    for rowIndex in range(1, maxLen):
        row = []

        # Zeit aus erster Tabelle
        try:
            time = listOfTables[0][rowIndex][0]
        except:
            time = ""
        row.append(time)

        fx_total = 0.0
        fy_total = 0.0
        fz_total = 0.0

        for table in listOfTables:
            try:
                fx = float(table[rowIndex][1].replace(",", ".")) if len(table[rowIndex]) > 1 else 0.0
                fy = float(table[rowIndex][2].replace(",", ".")) if len(table[rowIndex]) > 2 else 0.0
                fz = float(table[rowIndex][3].replace(",", ".")) if len(table[rowIndex]) > 3 else 0.0
            except:
                fx, fy, fz = 0.0, 0.0, 0.0

            row += [str(fx), str(fy), str(fz)]

            fx_total += fx
            fy_total += fy
            fz_total += fz

        # Summenspalten nur anhängen, wenn nameSum gesetzt ist
        if nameSum:
            row += [str(fx_total), str(fy_total), str(fz_total)]

        combinedTable.append(row)

    return combinedTable



def cleanBCForceTable(rawTable):
    """
    Wandelt ein Raw-Table mit möglichen X, Y, Z und Total-Spalten in ein sauberes Format um:
    ["Time [s]", "X [N]", "Y [N]", "Z [N]"]

    Notewendig bei Lagerreaktionen, wo der Input Total mitgeschleppt wird

    Alle fehlenden Richtungen werden mit 0 ersetzt.
    Die Total-Spalte wird ignoriert.
    """
    if not rawTable or len(rawTable) < 2:
        return []

    header = rawTable[0]
    time_index = -1
    x_index = -1
    y_index = -1
    z_index = -1

    # Spaltenindexe ermitteln
    for i, col in enumerate(header):
        col_lower = col.lower()
        if "time" in col_lower:
            time_index = i
        elif "(x)" in col_lower:
            x_index = i
        elif "(y)" in col_lower:
            y_index = i
        elif "(z)" in col_lower:
            z_index = i
        elif "(total)" in col_lower:
            continue  # explizit ignorieren

    # Header für neue Tabelle
    cleanedTable = []
    cleanedTable.append(["Time [s]", "X [N]", "Y [N]", "Z [N]"])

    # Datenzeilen aufbauen
    for row in rawTable[1:]:
        newRow = []

        # Zeitwert
        try:
            time = row[time_index] if time_index != -1 else ""
        except:
            time = ""

        newRow.append(time)

        # X/Y/Z – prüfen, ob vorhanden, sonst 0
        for idx in [x_index, y_index, z_index]:
            try:
                val = row[idx].replace(",", ".")
                val = float(val)
                newRow.append(str(float(val)))
            except:
                newRow.append("0.0")

        cleanedTable.append(newRow)

    return cleanedTable


def cleanBCMomentTable(rawTable):
    """
    Wandelt ein Raw-Table mit möglichen X, Y, Z und Total-Spalten in ein sauberes Format um:
    ["Time [s]", "X [Nmm]", "Y [Nmm]", "Z [Nmm]"]

    Notewendig bei Lagerreaktionen, wo der Input Total mitgeschleppt wird

    Alle fehlenden Richtungen werden mit 0 ersetzt.
    Die Total-Spalte wird ignoriert.
    """
    if not rawTable or len(rawTable) < 2:
        return []

    header = rawTable[0]
    time_index = -1
    x_index = -1
    y_index = -1
    z_index = -1

    # Spaltenindexe ermitteln
    for i, col in enumerate(header):
        col_lower = col.lower()
        if "time" in col_lower:
            time_index = i
        elif "(x)" in col_lower:
            x_index = i
        elif "(y)" in col_lower:
            y_index = i
        elif "(z)" in col_lower:
            z_index = i
        elif "(total)" in col_lower:
            continue  # explizit ignorieren

    # Header für neue Tabelle
    cleanedTable = []
    cleanedTable.append(["Time [s]", "X [Nmm]", "Y [Nmm]", "Z [Nmm]"])

    # Datenzeilen aufbauen
    for row in rawTable[1:]:
        newRow = []

        # Zeitwert
        try:
            time = row[time_index] if time_index != -1 else ""
        except:
            time = ""

        newRow.append(time)

        # X/Y/Z – prüfen, ob vorhanden, sonst 0
        for idx in [x_index, y_index, z_index]:
            try:
                val = row[idx].replace(",", ".")
                val = float(val)
                newRow.append(str(float(val)))
            except:
                newRow.append("0.0")

        cleanedTable.append(newRow)

    return cleanedTable


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
model_DIR = makeFolder(os.path.join(user_DIR,date+"_Lagerreaktionen"))

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
probeForceType = ProbeResultType.ForceReaction
probeMomentType = ProbeResultType.MomentReaction


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

    combinedLoadTable = combineForceTables(tableListLoads, nameListLoads, nameSumLoad)


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
                cleanedTable = cleanBCForceTable(rawTable)
                tableListBc.append(cleanedTable)

    combinedBcTable = combineForceTables(tableListBc, nameListBc, nameSumBc)


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
                cleanedTable = cleanBCForceTable(rawTable)
                tableListJointForce.append(cleanedTable)

    combinedJointForceTable = combineForceTables(tableListJointForce, nameListJointForce, nameSumJointForce)

    # Moment
    nameListJointMoment = []
    tableListJointMoment  = []

    for item in analysis.Solution.Children:
        if item.GetType() == jointType:
            if item.ResultType == probeMomentType:
                nameListJointMoment.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCMomentTable(rawTable)
                tableListJointMoment.append(cleanedTable)

    combinedJointMomentTable = combineMomentTables(tableListJointMoment, nameListJointMoment, nameSumJointMoment)


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
                cleanedTable = cleanBCForceTable(rawTable)
                tableListContactForce.append(cleanedTable)

    combinedContactForceTable = combineForceTables(tableListContactForce, nameListContactForce, nameSumContactForce)

    # Moment
    nameListContactMoment = []
    tableListContactMoment  = []

    for item in analysis.Solution.Children:
        if item.GetType() == resultMomentType:
            if item.LocationMethod == contactType:
                nameListContactMoment.append(adjustName(item.Name))

                rawTable = readTabularData(item, ".")
                cleanedTable = cleanBCMomentTable(rawTable)
                tableListContactMoment.append(cleanedTable)

    combinedContactMomentTable = combineMomentTables(tableListContactMoment, nameListContactMoment, nameSumContactMoment)


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


# Stelle initiales Einheitensstem wieder her
ExtAPI.Application.ActiveUnitSystem = initialUnitSystem