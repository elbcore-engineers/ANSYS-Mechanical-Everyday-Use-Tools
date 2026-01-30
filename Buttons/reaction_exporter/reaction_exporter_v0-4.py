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


def cleanBCTable(rawTable, data_type):
    """
    Wandelt ein Raw-Table mit möglichen X, Y, Z und Total-Spalten in ein sauberes Format um.
    Gibt eine Tabelle im Format zurück:
        - "force"  → ["Time [s]", "X [N]", "Y [N]", "Z [N]"]
        - "moment" → ["Time [s]", "X [Nmm]", "Y [Nmm]", "Z [Nmm]"]

    Alle fehlenden Richtungen werden mit 0 ersetzt.
    Die echte Summenspalte (z.B. ".. (Total) [N]") wird ignoriert.
    Standardmäßig:
        - erste Spalte → "Time [s]"
        - zweite Spalte → X
        - dritte Spalte → Y
        - vierte Spalte → Z
    """

    if not rawTable or len(rawTable) < 2:
        return []

    # Einheit je nach Datentyp
    data_type_lower = str(data_type).lower()
    if data_type_lower == "moment":
        unit = "Nmm"
    else:
        unit = "N"

    # Standardmäßig Spaltenindizes
    time_index = 0
    x_index = 1
    y_index = 2
    z_index = 3

    # Neue Tabelle initialisieren
    cleanedTable = []
    cleanedTable.append(["Time [s]", "X [" + unit + "]", "Y [" + unit + "]", "Z [" + unit + "]"])

    for row in rawTable[1:]:
        newRow = []

        # Zeit
        try:
            timeVal = row[time_index]
            timeStr = str(timeVal).replace(",", ".").replace("s", "").strip()
            newRow.append(float(timeStr))
        except:
            newRow.append(0.0)

        # X, Y, Z – mit Fallback 0.0
        for idx in [x_index, y_index, z_index]:
            try:
                val = row[idx]
                valStr = str(val).replace(",", ".")
                newRow.append(float(valStr))
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

def collectAllConnections(container, contactObjectList, jointObjectList, beamObjectList, connectionGroupTypes):
    """
    Sammelt rekursiv alle Connections aus einer beliebig
    verschachtelten Containerstruktur.

    container      : aktuelles Connection-Objekt oder -Gruppe
    connectionList : Liste zum Auffüllen
    """
    # Manche Container haben Children, andere sind selbst schon Connections
    print("Open connection object with Name: " + container.Name)

    if hasattr(container, "Children") and len(container.Children) > 0:

        for item in container.Children:
            itemType = item.GetType()

            if itemType in connectionGroupTypes:
                # rekursiv in Untergruppe
                print("\t -> open antoher folder")
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
                print("\t -> didnt found a proper type")

    else:
        # kein Container, sondern direkte Connection
        itemType = container.GetType()

        if itemType == Ansys.ACT.Automation.Mechanical.Connections.ContactRegion:
            print("\t -> assigned to contact type")
            contactObjectList.append(container)

        elif itemType == Ansys.ACT.Automation.Mechanical.Connections.Joint:
            print("\t -> assigned to joint type")
            jointObjectList.append(container)

        elif itemType == Ansys.ACT.Automation.Mechanical.Connections.Beam:
            print("\t -> assigned to beam type")
            beamObjectList.append(container)

        else:
            print("\t -> didnt found a proper type")


def collectAllBoundaries(container, bouboundaryObjectList, boundaryTypes):
    # Manche Container haben Children, andere sind selbst schon boundaries
    print("Read boundary: " + container.Name)

    if item.GetType() in boundaryTypes:
        bouboundaryObjectList.append(item)

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


# --------------------
# Datenpfade 
# --------------------


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

    # Analysistype ermitteln -> unterschiedliche Exportformate
    if analysis.AnalysisType in [AnalysisType.Static, AnalysisType.Transient]:
        domain = "time"
        headerEnum = "Time [s]"
    elif analysis.AnalysisType in [AnalysisType.Harmonic, AnalysisType.ResponseSpectrum]:
        domain = "frequency"
        headerEnum = "Freq [Hz]"
    else:
        domain = False
        continue

    header = [
        "Name", "Typ", headerEnum,
        "Fx [N]", "Fy [N]", "Fz [N]", 
        "Mx [Nmm]", "My [Nmm]", "Mz [Nmm]"
    ]

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
                if item.ContactRegionSelection == contactObject and item not in usedResults and item.ResultSelection == ProbeDisplayFilter.All and item.By == SetDriverStyle.Time:
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
                    elif item.GetType() == resultMomentType and not hasMoment and item.ResultSelection == ProbeDisplayFilter.All and item.By == SetDriverStyle.Time:
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
                #printTable(resultRow)

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
            print("check item: "+str(item.Name))
            try:
                if item.BoundaryConditionSelection == bcObject and item not in usedResults:
                    print("started loop")
                    # Force
                    if item.GetType() == resultForceType and not hasForce and item.ResultSelection == ProbeDisplayFilter.All:
                        print("Got Force")
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
                    elif item.GetType() == resultMomentType and not hasMoment and item.ResultSelection == ProbeDisplayFilter.All:
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
                    
                    if item.ResultType == probeForceType and not hasForce and item.ResultSelection == ProbeDisplayFilter.All:
                        hasForce = True
                        usedResults.append(item)
                        rawTable = readTabularData(item, ".")
                        jointForceTable = cleanBCTable(rawTable, "force")
                        Fx_list = tryGetData("X [N]", jointForceTable)
                        Fy_list = tryGetData("Y [N]", jointForceTable)
                        Fz_list = tryGetData("Z [N]", jointForceTable)
                        if isEmpty(time_list):
                            time_list = tryGetData("Time [s]", jointForceTable)

                    elif item.ResultType == probeMomentType and not hasMoment and item.ResultSelection == ProbeDisplayFilter.All:
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

    if domain == "time":
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
                    if item.BoundaryConditionSelection == beamObject and item not in usedResults and not hasBeenProcessed and item.ResultSelection == ProbeDisplayFilter.All:
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

                    # Korrektur:
                    # - Vorher wurde angenommen, dass "Shear I [N]" "Shear J [N]" sich auf die Hauptrichtung beziehen.
                    # - Sie beziehen sich aber auf die Knoten.
                    # - Daher pauschal y als Maximalwert und z als 0-Wert
                    if abs(Fy) >= abs(Fz):
                        Fy = Fy
                        Fz = 0
                    else:
                        Fy = Fz
                        Fz = 0

                    if abs(My) >= abs(Mz):
                        My = My
                        Mz = 0
                    else:
                        My = Mz
                        Mz = 0

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
        
    elif domain == "frequency":
        for beamObject in beamObjectList:

            # initialize lists and flags
            print("Read beam: " + beamObject.Name)
            check_for_stop(user_DIR)

            # initialize lists and flag
            Fx_list, Fy_list, Fz_list = [], [], []
            Mx_list, My_list, Mz_list = [], [], []
            time_list = []
            
            hasForce = False
            hasMoment = False

            for item in analysis.Solution.Children:
                try:
                    if item.Beam == beamObject and item not in usedResults:
                        # Force
                        if item.GetType() == resultForceType and not hasForce and item.ResultSelection == ProbeDisplayFilter.All and item.By == SetDriverStyle.Time:
                            hasForce = True
                            usedResults.append(item)
                            rawTable = readTabularData(item, ".")
                            beamForceTable = cleanBCTable(rawTable, "force")
                            Fx_list = tryGetData("X [N]", beamForceTable)
                            Fy_list = tryGetData("Y [N]", beamForceTable)
                            Fz_list = tryGetData("Z [N]", beamForceTable)
                            if isEmpty(time_list):
                                time_list = tryGetData("Time [s]", beamForceTable)

                        # Moment
                        elif item.GetType() == resultMomentType and not hasMoment and item.ResultSelection == ProbeDisplayFilter.All and item.By == SetDriverStyle.Time:
                            hasMoment = True
                            usedResults.append(item)
                            rawTable = readTabularData(item, ".")
                            beamMomentTable = cleanBCTable(rawTable, "moment")
                            Mx_list = tryGetData("X [Nmm]", beamMomentTable)
                            My_list = tryGetData("Y [Nmm]", beamMomentTable)
                            Mz_list = tryGetData("Z [Nmm]", beamMomentTable)
                            if isEmpty(time_list):
                                time_list = tryGetData("Time [s]", beamMomentTable)
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


    connectionTable = []
    connectionTable.append(header)
    connectionTable.extend(resultRows)

    printTable(connectionTable)

    # -------
    # CSV-Datein herausschreiben
    # -------

    if not isEmpty(resultRows):
            write_CSV(analysis_DIR, "Reactions", connectionTable, seperator)

     

# Stelle initiales Einheitensstem wieder her
ExtAPI.Application.ActiveUnitSystem = initialUnitSystem