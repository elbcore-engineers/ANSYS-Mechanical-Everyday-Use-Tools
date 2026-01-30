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

from System.Windows.Forms import Application, Form, Button, Label, CheckBox, DialogResult, FormStartPosition
from System.Drawing import Size, Point

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

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Application, Form, Label, Button, CheckBox, DialogResult, FormStartPosition
from System.Drawing import Size, Point


# --- Form-Klasse ---
class ReactionForm(Form):
    def __init__(self):
        self.Text = "Reaktionen generieren"
        self.Size = Size(400,250)  # etwas höher für zweite Checkbox
        self.StartPosition = FormStartPosition.CenterScreen

        # Standardwerte für Optionen
        self.renameResults = True
        self.groupNewResults = False

        # Label
        self.label = Label()
        self.label.Text = "Optionen:"
        self.label.Location = Point(20,20)
        self.label.Size = Size(300,20)
        self.Controls.Add(self.label)

        # Checkbox 1: Reaktionen umbenennen
        self.checkBoxRename = CheckBox()
        self.checkBoxRename.Text = "Vorhandene Reaktionen umbenennen?"
        self.checkBoxRename.Location = Point(20,50)
        self.checkBoxRename.Size = Size(300,20)
        self.checkBoxRename.Checked = self.renameResults
        self.checkBoxRename.CheckedChanged += self.toggleRename
        self.Controls.Add(self.checkBoxRename)

        # Checkbox 2: Reaktionen gruppieren
        self.checkBoxGroup = CheckBox()
        self.checkBoxGroup.Text = "Neu generierte Reaktionen gruppieren?"
        self.checkBoxGroup.Location = Point(20,80)
        self.checkBoxGroup.Size = Size(350,20)
        self.checkBoxGroup.Checked = self.groupNewResults
        self.checkBoxGroup.CheckedChanged += self.toggleGroup
        self.Controls.Add(self.checkBoxGroup)

        # Button: Generiere Reaktionen
        self.generateButton = Button()
        self.generateButton.Text = "Generiere Reaktionen"
        self.generateButton.Location = Point(50,130)
        self.generateButton.Size = Size(130,30)
        self.generateButton.Click += self.startScript
        self.Controls.Add(self.generateButton)

        # Button: Abbrechen
        self.cancelButton = Button()
        self.cancelButton.Text = "Abbrechen"
        self.cancelButton.Location = Point(200,130)
        self.cancelButton.Size = Size(100,30)
        self.cancelButton.Click += self.cancelScript
        self.Controls.Add(self.cancelButton)

    # Event-Handler
    def toggleRename(self, sender, args):
        self.renameResults = self.checkBoxRename.Checked

    def toggleGroup(self, sender, args):
        self.groupNewResults = self.checkBoxGroup.Checked

    def startScript(self, sender, args):
        # hier würde dein Skript starten
        print("Starte Skript...")
        print("renameResults =", self.renameResults)
        print("groupNewResults =", self.groupNewResults)
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancelScript(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ----------------------------------------
# ----------------------------------------
# MAIN
# ----------------------------------------
# ----------------------------------------

# --------------------
# Datenpfade 
# --------------------
user_DIR = wbjn.ExecuteCommand(ExtAPI, "returnValue(GetUserFilesDirectory())")

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

form = ReactionForm()
result = form.ShowDialog()

if result == DialogResult.OK:
    renameResults = form.renameResults
    groupNewResults = form.groupNewResults

    try:
        for item in Model.Connections.Children:
            collectAllConnections(item, contactObjectList, jointObjectList, beamObjectList, connectionGroupTypes)
    except:
        print("SOMETHING WENT WRONG DURING GATHERING OF CONNECTION OBJECTS")


# --------------------

    for analysisIndex, analysis in enumerate(ExtAPI.DataModel.AnalysisList):
        print("Read analysis: " +analysis.Name)

        # Listen fuer Gruppierungen
        ContactResultList = []
        JointResultList = []
        BeamResultList = []
        BCResultList = []
        listOfGroups = []

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

        # Zur Laufzeitoptimierung bereits benutze Reaktion herausnehmen
        usedResults = []


# --------------------

        # Boundary Conditions
        boundaryObjectList = []
        try:
            for item in analysis.Children:
                collectAllBoundaries(item, boundaryObjectList, boundaryTypes)
        except:
            print("SOMETHING WENT WRONG DURING GATHERING OF BOUNDARY OBJECTS")


        for bcObject in boundaryObjectList:
            print("Read Boundary: " + bcObject.Name)
            check_for_stop(user_DIR)

            hasForce = False
            hasMoment = False

            for item in analysis.Solution.Children:
                try:
                    if item.BoundaryConditionSelection == bcObject:
                        # Force
                        if item.GetType() == resultForceType:
                            usedResults.append(item)
                            hasForce = True
                            if renameResults:
                                item.Name = bcObject.Name + " (Force)"
                                #BCResultList.append(item)

                        # Moment
                        elif item.GetType() == resultMomentType:
                            usedResults.append(item)
                            hasMoment = True
                            if renameResults:
                                item.Name = bcObject.Name + " (Moment)"
                                #BCResultList.append(item)

                except:
                    pass

                if hasForce and hasMoment:
                    break

            if not hasForce:
                forceSolution = analysis.Solution.AddForceReaction()
                forceSolution.BoundaryConditionSelection = bcObject
                forceSolution.Name = bcObject.Name + " (Force)"
                usedResults.append(forceSolution)
                BCResultList.append(forceSolution)

            if not hasMoment:
                momentSolution = analysis.Solution.AddMomentReaction()
                momentSolution.BoundaryConditionSelection = bcObject
                momentSolution.Name = bcObject.Name + " (Moment)"
                usedResults.append(momentSolution)
                BCResultList.append(momentSolution)


        # Contacts
        for contactObject in contactObjectList:
            print("Read contact: " + contactObject.Name)
            check_for_stop(user_DIR)

            hasForce = False
            hasMoment = False

            for item in analysis.Solution.Children:
                try:
                    if item.ContactRegionSelection == contactObject:
                        # Force
                        if item.GetType() == resultForceType:
                            usedResults.append(item)
                            hasForce = True
                            if renameResults:
                                #ContactResultList.append(item)
                                item.Name = contactObject.Name + " (Force)"

                        # Moment
                        elif item.GetType() == resultMomentType:
                            usedResults.append(item)
                            hasMoment = True
                            if renameResults:
                                #ContactResultList.append(item)
                                item.Name = contactObject.Name + " (Moment)"

                except:
                    pass

                if hasForce and hasMoment:
                    break

            if not hasForce:
                forceSolution = analysis.Solution.AddForceReaction()
                forceSolution.ContactRegionSelection = contactObject
                forceSolution.Name = contactObject.Name + " (Force)"
                usedResults.append(forceSolution)
                ContactResultList.append(forceSolution)

            if not hasMoment:
                momentSolution = analysis.Solution.AddMomentReaction()
                momentSolution.ContactRegionSelection = contactObject
                momentSolution.Name = contactObject.Name + " (Moment)"
                usedResults.append(momentSolution)
                ContactResultList.append(momentSolution)


        if domain == "time":
            # Joints
            for jointObject in jointObjectList:
                print("Read joint: " + jointObject.Name)
                check_for_stop(user_DIR)

                # initialize lists
                hasForce = False
                hasMoment = False

                for item in analysis.Solution.Children:
                    try:
                        if item.BoundaryConditionSelection == jointObject:
                            
                            if item.ResultType == probeForceType:
                                usedResults.append(item)
                                hasForce = True
                                if renameResults:
                                    item.Name = jointObject.Name + " (Force)"
                                    #JointResultList.append(item)

                            elif item.ResultType == probeMomentType:
                                usedResults.append(item)
                                hasMoment = True
                                if renameResults:
                                    item.Name = jointObject.Name + " (Moment)"
                                    #JointResultList.append(item)

                    except:
                        pass

                    if hasForce and hasMoment:
                        break

                if not hasForce:
                    forceSolution = analysis.Solution.AddJointProbe()
                    forceSolution.ResultType = probeForceType
                    forceSolution.Name = jointObject.Name + " (Force)"
                    forceSolution.BoundaryConditionSelection = jointObject
                    usedResults.append(forceSolution)
                    JointResultList.append(forceSolution)

                if not hasMoment:
                    momentSolution = analysis.Solution.AddJointProbe()
                    momentSolution.ResultType = probeMomentType
                    momentSolution.Name = jointObject.Name + " (Moment)"
                    momentSolution.BoundaryConditionSelection = jointObject
                    usedResults.append(momentSolution)
                    JointResultList.append(momentSolution)


            # Beams
            for beamObject in beamObjectList:
                print("Read beam: " + beamObject.Name)
                check_for_stop(user_DIR)

                hasBeamResult = False

                for item in analysis.Solution.Children:
                    try:
                        if item.BoundaryConditionSelection == beamObject:
                            hasBeamResult = True
                            usedResults.append(item)
                            if renameResults:
                                item.Name = beamObject.Name + " (Beam)"
                                #BeamResultList.append(item)

                    except:
                        pass

                    if hasBeamResult:
                        break

                if not hasBeamResult:
                    beamSolution = analysis.Solution.AddBeamProbe()
                    beamSolution.BoundaryConditionSelection = beamObject
                    beamSolution.Name = beamObject.Name + " (Beam)"
                    usedResults.append(momentSolution)
                    BeamResultList.append(beamSolution)

        elif domain == "frequency":

            # Beams
            for beamObject in beamObjectList:
                print("Read beam: " + beamObject.Name)
                check_for_stop(user_DIR)

                hasForce = False
                hasMoment = False

                for item in analysis.Solution.Children:
                    try:
                        if item.Beam == beamObject:
                            # Force
                            if item.GetType() == resultForceType:
                                usedResults.append(item)
                                hasForce = True
                                if renameResults:
                                    item.Name = beamObject.Name + " (Force)"
                                    #BeamResultList.append(item)

                            # Moment
                            elif item.GetType() == resultMomentType:
                                usedResults.append(item)
                                hasMoment = True
                                if renameResults:
                                    item.Name = beamObject.Name + " (Moment)"
                                    #BeamResultList.append(item)
                    except:
                        pass

                    if hasForce and hasMoment:
                        break

                if not hasForce:
                    forceSolution = analysis.Solution.AddForceReaction()
                    forceSolution.LocationMethod = LocationDefinitionMethod.Beam
                    forceSolution.Beam = beamObject
                    forceSolution.Name = beamObject.Name + " (Force)"
                    usedResults.append(forceSolution)
                    BeamResultList.append(forceSolution)

                if not hasMoment:
                    momentSolution = analysis.Solution.AddMomentReaction()
                    momentSolution.LocationMethod = LocationDefinitionMethod.Beam
                    momentSolution.Beam = beamObject
                    momentSolution.Name = beamObject.Name + " (Moment)"
                    usedResults.append(momentSolution)
                    BeamResultList.append(momentSolution)


        if groupNewResults:

            if not isEmpty(ContactResultList):
                contactGroupFolder = Tree.Group(ContactResultList)
                contactGroupFolder.Name = "Kontakte"
                listOfGroups.append(contactGroupFolder)

            if not isEmpty(JointResultList):
                jointGroupFolder = Tree.Group(JointResultList)
                jointGroupFolder.Name = "Joints"
                listOfGroups.append(jointGroupFolder)

            if not isEmpty(BeamResultList):
                beamGroupFolder = Tree.Group(BeamResultList)
                beamGroupFolder.Name = "Beams"
                listOfGroups.append(beamGroupFolder)

            if not isEmpty(BCResultList):
                bcGroupFolder = Tree.Group(BCResultList)
                bcGroupFolder.Name = "Randbedingungen"
                listOfGroups.append(bcGroupFolder)

            if not isEmpty(listOfGroups):
                reactionGroupFolder = Tree.Group(listOfGroups)
                reactionGroupFolder.Name = "Reaktionen"

    
else:
    print("Abgebrochen.")
    