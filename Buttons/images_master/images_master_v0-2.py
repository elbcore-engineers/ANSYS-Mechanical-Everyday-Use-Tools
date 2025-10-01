# Script

# ----------------
# Importiere Module
# ----------------
import os, sys
import clr
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

# --- Form-Klasse ---
class PathSelectionForm(Form):
    def __init__(self, default_path):
        self.Text = "Bilder-Pfad auswählen"
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


# ------
# Paths
# ------

# --- Projekt-spezifisch: UserFiles-Verzeichnis holen ---
#user_DIR = wbjn.ExecuteCommand(ExtAPI,"returnValue(GetUserFilesDirectory())")
projectDIR = ExtAPI.DataModel.Project.ProjectDirectory
user_DIR = os.path.join(projectDIR, "exports")

# Default-Ordnername mit Zeitstempel
date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
default_model_DIR = os.path.join(user_DIR, "Images_" + date)

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

# ------
# Graphic settings
# ------
gset = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
gset.Background = GraphicsBackgroundType.White
gset.CurrentGraphicsDisplay = False
gset.Width = 1920
gset.Height = 1080

# ------
# Store Mesh
# ------
Model.Mesh.Activate()
fName = "mesh.png"
fPath = os.path.join(model_DIR, fName)
Graphics.ExportImage(fPath, GraphicsImageExportFormat.PNG, gset)

stressCategory = [DataModelObjectCategory.EquivalentStress, DataModelObjectCategory.MaximumPrincipalStress, DataModelObjectCategory.MiddlePrincipalStress, DataModelObjectCategory.MinimumPrincipalStress]

strainCategory = [DataModelObjectCategory.EquivalentElasticStrainRST, DataModelObjectCategory.EquivalentPlasticStrainRST, DataModelObjectCategory.EquivalentTotalStrain]

deformationCategory = [DataModelObjectCategory.TotalDeformation, DataModelObjectCategory.DirectionalDeformation, DataModelObjectCategory.VectorDeformation]

# ------
# Store Results
# ------
for analysis in ExtAPI.DataModel.AnalysisList:
        
    analysis_DIR = makeFolder(os.path.join(model_DIR, adjustPath(analysis.Name)))

    #Get All Stress Objects in all the Analyses in the Tree
    analstressResults = [child for child in analysis.Solution.Children if child.DataModelObjectCategory in stressCategory]

    #Get All Strain Objects in all the Analyses in the Tree
    analStrainResults = [child for child in analysis.Solution.Children if child.DataModelObjectCategory in strainCategory]

    #Get All Deformation Objects in all the Analyses in the Tree
    analDeformationResults = [child for child in analysis.Solution.Children if child.DataModelObjectCategory in deformationCategory]
    
    analAllResults = analstressResults + analStrainResults + analDeformationResults
    for result in analAllResults:
        check_for_stop(user_DIR)

        result.Activate()
        fName = adjustPath(result.Name) + ".png"
        fPath = os.path.join(analysis_DIR, fName)
        if os.path.exists(fPath):
            os.remove(fPath)
        Graphics.ExportImage(fPath, GraphicsImageExportFormat.PNG, gset)

# ------
# Exception for Solution Combination
# ------
for item in Model.Children:
 
    if item.GetType() == Ansys.ACT.Automation.Mechanical.SolutionCombination:

        item_DIR = makeFolder(os.path.join(model_DIR, adjustPath(item.Name)))

        #Get All Stress Objects in all the Analyses in the Tree
        scStressResults = [child for child in item.Children if child.DataModelObjectCategory in stressCategory]

        #Get All Strain Objects in all the Analyses in the Tree
        scStrainResults = [child for child in item.Children if child.DataModelObjectCategory in strainCategory]

        #Get All Deformation Objects in all the Analyses in the Tree
        scDeformationResults = [child for child in item.Children if child.DataModelObjectCategory in deformationCategory]

        scAllResults = scStressResults + scStrainResults + scDeformationResults
        for result in scAllResults:
            check_for_stop(user_DIR)

            result.Activate()
            fName = adjustPath(result.Name) + ".png"
            fPath = os.path.join(item_DIR, fName)
            if os.path.exists(fPath):
                os.remove(fPath)
            Graphics.ExportImage(fPath, GraphicsImageExportFormat.PNG, gset)