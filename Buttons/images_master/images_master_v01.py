# Script

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
    text = text.replace("   "," ")
    text = text.replace("  "," ")
    text = text.replace(" ","-")
    text = text.replace("/",":")
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

# ------
# Paths
# ------

# Get the User Files Directory of the Project
user_DIR = wbjn.ExecuteCommand(ExtAPI,"returnValue(GetUserFilesDirectory())")

# Get the directory to store the images
model_DIR = makeFolder(os.path.join(user_DIR,"Result_Images"))

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