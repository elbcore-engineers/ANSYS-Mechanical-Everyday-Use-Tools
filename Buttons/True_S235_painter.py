# Script

#maxValue = 235

# ----------------
# Importiere Module
# ----------------
import os
import sys
import wbjn
import context_menu

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Application, Form, Label, TextBox, Button, DialogResult, FormStartPosition
from System.Drawing import Point, Size

# ----------------
# Functions
# ----------------
def apply_legend_Scheme(maxValue):
    BANDS = 9
    legendSettings = Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings()
    legendSettings.Reset()
    legendSettings.NumberOfBands = BANDS
    legendSettings.ColorScheme = LegendColorSchemeType.Rainbow
    legendSettings.SetBandColor(0, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(192,192,192))
    if is_float(maxValue):
        legendSettings.SetLowerBound(BANDS-1,Quantity(maxValue, "MPa"))
        legendSettings.SetUpperBound(BANDS-1,Quantity(maxValue*1.0000000001, "MPa"))
    if is_float(minValue):
        legendSettings.SetLowerBound(0,Quantity(minValue, "MPa"))
    
    for i in range(1, BANDS):
        legendSettings.SetBandColorAuto(i, True)


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

def is_float(value):
    try:
        float(value)
        return True
    except:
        return False


def get_maxValue():
    class MaxInputForm(Form):
        def __init__(self):
            self.Text = "Bemessungsgrenze"
            self.Size = Size(300,150)
            self.StartPosition = FormStartPosition.CenterScreen

            # Label
            self.label = Label()
            self.label.Text = "Bitte den maximalen Bemessungswert eingeben:"
            self.label.Location = Point(10,20)
            self.label.Size = Size(260,20)
            self.Controls.Add(self.label)

            # TextBox
            self.textbox = TextBox()
            self.textbox.Location = Point(10,45)
            self.textbox.Size = Size(260,20)
            self.Controls.Add(self.textbox)

            # OK Button
            self.okButton = Button()
            self.okButton.Text = "OK"
            self.okButton.Location = Point(100,80)
            self.okButton.DialogResult = DialogResult.OK
            self.Controls.Add(self.okButton)

            self.AcceptButton = self.okButton

    form = MaxInputForm()
    if form.ShowDialog() == DialogResult.OK:
        try:
            return float(form.textbox.Text)
        except:
            return None
    return None

# ------
# MAIN
# ------

# --- Eingabe einholen ---
maxValue = get_maxValue()

# ------
# Parameter
# ------

#stressCategory = [DataModelObjectCategory.EquivalentStress, DataModelObjectCategory.MaximumPrincipalStress, DataModelObjectCategory.MiddlePrincipalStress, DataModelObjectCategory.MinimumPrincipalStress]
stressCategory = [DataModelObjectCategory.EquivalentStress]

user_DIR = wbjn.ExecuteCommand(ExtAPI,"returnValue(GetUserFilesDirectory())")

# ------
# Analysis
# ------
for analysis in ExtAPI.DataModel.AnalysisList:

    #Get All Stress Objects in all the Analyses in the Tree
    analstressResults = [child for child in analysis.Solution.Children if child.DataModelObjectCategory in stressCategory]

    for result in analstressResults:
        check_for_stop(user_DIR)
        result.Activate()
        apply_legend_Scheme(maxValue)


# ------
# Exception for Solution Combination
# ------
for item in Model.Children:
 
    if item.GetType() == Ansys.ACT.Automation.Mechanical.SolutionCombination:

        #Get All Stress Objects in all the Analyses in the Tree
        scStressResults = [child for child in item.Children if child.DataModelObjectCategory in stressCategory]

        for result in scStressResults:
            check_for_stop(user_DIR)
            result.Activate()
            apply_legend_Scheme(maxValue)
