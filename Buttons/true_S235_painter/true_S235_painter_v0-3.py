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

from System.Windows.Forms import (Application, Form, Label, ComboBox, TextBox, Button)
from System.Drawing import Point, Size
import System


# ----------------
# Class
# ----------------
class PaintDefintions:
    def __init__(self, category = [999], unit = None, maxValue = None, minValue = None, applyToAll = True, applyToSelectedAnalysis = None):
        self.category = category
        if category[0] in [DataModelObjectCategory.EquivalentStress, DataModelObjectCategory.MaximumPrincipalStress, DataModelObjectCategory.MiddlePrincipalStress, DataModelObjectCategory.MinimumPrincipalStress]:
            self.unit = "MPa"
        elif category[0] in [DataModelObjectCategory.TotalDeformation, DataModelObjectCategory.DirectionalDeformation]:
            self.unit = "mm"
        elif category[0] in [DataModelObjectCategory.EquivalentTotalStrain, DataModelObjectCategory.EquivalentPlasticStrainRST]:
            self.unit = "mm/mm"

        self.maxValue = maxValue
        self.minValue = minValue
        self.applyToAll = applyToAll
        self.applyToSelectedAnalysis = applyToSelectedAnalysis


class PaintGUI(Form):
    def __init__(self, categoryOptions, categoryMap, analysisObjects):
        self.categoryOptions = categoryOptions
        self.categoryMap = categoryMap
        self.analysisObjects = analysisObjects

        self.analysisNames = ["All"] + [a.Name for a in analysisObjects]
        self.paintDefinition = None

        self.Text = "Paint Definitions"
        self.Size = Size(360, 260)

        # Result Type
        Label(Text="Choose Result Type:", Location=Point(10, 15), Parent=self)
        self.cbResult = ComboBox(Location=Point(150, 10), Width=180)
        self.cbResult.Items.AddRange(
            System.Array[object](self.categoryOptions)
        )
        self.cbResult.SelectedIndex = 0
        self.Controls.Add(self.cbResult)

        # Analysis ComboBox
        self.cbAnalysis = ComboBox(Location=Point(150, 120), Width=180)
        self.cbAnalysis.Items.AddRange(
            System.Array[object](self.analysisNames)
        )
        self.cbAnalysis.SelectedIndex = 0
        self.Controls.Add(self.cbAnalysis)

        # Max Value
        Label(Text="Max. Value:", Location=Point(10, 55), Parent=self)
        self.tbMax = TextBox(Location=Point(150, 50), Width=180)
        self.Controls.Add(self.tbMax)

        # Min Value
        Label(Text="Min. Value:", Location=Point(10, 85), Parent=self)
        self.tbMin = TextBox(Location=Point(150, 80), Width=180)
        self.Controls.Add(self.tbMin)

        # Analysis
        Label(Text="Choose Analysis:", Location=Point(10, 125), Parent=self)
        self.cbAnalysis = ComboBox(Location=Point(150, 120), Width=180)
        self.cbAnalysis.Items.AddRange(System.Array[object](self.analysisNames))
        self.cbAnalysis.SelectedIndex = 0
        self.Controls.Add(self.cbAnalysis)

        # OK Button
        btnOK = Button(Text="OK", Location=Point(120, 170), Width=100)
        btnOK.Click += self.onOK
        self.Controls.Add(btnOK)

    def _parseFloat(self, text):
        try:
            return float(text)
        except:
            return None

    def onOK(self, sender, args):
        # --- Result Category ---
        resultType = self.cbResult.SelectedItem
        category = self.categoryMap[resultType]

        # --- Values ---
        maxValue = self._parseFloat(self.tbMax.Text)
        minValue = self._parseFloat(self.tbMin.Text)

        # --- Analysis Selection ---
        analysisSelection = self.cbAnalysis.SelectedItem

        if analysisSelection == "All":
            applyToAll = True
            applyToSelectedAnalysis = None
        else:
            applyToAll = False
            applyToSelectedAnalysis = next(
                a for a in self.analysisObjects if a.Name == analysisSelection
            )

        # --- Create PaintDefinition ---
        self.paintDefinition = PaintDefintions(
            category=category,
            maxValue=maxValue,
            minValue=minValue,
            applyToAll=applyToAll,
            applyToSelectedAnalysis=applyToSelectedAnalysis
        )

        self.Close()


# ----------------
# Functions
# ----------------
def apply_legend_Scheme(paintDef = PaintDefintions(), resultType = None):
    BANDS = 9
    legendSettings = Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings()
    legendSettings.ColorScheme = LegendColorSchemeType.Rainbow
    legendSettings.NumberOfBands = BANDS
    legendSettings.Digits = 3

    if resultType in [DataModelObjectCategory.EquivalentStress, DataModelObjectCategory.TotalDeformation, DataModelObjectCategory.EquivalentTotalStrain, DataModelObjectCategory.EquivalentPlasticStrainRST]:
        legendSettings.SetBandColor(0, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(192,192,192))

        for i in range(1, BANDS):
            legendSettings.SetBandColorAuto(i, True)

        if is_float(paintDef.maxValue):
            legendSettings.SetLowerBound(BANDS-1,Quantity(paintDef.maxValue, paintDef.unit))
        if is_float(paintDef.minValue):
            legendSettings.SetUpperBound(0,Quantity(paintDef.minValue, paintDef.unit))


    elif resultType in [DataModelObjectCategory.DirectionalDeformation]:
        
        for i in range(1, BANDS):
            legendSettings.SetBandColorAuto(i, True)

        if is_float(paintDef.maxValue):
            legendSettings.SetLowerBound(BANDS-1,Quantity(paintDef.maxValue, paintDef.unit))
        if is_float(paintDef.minValue):
            legendSettings.SetUpperBound(0,Quantity(paintDef.minValue, paintDef.unit))


    elif resultType == DataModelObjectCategory.MaximumPrincipalStress:

        legendSettings.SetBandColor(0, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(192,192,192))

        for i in range(1, BANDS):
            legendSettings.SetBandColorAuto(i, True)

        if is_float(paintDef.maxValue):
            legendSettings.SetLowerBound(BANDS-1,Quantity(paintDef.maxValue, paintDef.unit))
            legendSettings.SetUpperBound(0,Quantity(0, paintDef.unit))


    elif resultType == DataModelObjectCategory.MiddlePrincipalStress:

        legendSettings.SetBandColor(BANDS-1, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(255,0,0))
        legendSettings.SetBandColor(int((BANDS-1)/2.0), Ansys.Mechanical.DataModel.Utilities.Colors.RGB(192,192,192))
        legendSettings.SetBandColor(0, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(0,0,255))

        for i in range(1, BANDS-1):
            legendSettings.SetBandColorAuto(i, True)

        if is_float(paintDef.maxValue):
            legendSettings.SetLowerBound(BANDS-1,Quantity(paintDef.maxValue, paintDef.unit))
        if is_float(paintDef.minValue):
            legendSettings.SetUpperBound(0,Quantity(paintDef.minValue, paintDef.unit))


    elif resultType == DataModelObjectCategory.MinimumPrincipalStress:
        legendSettings.ColorScheme = LegendColorSchemeType.ReverseRainbow
        legendSettings.SetBandColor(BANDS-1, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(192,192,192))
        legendSettings.SetBandColor(0, Ansys.Mechanical.DataModel.Utilities.Colors.RGB(255,0,0))

        for i in range(1, BANDS-1):
            legendSettings.SetBandColorAuto(i, True)

        if is_float(paintDef.minValue):
            legendSettings.SetLowerBound(BANDS-1, Quantity(0, paintDef.unit))
            legendSettings.SetUpperBound(0, Quantity(paintDef.minValue, paintDef.unit))


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

def get_maxValue_and_scope():
    class MaxInputForm(Form):
        def __init__(self):
            self.Text = "Bemessungsgrenze"
            self.Size = Size(320,200)
            self.StartPosition = FormStartPosition.CenterScreen

            # Label für maxValue
            self.label = Label()
            self.label.Text = "Bitte den maximalen Bemessungswert eingeben:"
            self.label.Location = Point(10,20)
            self.label.Size = Size(280,20)
            self.Controls.Add(self.label)

            # TextBox für maxValue
            self.textbox = TextBox()
            self.textbox.Location = Point(10,45)
            self.textbox.Size = Size(280,20)
            self.Controls.Add(self.textbox)

            # CheckBox für Scope-Auswahl
            self.checkbox = CheckBox()
            self.checkbox.Text = "Alle Ergebnisse anpassen"
            self.checkbox.Location = Point(10,80)
            self.checkbox.Size = Size(280,20)
            self.Controls.Add(self.checkbox)

            # OK Button
            self.okButton = Button()
            self.okButton.Text = "OK"
            self.okButton.Location = Point(110,120)
            self.okButton.DialogResult = DialogResult.OK
            self.Controls.Add(self.okButton)

            self.AcceptButton = self.okButton

    form = MaxInputForm()
    if form.ShowDialog() == DialogResult.OK:
        try:
            max_val = float(form.textbox.Text)
        except:
            max_val = None

        applyToAll = form.checkbox.Checked
        return max_val, applyToAll

    return None, None

# ------
# MAIN
# ------

# --------------------------------------------------
# Daten vorbereiten
# --------------------------------------------------
categoryOptions = [
    "Equivalent Stress",
    "Principle Stresses",
    "Total Deformation",
    "Directional Deformations",
    "Total Strain",
    "Plastic Strain"
]

categoryMap = {
    "Equivalent Stress": [DataModelObjectCategory.EquivalentStress],
    "Principle Stresses": [
        DataModelObjectCategory.MaximumPrincipalStress,
        DataModelObjectCategory.MiddlePrincipalStress,
        DataModelObjectCategory.MinimumPrincipalStress
    ],
    "Total Deformation": [DataModelObjectCategory.TotalDeformation],
    "Directional Deformations": [DataModelObjectCategory.DirectionalDeformation],
    "Total Strain": [DataModelObjectCategory.EquivalentTotalStrain],
    "Plastic Strain": [DataModelObjectCategory.EquivalentPlasticStrainRST]
}

analysisObjects = []
for analysis in ExtAPI.DataModel.AnalysisList:
    analysisObjects.append(analysis)

for item in Model.Children:
    if item.GetType() == Ansys.ACT.Automation.Mechanical.SolutionCombination:
        analysisObjects.append(item)

# --------------------------------------------------
# GUI starten
# --------------------------------------------------
form = PaintGUI(categoryOptions, categoryMap, analysisObjects)
Application.Run(form)

userDefinitions = form.paintDefinition


# ------- RUN -----------

user_DIR = wbjn.ExecuteCommand(ExtAPI,"returnValue(GetUserFilesDirectory())")

if userDefinitions.applyToAll:
    for analysis in analysisObjects:
        if analysis.GetType() == Ansys.ACT.Automation.Mechanical.Analysis:
            for result in analysis.Solution.Children:
                if result.DataModelObjectCategory in userDefinitions.category:
                    check_for_stop(user_DIR)
                    result.Activate()
                    apply_legend_Scheme(userDefinitions, result.DataModelObjectCategory)

        elif analysis.GetType() == Ansys.ACT.Automation.Mechanical.SolutionCombination:
            for result in analysis.Children:
                if result.DataModelObjectCategory in userDefinitions.category:
                    check_for_stop(user_DIR)
                    result.Activate()
                    apply_legend_Scheme(userDefinitions, result.DataModelObjectCategory)


elif userDefinitions.applyToSelectedAnalysis:
        analysis = userDefinitions.applyToSelectedAnalysis
        if analysis.GetType() == Ansys.ACT.Automation.Mechanical.Analysis:
            for result in analysis.Solution.Children:
                if result.DataModelObjectCategory in userDefinitions.category:
                    check_for_stop(user_DIR)
                    result.Activate()
                    apply_legend_Scheme(userDefinitions, result.DataModelObjectCategory)

        elif analysis.GetType() == Ansys.ACT.Automation.Mechanical.SolutionCombination:
            for result in analysis.Children:
                if result.DataModelObjectCategory in userDefinitions.category:
                    check_for_stop(user_DIR)
                    result.Activate()
                    apply_legend_Scheme(userDefinitions, result.DataModelObjectCategory)


    
