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
import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, Application, Button, Label, CheckBox, FormStartPosition
from System.Drawing import Size, Point

# ----------------
# Einstellungen
# ----------------
seperator = ","

# --- Form-Klasse ---
class PretensionForm(Form):
    def __init__(self, nPretensions):
        self.Text = "Pretension einstellen"
        self.Size = Size(400,200)
        self.StartPosition = FormStartPosition.CenterScreen

        self.pretensionForceValue = None

        # Label: Info
        self.labelInfo = Label()
        self.labelInfo.Text = "Gefundene Pretensions: " + str(nPretensions)
        self.labelInfo.Location = Point(20,20)
        self.labelInfo.Size = Size(350,20)
        self.Controls.Add(self.labelInfo)

        # Label: Eingabe
        self.labelInput = Label()
        self.labelInput.Text = "Setze Vorspannung auf [N]:"
        self.labelInput.Location = Point(20,50)
        self.labelInput.Size = Size(200,20)
        self.Controls.Add(self.labelInput)

        # TextBox: Eingabe
        self.textBoxForce = TextBox()
        self.textBoxForce.Location = Point(220,50)
        self.textBoxForce.Size = Size(100,20)
        self.Controls.Add(self.textBoxForce)

        # Button: OK
        self.okButton = Button()
        self.okButton.Text = "OK"
        self.okButton.Location = Point(50,100)
        self.okButton.Size = Size(100,30)
        self.okButton.Click += self.onOK
        self.Controls.Add(self.okButton)

        # Button: Abbrechen
        self.cancelButton = Button()
        self.cancelButton.Text = "Abbrechen"
        self.cancelButton.Location = Point(200,100)
        self.cancelButton.Size = Size(100,30)
        self.cancelButton.Click += self.onCancel
        self.Controls.Add(self.cancelButton)

    def onOK(self, sender, args):
        try:
            # Float aus TextBox auslesen
            value = float(self.textBoxForce.Text)
            self.pretensionForceValue = str(value)
            self.DialogResult = DialogResult.OK
            self.Close()
        except:
            from System.Windows.Forms import MessageBox
            MessageBox.Show("Bitte eine gültige Zahl eingeben.", "Fehler")
            return

    def onCancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# --------------------
# MAIN
# --------------------

# Anzahl aller Pretensions ermitteln
totalPretensions = 0
for analysis in ExtAPI.DataModel.AnalysisList:
    pretensions = analysis.GetChildren(DataModelObjectCategory.BoltPretension, True)
    totalPretensions += len(pretensions)

# GUI starten
form = PretensionForm(totalPretensions)
result = form.ShowDialog()

if result == DialogResult.OK and form.pretensionForceValue is not None:
    pretensionForceValue = form.pretensionForceValue
    pretensionForce = Quantity(pretensionForceValue + " [N]")

    # Alle Analysen durchlaufen
    for analysisIndex, analysis in enumerate(ExtAPI.DataModel.AnalysisList):
        print("Read analysis: " + analysis.Name)

        # Alle Pretensions in der aktuellen Analyse
        pretensions = analysis.GetChildren(DataModelObjectCategory.BoltPretension, True)
        print("Gefundene Pretensions:", len(pretensions))

        # Pretensions durchgehen
        for pretensionObject in pretensions:

            # Dynamische Anzahl Steps ermitteln
            nSteps = len(pretensionObject.Preload.Output.DiscreteValues)
            if nSteps < 1:
                continue

            # DiscreteValues automatisch erzeugen: Step 1 = pretensionForce, Rest = 0[N]
            newValues = [pretensionForce] + [Quantity("0 [N]")] * (nSteps - 1)
            pretensionObject.Preload.Output.DiscreteValues = newValues

            # Alle weiteren Steps auf Locked setzen
            for i in range(1, nSteps):
                pretensionObject.SetDefineBy(i + 1, BoltLoadDefineBy.Lock)

            print("Pretension aktualisiert:", pretensionObject.Name)

    print("=== Script abgeschlossen ===")
else:
    print("Abbruch durch Benutzer.")