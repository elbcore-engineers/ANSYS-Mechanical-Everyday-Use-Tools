# Script
# Einheitensystem: s,t,mm,N

# ----------------
# Importiere Module
# ----------------
import os
import sys
import context_menu
import subprocess
import csv
import math
from datetime import datetime
import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")


from System.Windows.Forms import *
from System.Drawing import *



# ----------------
# Einstellungen
# ----------------
seperator = ","

class metric_bolt:
    """
    Klasse zur Beschreibung einer metrischen Schraube (nach VDI 2230 Anlehnung)
    Kompatibel zu IronPython (z. B. in Ansys Mechanical)
    """

    # ------------------------------------------------------------
    # Dictionaries mit Daten nach ISO-Regelgewinden
    # ------------------------------------------------------------

    bolt_sizes = {
        "M39": 39.0, "M36": 36.0, "M33": 33.0, "M30": 30.0, "M27": 27.0,
        "M24": 24.0, "M22": 22.0, "M20": 20.0, "M18": 18.0, "M16": 16.0,
        "M14": 14.0, "M12": 12.0, "M10": 10.0, "M8": 8.0, "M7": 7.0,
        "M6": 6.0, "M5": 5.0, "M4": 4.0
    }

    thread_pitch = {
        "M39": 4.00, "M36": 4.00, "M33": 3.50, "M30": 3.50, "M27": 3.00,
        "M24": 3.00, "M22": 2.50, "M20": 2.50, "M18": 2.50, "M16": 2.00,
        "M14": 2.00, "M12": 1.75, "M10": 1.50, "M8": 1.25, "M7": 1.00,
        "M6": 1.00, "M5": 0.80, "M4": 0.70
    }

    stress_areas = {
        "M39": 976.00, "M36": 817.00, "M33": 694.00, "M30": 561.00, "M27": 459.00,
        "M24": 353.00, "M22": 303.00, "M20": 245.00, "M18": 193.00, "M16": 157.00,
        "M14": 115.00, "M12": 84.30, "M10": 58.00, "M8": 36.60, "M7": 28.90,
        "M6": 20.10, "M5": 14.20, "M4": 8.78
    }

    yield_strengths = {
        "12.9": 1080, "10.9": 900, "8.8": 640
    }

    tensile_strengths = {
        "12.9": 1200, "10.9": 1000, "8.8": 800
    }

    shear_coefficient = {
        "12.9": 0.5, "10.9": 0.5, "8.8": 0.6
    }

    d2_table = {
        "M39": 36.402, "M36": 33.402, "M33": 30.727, "M30": 27.051,
        "M27": 25.051, "M24": 22.051, "M22": 20.376, "M20": 18.376,
        "M18": 16.376, "M16": 14.701, "M14": 12.701, "M12": 10.863,
        "M10": 9.026, "M8": 7.188, "M7": 6.350, "M6": 5.350,
        "M5": 4.480, "M4": 3.545
    }

    A0_table = {
        "M39": 913.0, "M36": 759.3, "M33": 647.2, "M30": 519.0,
        "M27": 427.1, "M24": 324.3, "M22": 281.5, "M20": 225.2,
        "M18": 175.1, "M16": 144.1, "M14": 104.7, "M12": 76.25,
        "M10": 52.3, "M8": 32.84, "M7": 26.18, "M6": 17.89,
        "M5": 12.69, "M4": 7.749
    }

    preload_table = {
        "8.8": {
            "M39": {0.08:548.0, 0.10:537.0, 0.12:525.0, 0.14:512.0, 0.16:498.0, 0.20:470.0, 0.24:443.0},
            "M36": {0.08:458.0, 0.10:448.0, 0.12:438.0, 0.14:427.0, 0.16:415.0, 0.20:392.0, 0.24:368.0},
            "M33": {0.08:389.0, 0.10:381.0, 0.12:373.0, 0.14:363.0, 0.16:354.0, 0.20:334.0, 0.24:314.0},
            "M30": {0.08:313.0, 0.10:307.0, 0.12:300.0, 0.14:292.0, 0.16:284.0, 0.20:268.0, 0.24:252.0},
            "M27": {0.08:257.0, 0.10:252.0, 0.12:246.0, 0.14:240.0, 0.16:234.0, 0.20:220.0, 0.24:207.0},
            "M24": {0.08:196.0, 0.10:192.0, 0.12:188.0, 0.14:183.0, 0.16:178.0, 0.20:168.0, 0.24:157.0},
            "M22": {0.08:170.0, 0.10:166.0, 0.12:162.0, 0.14:158.0, 0.16:154.0, 0.20:145.0, 0.24:137.0},
            "M20": {0.08:136.0, 0.10:134.0, 0.12:130.0, 0.14:127.0, 0.16:123.0, 0.20:116.0, 0.24:109.0},
            "M18": {0.08:107.0, 0.10:104.0, 0.12:102.0, 0.14:99.0, 0.16:96.0, 0.20:91.0, 0.24:85.0},
            "M16": {0.08:84.7, 0.10:82.9, 0.12:80.9, 0.14:78.8, 0.16:76.6, 0.20:72.2, 0.24:67.8},
            "M14": {0.08:62.0, 0.10:60.6, 0.12:59.1, 0.14:57.5, 0.16:55.9, 0.20:52.6, 0.24:49.3},
            "M12": {0.08:45.2, 0.10:44.1, 0.12:43.0, 0.14:41.9, 0.16:40.7, 0.20:38.3, 0.24:35.9},
            "M10": {0.08:31.0, 0.10:30.3, 0.12:29.6, 0.14:28.8, 0.16:27.9, 0.20:26.3, 0.24:24.7},
            "M8":  {0.08:19.5, 0.10:19.1, 0.12:18.6, 0.14:18.1, 0.16:17.6, 0.20:16.5, 0.24:15.5},
            "M7":  {0.08:15.5, 0.10:15.1, 0.12:14.8, 0.14:14.4, 0.16:14.0, 0.20:13.1, 0.24:12.3},
            "M6":  {0.08:10.7, 0.10:10.4, 0.12:10.2, 0.14:9.9, 0.16:9.6, 0.20:9.0, 0.24:8.4},
            "M5":  {0.08:7.6,  0.10:7.4, 0.12:7.2, 0.14:7.0, 0.16:6.8, 0.20:6.4, 0.24:6.0},
            "M4":  {0.08:4.6,  0.10:4.5, 0.12:4.4, 0.14:4.3, 0.16:4.2, 0.20:3.9, 0.24:3.7}
        },
        "10.9": {
            "M39": {0.08:781.0, 0.10:765.0, 0.12:748.0, 0.14:729.0, 0.16:710.0, 0.20:670.0, 0.24:630.0},
            "M36": {0.08:652.0, 0.10:638.0, 0.12:623.0, 0.14:608.0, 0.16:591.0, 0.20:558.0, 0.24:524.0},
            "M33": {0.08:554.0, 0.10:543.0, 0.12:531.0, 0.14:517.0, 0.16:504.0, 0.20:475.0, 0.24:447.0},
            "M30": {0.08:446.0, 0.10:437.0, 0.12:427.0, 0.14:416.0, 0.16:405.0, 0.20:382.0, 0.24:359.0},
            "M27": {0.08:367.0, 0.10:359.0, 0.12:351.0, 0.14:342.0, 0.16:333.0, 0.20:314.0, 0.24:295.0},
            "M24": {0.08:280.0, 0.10:274.0, 0.12:267.0, 0.14:260.0, 0.16:253.0, 0.20:239.0, 0.24:224.0},
            "M22": {0.08:242.0, 0.10:237.0, 0.12:231.0, 0.14:225.0, 0.16:219.0, 0.20:207.0, 0.24:194.0},
            "M20": {0.08:194.0, 0.10:190.0, 0.12:186.0, 0.14:181.0, 0.16:176.0, 0.20:166.0, 0.24:156.0},
            "M18": {0.08:152.0, 0.10:149.0, 0.12:145.0, 0.14:141.0, 0.16:137.0, 0.20:129.0, 0.24:121.0},
            "M16": {0.08:124.4, 0.10:121.7, 0.12:118.8, 0.14:115.7, 0.16:112.6, 0.20:106.1, 0.24:99.6},
            "M14": {0.08:91.0,  0.10:88.9,  0.12:86.7, 0.14:84.4, 0.16:82.1, 0.20:77.2, 0.24:72.5},
            "M12": {0.08:66.3,  0.10:64.8,  0.12:63.2, 0.14:61.5, 0.16:59.8, 0.20:56.3, 0.24:52.8},
            "M10": {0.08:45.6,  0.10:44.5,  0.12:43.4, 0.14:42.2, 0.16:41.0, 0.20:38.6, 0.24:36.2},
            "M8":  {0.08:28.7,  0.10:28.0,  0.12:27.3, 0.14:26.6, 0.16:25.8, 0.20:24.3, 0.24:22.7},
            "M7":  {0.08:22.7,  0.10:22.5,  0.12:21.7, 0.14:21.1, 0.16:20.5, 0.20:19.3, 0.24:18.1},
            "M6":  {0.08:15.7,  0.10:15.3,  0.12:14.9, 0.14:14.5, 0.16:14.1, 0.20:13.2, 0.24:12.4},
            "M5":  {0.08:11.1,  0.10:10.8,  0.12:10.6, 0.14:10.3, 0.16:10.0, 0.20:9.4,  0.24:8.8},
            "M4":  {0.08:6.8,   0.10:6.7,   0.12:6.5,  0.14:6.3,  0.16:6.1,  0.20:5.7,  0.24:5.4}
        },
        "12.9": {
            "M39": {0.08:914.0, 0.10:895.0, 0.12:875.0, 0.14:853.0, 0.16:831.0, 0.20:784.0, 0.24:738.0},
            "M36": {0.08:763.0, 0.10:747.0, 0.12:729.0, 0.14:711.0, 0.16:692.0, 0.20:653.0, 0.24:614.0},
            "M33": {0.08:649.0, 0.10:635.0, 0.12:621.0, 0.14:605.0, 0.16:589.0, 0.20:556.0, 0.24:523.0},
            "M30": {0.08:522.0, 0.10:511.0, 0.12:499.0, 0.14:487.0, 0.16:474.0, 0.20:447.0, 0.24:420.0},
            "M27": {0.08:429.0, 0.10:420.0, 0.12:410.0, 0.14:400.0, 0.16:389.0, 0.20:367.0, 0.24:345.0},
            "M24": {0.08:327.0, 0.10:320.0, 0.12:313.0, 0.14:305.0, 0.16:296.0, 0.20:279.0, 0.24:262.0},
            "M22": {0.08:283.0, 0.10:277.0, 0.12:271.0, 0.14:264.0, 0.16:257.0, 0.20:242.0, 0.24:228.0},
            "M20": {0.08:227.0, 0.10:223.0, 0.12:217.0, 0.14:212.0, 0.16:206.0, 0.20:194.0, 0.24:182.0},
            "M18": {0.08:178.0, 0.10:174.0, 0.12:170.0, 0.14:165.0, 0.16:160.0, 0.20:151.0, 0.24:142.0},
            "M16": {0.08:145.5, 0.10:142.4, 0.12:139.0, 0.14:135.4, 0.16:131.7, 0.20:124.1, 0.24:116.6},
            "M14": {0.08:106.5, 0.10:104.1, 0.12:101.5, 0.14:98.8, 0.16:96.0, 0.20:90.4, 0.24:84.8},
            "M12": {0.08:77.6,  0.10:75.9,  0.12:74.0,  0.14:72.0,  0.16:70.0,  0.20:65.8,  0.24:61.8},
            "M10": {0.08:53.3,  0.10:52.1,  0.12:50.8,  0.14:49.4,  0.16:48.0,  0.20:45.2,  0.24:42.4},
            "M8":  {0.08:33.6,  0.10:32.8,  0.12:32.0,  0.14:31.1,  0.16:30.2,  0.20:28.4,  0.24:26.6},
            "M7":  {0.08:26.6,  0.10:26.0,  0.12:25.4,  0.14:24.7,  0.16:24.0,  0.20:22.6,  0.24:21.2},
            "M6":  {0.08:18.4,  0.10:17.9,  0.12:17.5,  0.14:17.0,  0.16:16.5,  0.20:15.5,  0.24:14.5},
            "M5":  {0.08:13.0,  0.10:12.7,  0.12:12.4,  0.14:12.0,  0.16:11.7,  0.20:11.0,  0.24:10.3},
            "M4":  {0.08:8.0,   0.10:7.8,   0.12:7.6,   0.14:7.4,   0.16:7.1,   0.20:6.7,   0.24:6.3}
        }
    }

    # ------------------------------------------------------------

    def __init__(self, size=None, strengthGrade=None, friction=0.12,
                 preload=None, torque=None, Pitch=None,
                 stressAreas=None, diameter=None, utilizationFactor=0.9):

        self.size = size
        self.strengthGrade = strengthGrade
        self.friction = friction
        self.preload = preload
        self.torque = torque
        self.Pitch = Pitch
        self.stressAreas = stressAreas
        self.diameter = diameter
        self.utilizationFactor = utilizationFactor
        self.Rp02 = None
        self.Rm = None
        self.shearCoeff = None

        if size:
            self.setSize(size)
        if strengthGrade:
            self.setStrength(strengthGrade)
        if Pitch:
            self.setPitch(Pitch)

    # ------------------------------------------------------------

    def setSize(self, size):
        """Setzt Parameter, die von der Schraubengröße abhängen"""
        if size not in self.bolt_sizes:
            raise ValueError("Unbekannte Größe: " + str(size))
        self.size = size
        self.stressAreas = self.stress_areas[size]
        self.diameter = self.bolt_sizes[size]

        # Automatische Regelgewinde-Steigung
        if size in self.thread_pitch:
            self.Pitch = self.thread_pitch[size]

    def setStrength(self, grade):
        """Setzt Materialkennwerte basierend auf der Festigkeitsklasse"""
        if grade not in self.yield_strengths:
            raise ValueError("Unbekannte Festigkeitsklasse: " + str(grade))
        self.strengthGrade = grade
        self.Rp02 = self.yield_strengths[grade]
        self.Rm = self.tensile_strengths[grade]
        self.shearCoeff = self.shear_coefficient[grade]

    def setPitch(self, pitch):
        """Setzt das Gewindesteigungsmaß (manuell oder überschreibt Regelgewinde)"""
        self.Pitch = pitch

    def setUtilizationFactor(self, utilizationFactor):
        """Setzt Auslastung"""
        self.utilizationFactor = utilizationFactor

    def calc_preload(self):
        """
        Berechnet die Vorspannkraft nach vereinfachter VDI-2230-Formel.
        """
        if not (self.size and self.utilizationFactor and self.Rp02 and self.friction and self.Pitch):
            raise ValueError("Nicht alle Parameter für die Berechnung der Vorspannkraft sind gesetzt.")

        v = self.utilizationFactor

        if v == 0.9:
            self.preload = self.calc_preload_from_table()
        else:
            Rp02 = self.Rp02
            mu = self.friction
            P = self.Pitch

            # Werte aus den eigenen Dictionaries
            if self.size not in self.d2_table or self.size not in self.stress_areas:
                raise ValueError("Größe nicht in Tabelle gefunden: " + str(self.size))
            d2 = self.d2_table[self.size]
            As = self.stress_areas[self.size] # Annahme A0 = As

            # Kerndurchmesser ds aus As berechnen
            ds = math.sqrt(4 * As/ math.pi)

            term = (1.5 * d2 / ds * (P / (math.pi * d2) + 1.155 * mu))
            self.preload = As * v * Rp02 / math.sqrt(1 + 3 * term**2)
            self.preload = int(self.preload/1000)
            self.preload = self.preload*1000

        return self.preload

    def calc_preload_from_table(self):
        """Gibt die Vorspannkraft in N basierend auf exakten Tabellenwert zurück."""
        # Prüfen, ob Größe und Festigkeit in der Tabelle existieren
        if self.strengthGrade not in self.preload_table:
            raise ValueError("Strength grade nicht in Tabelle")
        if self.size not in self.preload_table[self.strengthGrade]:
            raise ValueError("Bolt size nicht in Tabelle")
        if self.friction not in self.preload_table[self.strengthGrade][self.size]:
            raise ValueError("Reibwert nicht in Tabelle")
        
        # Exakten Tabellenwert auslesen
        FkN = self.preload_table[self.strengthGrade][self.size][self.friction]
        return FkN * 1000  # Umrechnung in N


    def __repr__(self):
        s = "metric_bolt("
        s += "size=" + str(self.size)
        s += ", diameter=" + str(self.diameter)
        s += ", stress area=" + str(self.stressAreas)
        s += ", strengthGrade=" + str(self.strengthGrade)
        s += ", friction=" + str(self.friction)
        s += ", preload=" + str(round(self.preload, 2) if self.preload else "None")
        s += ", Pitch=" + str(self.Pitch)
        s += ", utilizationFactor=" + str(self.utilizationFactor)
        s += ")"
        return s

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

# --- Hilfsfunktion: Größe und Festigkeit aus Text extrahieren ---
def get_bolt_data(name, bolt_sizes, yield_strengths):
    size = "Default"
    strength = "Default"

    # Zuerst nach Bolt-Size suchen
    for key in bolt_sizes.keys():
        if key in name:
            size = key
            break

    # Dann nach Festigkeitsklasse suchen
    for key in yield_strengths.keys():
        if key in name:
            strength = key
            break

    return (size, strength)


# --- GUI Klasse ---
class PretensionForm(Form):
    def __init__(self, pretensionNames):
        Form.__init__(self) # WICHTIG: ruft den Form-Konstruktor auf

        self.Text = "Pretension einstellen"
        self.Size = Size(800, 800)
        self.StartPosition = FormStartPosition.CenterScreen

        self.pretensionNames = pretensionNames
        self.defaultSize = "M12"
        self.defaultStrength = "8.8"
        self.defaultFriction = 0.14
        self.defaultUtilization = 0.9
        self.manualPretensionForce = None
        self.groupTable = []

        # --- Default-Werte Controls ---
        y = 20
        lbDefault = Label()
        lbDefault.Text = "Default Werte:"
        lbDefault.Location = Point(20, y)
        lbDefault.Size = Size(150, 20)
        self.Controls.Add(lbDefault)

        y += 30
        lblSize = Label()
        lblSize.Text = "Nenndurchmesser:"
        lblSize.Location = Point(20, y)
        lblSize.Size = Size(150, 20)
        self.Controls.Add(lblSize)

        self.comboSize = ComboBox()
        self.comboSize.Location = Point(180, y)
        self.comboSize.Size = Size(120, 20)
        self.comboSize.DropDownStyle = ComboBoxStyle.DropDownList
        for s in ["M39","M36","M33","M30","M27","M24","M22","M20","M18","M16","M14","M12","M10","M8","M7","M6","M5","M4"]:
            self.comboSize.Items.Add(s)
        self.comboSize.SelectedItem = self.defaultSize
        self.Controls.Add(self.comboSize)

        y += 30
        lblStrength = Label()
        lblStrength.Text = "Festigkeitsklasse:"
        lblStrength.Location = Point(20, y)
        lblStrength.Size = Size(150, 20)
        self.Controls.Add(lblStrength)

        self.comboStrength = ComboBox()
        self.comboStrength.Location = Point(180, y)
        self.comboStrength.Size = Size(120, 20)
        self.comboStrength.DropDownStyle = ComboBoxStyle.DropDownList
        for s in metric_bolt.yield_strengths.keys():
            self.comboStrength.Items.Add(s)
        self.comboStrength.SelectedItem = self.defaultStrength
        self.Controls.Add(self.comboStrength)

        y += 30
        lblFriction = Label()
        lblFriction.Text = "Gewindereibung:"
        lblFriction.Location = Point(20, y)
        lblFriction.Size = Size(150, 20)
        self.Controls.Add(lblFriction)

        self.comboFriction = ComboBox()
        self.comboFriction.Location = Point(180, y)
        self.comboFriction.Size = Size(120, 20)
        self.comboFriction.DropDownStyle = ComboBoxStyle.DropDownList
        for mu in ["0.08","0.10","0.12","0.14","0.16","0.20","0.24"]:
            self.comboFriction.Items.Add(mu)
        self.comboFriction.SelectedItem = str(self.defaultFriction)
        self.Controls.Add(self.comboFriction)

        y += 30
        lblUtilization = Label()
        lblUtilization.Text = "Auslastung Rp*:"
        lblUtilization.Location = Point(20, y)
        lblUtilization.Size = Size(150, 20)
        self.Controls.Add(lblUtilization)

        self.comboUtilization = ComboBox()
        self.comboUtilization.Location = Point(180, y)
        self.comboUtilization.Size = Size(120, 20)
        self.comboUtilization.DropDownStyle = ComboBoxStyle.DropDownList
        for ut in ["1.0","0.95","0.9","0.85","0.8","0.75","0.7","0.65","0.6","0.55","0.5","0.45","0.4","0.3","0.2","0.1"]:
            self.comboUtilization.Items.Add(ut)
        self.comboUtilization.SelectedItem = str(self.defaultUtilization)
        self.Controls.Add(self.comboUtilization)

        y += 30
        lblUtilizationRemark = Label()
        lblUtilizationRemark.Text = "*Bei Ausnutzungsfaktor 0,9 wird der Vorspannwert direkt aus Tab. A1 (VDI) übernommen."\
                                    "Bei anderen Faktoren wird die vereinfachte Formel R7/2 verwendet." 
        lblUtilizationRemark.Location = Point(20, y)
        lblUtilizationRemark.Size = Size(300, 60)
        self.Controls.Add(lblUtilizationRemark)

        y += 10
        self.buttonDefaultBolt = Button()
        self.buttonDefaultBolt.Text = "Defaultschraube anzeigen"
        self.buttonDefaultBolt.Location = Point(400, 220)
        self.buttonDefaultBolt.Size = Size(200, 30)
        self.buttonDefaultBolt.Click += self.onShowDefaultBolt
        self.Controls.Add(self.buttonDefaultBolt)

        self.infoPanel = GroupBox()
        self.infoPanel.Text = "Default Schraubendaten (VDI 2230)"
        self.infoPanel.Location = Point(350, 20)  # links / rechts anpassbar
        self.infoPanel.Size = Size(300, 180)
        self.Controls.Add(self.infoPanel)

        self.lblInfoSize = Label()
        self.lblInfoSize.Location = Point(10, 20)
        self.lblInfoSize.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoSize)

        self.lblInfoStrength = Label()
        self.lblInfoStrength.Location = Point(10, 40)
        self.lblInfoStrength.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoStrength)

        self.lblInfoPitch = Label()
        self.lblInfoPitch.Location = Point(10, 60)
        self.lblInfoPitch.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoPitch)

        self.lblInfoAreaS = Label()
        self.lblInfoAreaS.Location = Point(10, 80)
        self.lblInfoAreaS.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoAreaS)

        self.lblInfoArea0 = Label()
        self.lblInfoArea0.Location = Point(10, 100)
        self.lblInfoArea0.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoArea0)

        self.lblInfoDiamter2 = Label()
        self.lblInfoDiamter2.Location = Point(10, 120)
        self.lblInfoDiamter2.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoDiamter2)

        self.lblInfoUtil = Label()
        self.lblInfoUtil.Location = Point(10, 140)
        self.lblInfoUtil.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoUtil)

        self.lblInfoPreload = Label()
        self.lblInfoPreload.Location = Point(10, 160)
        self.lblInfoPreload.Size = Size(280, 20)
        self.infoPanel.Controls.Add(self.lblInfoPreload)
        # ---------------------------------------------------

        y += 100
        self.buttonCalc = Button()
        self.buttonCalc.Text = "Defaultwerte für Berechnung übernehmen (gelbe Zellen)"
        self.buttonCalc.Location = Point(20, y)
        self.buttonCalc.Size = Size(400, 30)
        self.buttonCalc.Click += self.onCalculate
        self.Controls.Add(self.buttonCalc)

        y += 50
        # --- DataGridView für Gruppen ---
        self.dgv = DataGridView()
        self.dgv.Location = Point(20, y)
        self.dgv.Size = Size(650, 300)
        self.dgv.ColumnCount = 6

        self.dgv.Columns[0].Name = "Größe"
        self.dgv.Columns[0].ReadOnly = True
        self.dgv.Columns[0].Width = 60 

        self.dgv.Columns[1].Name = "Festigkeit"
        self.dgv.Columns[1].ReadOnly = True
        self.dgv.Columns[1].Width = 80

        self.dgv.Columns[2].Name = "Anzahl"
        self.dgv.Columns[2].ReadOnly = True
        self.dgv.Columns[2].Width = 50

        self.dgv.Columns[3].Name = "Auslastung"
        self.dgv.Columns[3].ReadOnly = True
        self.dgv.Columns[3].Width = 70

        self.dgv.Columns[4].Name = "Reibung"
        self.dgv.Columns[4].ReadOnly = True
        self.dgv.Columns[4].Width = 60 

        self.dgv.Columns[5].Name = "Vorspannkraft (editierbar) [N]"
        self.dgv.Columns[5].ReadOnly = False  # editierbar
        self.dgv.Columns[5].ValueType = float
        self.dgv.Columns[5].Width = 270

        self.dgv.AllowUserToAddRows = False

        for i in range(self.dgv.ColumnCount):
            self.dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable

        self.Controls.Add(self.dgv)

        y += 370
        self.okButton = Button()
        self.okButton.Text = "OK"
        self.okButton.Location = Point(180, y)
        self.okButton.Size = Size(100, 30)
        self.okButton.Click += self.onOK
        self.Controls.Add(self.okButton)

        self.cancelButton = Button()
        self.cancelButton.Text = "Abbrechen"
        self.cancelButton.Location = Point(300, y)
        self.cancelButton.Size = Size(100, 30)
        self.cancelButton.Click += self.onCancel
        self.Controls.Add(self.cancelButton)

        self.populate_groups()


    def onCellFormatting(self, sender, e):
        # Spalten 0 = Größe, 1 = Festigkeit, ...
        if e.ColumnIndex in (0, 1, 2, 3, 4, 5):
            value = e.Value
            if value == "Default":
                e.CellStyle.BackColor = Color.Red
                e.CellStyle.ForeColor = Color.Black
            elif e.ColumnIndex in (0, 1, 2):   
                e.CellStyle.BackColor = Color.LightGreen
                e.CellStyle.ForeColor = Color.Black
            elif e.ColumnIndex != 5:
                e.CellStyle.BackColor = Color.Yellow
                e.CellStyle.ForeColor = Color.Black


    def onCalculate(self, sender, args):
        try:
            # 1. Default-Werte holen
            size = self.comboSize.SelectedItem
            strength = self.comboStrength.SelectedItem
            friction = float(self.comboFriction.SelectedItem)
            utilization = float(self.comboUtilization.SelectedItem)

            # 2. Tabelle aktualisieren
            for i, row in enumerate(self.groupTable):
                grp_size = row[0]
                grp_strength = row[1]
                b = metric_bolt()
                # Wenn Gruppe "Default" hat, nutze die neu gesetzten Default-Werte
                final_size = grp_size if grp_size != "Default" else size
                final_strength = grp_strength if grp_strength != "Default" else strength

                b.setSize(final_size)
                b.setStrength(final_strength)
                b.setUtilizationFactor(utilization)
                b.friction = friction
                preload = b.calc_preload()
                self.groupTable[i][5] = preload

                # DataGridView aktualisieren
                self.dgv.Rows[i].Cells[3].Value = utilization
                self.dgv.Rows[i].Cells[4].Value = friction
                self.dgv.Rows[i].Cells[5].Value = preload

                print(preload)

        except Exception as e:
            MessageBox.Show("Fehler bei Eingabe oder Berechnung.\n" + str(e), "Fehler")


    def onShowDefaultBolt(self, sender, args):
        try:
            size = self.comboSize.SelectedItem
            strength = self.comboStrength.SelectedItem
            friction = float(self.comboFriction.SelectedItem)
            utilization = float(self.comboUtilization.SelectedItem)

            bolt = metric_bolt(size, strength, friction, utilizationFactor=utilization)

            pitch = bolt.Pitch
            area = bolt.stressAreas
            preload = bolt.calc_preload()

            # Labels aktualisieren
            self.lblInfoSize.Text = "Nenndurchmesser: " + str(size)
            self.lblInfoStrength.Text = "Festigkeitsklasse: " + str(strength)
            self.lblInfoPitch.Text = "Steigung: " + str(pitch) + " mm"
            self.lblInfoAreaS.Text = "Spannungsquerschnitt: " + str(round(area,2)) + " mm² (d = " + str(round(math.sqrt(4*area/math.pi),2)) + " mm)"
            self.lblInfoArea0.Text = "Kernquerschnitt: " + str(round(bolt.A0_table[bolt.size], 2)) + " mm² (d = " + str(round(math.sqrt(4*bolt.A0_table[bolt.size]/math.pi),2)) + " mm)"
            self.lblInfoDiamter2.Text = "Flankendurchmesser: " + str(bolt.d2_table[bolt.size])+ " mm"
            self.lblInfoUtil.Text = "Auslastung: " + str(utilization*100) +  "%"
            self.lblInfoPreload.Text = "Vorspannkraft: " + str(int(round(preload))) + " N"

        except Exception as e:
            MessageBox.Show("Fehler: " + str(e))


    def populate_groups(self):
        group_dict = {}
        for name in self.pretensionNames:
            bolt = metric_bolt()
            size, strength = get_bolt_data(name, bolt.bolt_sizes, bolt.yield_strengths)
            key = (size,strength)
            group_dict[key] = group_dict.get(key,0) + 1

        self.groupTable = []
        for (size,strength), count in group_dict.items():
            bolt = metric_bolt()
            bolt.setSize(size if size!="Default" else self.defaultSize)
            bolt.setStrength(strength if strength!="Default" else self.defaultStrength)
            bolt.friction = float(self.comboFriction.SelectedItem)
            preload = bolt.calc_preload()
            self.groupTable.append([size, strength, count, bolt.utilizationFactor, bolt.friction, preload])

        self.dgv.Rows.Clear()
        for row in self.groupTable:
            self.dgv.Rows.Add(row[0], row[1], row[2], row[3], row[4], row[5])

        # Farbe
        self.dgv.CellFormatting += self.onCellFormatting


    def onOK(self, sender, args):
        for i,row in enumerate(self.dgv.Rows):
            try:
                val = float(row.Cells[5].Value)
                self.groupTable[i][5] = val
            except:
                pass
        self.DialogResult = DialogResult.OK
        self.Close()


    def onCancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()



# --- MAIN SCRIPT ---
totalPretensions = 0
pretensionNames = []

for analysis in ExtAPI.DataModel.AnalysisList:
    pretensions = analysis.GetChildren(DataModelObjectCategory.BoltPretension, True)
    totalPretensions += len(pretensions)
    for p in pretensions:
        pretensionNames.append(p.Name)


form = PretensionForm(pretensionNames)
result = form.ShowDialog()

if result == DialogResult.OK:

    for analysis in ExtAPI.DataModel.AnalysisList:
        pretensions = analysis.GetChildren(DataModelObjectCategory.BoltPretension, True)
        for p in pretensions:

            pretension = 0
             
            size, strength = get_bolt_data(p.Name, metric_bolt.bolt_sizes, metric_bolt.yield_strengths)
            
            for row in form.groupTable:
                if size == row[0] and strength == row[1]:
                    pretension = row[5]
                    break

            nSteps = len(p.Preload.Output.DiscreteValues)
            if nSteps < 1:
                continue

            newValues = [Quantity(str(pretension) + " [N]")] + [Quantity("0 [N]")] * (nSteps - 1)
            p.Preload.Output.DiscreteValues = newValues

            for i in range(1, nSteps):
                p.SetDefineBy(i + 1, BoltLoadDefineBy.Lock)
else:
    print("Abbruch durch Benutzer")
