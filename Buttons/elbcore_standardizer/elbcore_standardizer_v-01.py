# Script

# ----------------
# Importiere Module
# ----------------
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import *
from System.Drawing import *

class ElbcoreStandardizerForm(Form):
    def __init__(self, initialNumberOfSteps=1, initialLargeDeflection=False):
        Form.__init__(self)
        self.Text = "Elbcore Standardizer"
        self.Size = Size(420, 780)
        self.StartPosition = FormStartPosition.CenterScreen

        # === Defaultwerte ===
        self.userQuadraticMesh = True
        self.userDeactivateAutomaticConnections = True
        self.userDirectSolver = True
        self.userCreateWeakSprings = True
        self.userNumberOfLoadSteps = initialNumberOfSteps
        self.userWriteOnlyLastTimeStep = True
        self.userWriteNodalForces = True
        self.userNonLinearGeometry = initialLargeDeflection
        self.userNumberOfSubStepsInitial = 10
        self.userNumberOfSubStepsMin = 5
        self.userNumberOfSubStepsMax = 100
        self.userRunOnlyPostProcessingMode = False
        self.userCreateVonMisesResults = False
        self.userCreatePrincStressResults = False
        self.userCreateDirectionalStressResults = False
        self.userCreateTotalDeformationResults = False
        self.userCreateDirectionalDeformationResults = False

        y = 10

        # === Mesh Properties ===
        self.labelMode = Label()
        self.labelMode.Text = "Mode:"
        self.labelMode.Location = Point(20, y)
        self.labelMode.Size = Size(200, 20)
        self.Controls.Add(self.labelMode)
        y += 25

        self.chkPostProcessingMode = CheckBox()
        self.chkPostProcessingMode.Text = "Run only in Post Processing Mode"
        self.chkPostProcessingMode.Location = Point(40, y)
        self.chkPostProcessingMode.Size = Size(300, 20)
        self.chkPostProcessingMode.Checked = self.userRunOnlyPostProcessingMode
        self.chkPostProcessingMode.CheckedChanged += self.togglePostProcessingMode
        self.Controls.Add(self.chkPostProcessingMode)
        y += 40


        # === Mesh Properties ===
        self.labelMesh = Label()
        self.labelMesh.Text = "Mesh Properties:"
        self.labelMesh.Location = Point(20, y)
        self.labelMesh.Size = Size(200, 20)
        self.Controls.Add(self.labelMesh)
        y += 25

        self.chkQuadMesh = CheckBox()
        self.chkQuadMesh.Text = "Use Quadratic Mesh"
        self.chkQuadMesh.Location = Point(40, y)
        self.chkQuadMesh.Size = Size(300, 20)
        self.chkQuadMesh.Checked = self.userQuadraticMesh
        self.chkQuadMesh.CheckedChanged += self.toggleQuadMesh
        self.Controls.Add(self.chkQuadMesh)
        y += 40

        # === Connections ===
        self.labelConn = Label()
        self.labelConn.Text = "Connections:"
        self.labelConn.Location = Point(20, y)
        self.labelConn.Size = Size(200, 20)
        self.Controls.Add(self.labelConn)
        y += 25

        self.chkAutoConn = CheckBox()
        self.chkAutoConn.Text = "Deactivate Automatic Connections"
        self.chkAutoConn.Location = Point(40, y)
        self.chkAutoConn.Size = Size(300, 20)
        self.chkAutoConn.Checked = self.userDeactivateAutomaticConnections
        self.chkAutoConn.CheckedChanged += self.toggleAutoConnections
        self.Controls.Add(self.chkAutoConn)
        y += 40

        # === Solver ===
        self.labelSolver = Label()
        self.labelSolver.Text = "Solver:"
        self.labelSolver.Location = Point(20, y)
        self.labelSolver.Size = Size(200, 20)
        self.Controls.Add(self.labelSolver)
        y += 25

        self.chkDirectSolver= CheckBox()
        self.chkDirectSolver.Text = "Use Direct Solver"
        self.chkDirectSolver.Location = Point(40, y)
        self.chkDirectSolver.Size = Size(300, 20)
        self.chkDirectSolver.Checked = self.userDirectSolver
        self.chkDirectSolver.CheckedChanged += self.toggleDirectSolver
        self.Controls.Add(self.chkDirectSolver)
        y += 30

        self.chkWeakSprings = CheckBox()
        self.chkWeakSprings.Text = "Create Weak Springs"
        self.chkWeakSprings.Location = Point(40, y)
        self.chkWeakSprings.Size = Size(300, 20)
        self.chkWeakSprings.Checked = self.userCreateWeakSprings
        self.chkWeakSprings.CheckedChanged += self.toggleWeakSprings
        self.Controls.Add(self.chkWeakSprings)
        y += 30

        label_x = 40
        textbox_x = 200
        textbox_width = 60

        self.lblLoadSteps = Label()
        self.lblLoadSteps.Text = "Number of Load Steps:"
        self.lblLoadSteps.Location = Point(label_x, y + 3)
        self.lblLoadSteps.Size = Size(150, 20)
        self.Controls.Add(self.lblLoadSteps)

        self.txtLoadSteps = TextBox()
        self.txtLoadSteps.Text = str(self.userNumberOfLoadSteps)
        self.txtLoadSteps.Location = Point(textbox_x, y)
        self.txtLoadSteps.Size = Size(textbox_width, 20)
        self.Controls.Add(self.txtLoadSteps)
        y += 30

        self.chkWriteLast = CheckBox()
        self.chkWriteLast.Text = "Store Only Last Result of Each Load Step"
        self.chkWriteLast.Location = Point(40, y)
        self.chkWriteLast.Size = Size(300, 20)
        self.chkWriteLast.Checked = self.userWriteOnlyLastTimeStep
        self.chkWriteLast.CheckedChanged += self.toggleWriteLast
        self.Controls.Add(self.chkWriteLast)
        y += 30

        self.chkNodalForces = CheckBox()
        self.chkNodalForces.Text = "Write Nodal Forces"
        self.chkNodalForces.Location = Point(40, y)
        self.chkNodalForces.Size = Size(300, 20)
        self.chkNodalForces.Checked = self.userWriteNodalForces
        self.chkNodalForces.CheckedChanged += self.toggleNodalForces
        self.Controls.Add(self.chkNodalForces)
        y += 40

        # === Nonlinear Controls ===
        self.labelNL = Label()
        self.labelNL.Text = "Nonlinear Controls:"
        self.labelNL.Location = Point(20, y)
        self.labelNL.Size = Size(200, 20)
        self.Controls.Add(self.labelNL)
        y += 25

        self.chkNonlinear = CheckBox()
        self.chkNonlinear.Text = "Activate Nonlinear Geometry"
        self.chkNonlinear.Location = Point(40, y)
        self.chkNonlinear.Size = Size(300, 20)
        self.chkNonlinear.Checked = self.userNonLinearGeometry
        self.chkNonlinear.CheckedChanged += self.toggleNonlinear
        self.Controls.Add(self.chkNonlinear)
        y += 30

        label_x = 40
        textbox_x = 200
        textbox_width = 60

        self.lblSubInit = Label()
        self.lblSubInit.Text = "Substeps Initial:"
        self.lblSubInit.Location = Point(label_x, y + 3)
        self.lblSubInit.Size = Size(150, 20)
        self.Controls.Add(self.lblSubInit)

        self.txtSubInit = TextBox()
        self.txtSubInit.Text = str(self.userNumberOfSubStepsInitial)
        self.txtSubInit.Location = Point(textbox_x, y)
        self.txtSubInit.Size = Size(textbox_width, 20)
        self.Controls.Add(self.txtSubInit)
        y += 25

        self.lblSubMin = Label()
        self.lblSubMin.Text = "Substeps Minimum:"
        self.lblSubMin.Location = Point(label_x, y + 3)
        self.lblSubMin.Size = Size(150, 20)
        self.Controls.Add(self.lblSubMin)

        self.txtSubMin = TextBox()
        self.txtSubMin.Text = str(self.userNumberOfSubStepsMin)
        self.txtSubMin.Location = Point(textbox_x, y)
        self.txtSubMin.Size = Size(textbox_width, 20)
        self.Controls.Add(self.txtSubMin)
        y += 25

        self.lblSubMax = Label()
        self.lblSubMax.Text = "Substeps Maximum:"
        self.lblSubMax.Location = Point(label_x, y + 3)
        self.lblSubMax.Size = Size(150, 20)
        self.Controls.Add(self.lblSubMax)

        self.txtSubMax = TextBox()
        self.txtSubMax.Text = str(self.userNumberOfSubStepsMax)
        self.txtSubMax.Location = Point(textbox_x, y)
        self.txtSubMax.Size = Size(textbox_width, 20)
        self.Controls.Add(self.txtSubMax)
        y += 40

        # === Results ===
        self.labelResults = Label()
        self.labelResults.Text = "Result Types:"
        self.labelResults.Location = Point(20, y)
        self.labelResults.Size = Size(200, 20)
        self.Controls.Add(self.labelResults)
        y += 25


        self.chkVonMises = CheckBox()
        self.chkVonMises.Text = "Create von Mises Stresses"
        self.chkVonMises.Location = Point(40, y)
        self.chkVonMises.Size = Size(300, 20)
        self.chkVonMises.Checked = self.userCreateVonMisesResults
        self.chkVonMises.CheckedChanged += self.toggleVonMises
        self.Controls.Add(self.chkVonMises)
        y += 25

        self.chkPrinc = CheckBox()
        self.chkPrinc.Text = "Create Principal Stresses"
        self.chkPrinc.Location = Point(40, y)
        self.chkPrinc.Size = Size(300, 20)
        self.chkPrinc.Checked = self.userCreatePrincStressResults
        self.chkPrinc.CheckedChanged += self.togglePrinc
        self.Controls.Add(self.chkPrinc)
        y += 25

        self.chkDirStress = CheckBox()
        self.chkDirStress.Text = "Create Directional Stresses"
        self.chkDirStress.Location = Point(40, y)
        self.chkDirStress.Size = Size(300, 20)
        self.chkDirStress.Checked = self.userCreateDirectionalStressResults
        self.chkDirStress.CheckedChanged += self.toggleDirStress
        self.Controls.Add(self.chkDirStress)
        y += 25

        self.chkTotalDef = CheckBox()
        self.chkTotalDef.Text = "Create Total Deformations"
        self.chkTotalDef.Location = Point(40, y)
        self.chkTotalDef.Size = Size(300, 20)
        self.chkTotalDef.Checked = self.userCreateTotalDeformationResults
        self.chkTotalDef.CheckedChanged += self.toggleTotalDef
        self.Controls.Add(self.chkTotalDef)
        y += 25

        self.chkDirDef = CheckBox()
        self.chkDirDef.Text = "Create Directional Deformations"
        self.chkDirDef.Location = Point(40, y)
        self.chkDirDef.Size = Size(300, 20)
        self.chkDirDef.Checked = self.userCreateDirectionalDeformationResults
        self.chkDirDef.CheckedChanged += self.toggleDirDef
        self.Controls.Add(self.chkDirDef)
        y += 40

        # === Buttons ===
        self.btnOK = Button()
        self.btnOK.Text = "OK"
        self.btnOK.Location = Point(100, y)
        self.btnOK.Size = Size(80, 30)
        self.btnOK.DialogResult = DialogResult.OK
        self.btnOK.Click += self.onOK
        self.Controls.Add(self.btnOK)

        self.btnCancel = Button()
        self.btnCancel.Text = "Cancel"
        self.btnCancel.Location = Point(220, y)
        self.btnCancel.Size = Size(80, 30)
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.btnCancel.Click += self.onCancel
        self.Controls.Add(self.btnCancel)

        # Damit ESC und Enter korrekt funktionieren:
        self.AcceptButton = self.btnOK
        self.CancelButton = self.btnCancel


    # === Event Handlers ===
    def toggleQuadMesh(self, sender, args): self.userQuadraticMesh = self.chkQuadMesh.Checked
    def toggleAutoConnections(self, sender, args): self.userDeactivateAutomaticConnections = self.chkAutoConn.Checked
    def toggleWeakSprings(self, sender, args): self.userCreateWeakSprings = self.chkWeakSprings.Checked
    def toggleWriteLast(self, sender, args): self.userWriteOnlyLastTimeStep = self.chkWriteLast.Checked
    def toggleDirectSolver(self, sender, args): self.userDirectSolver = self.chkDirectSolver.Checked
    def toggleNodalForces(self, sender, args): self.userWriteNodalForces = self.chkNodalForces.Checked
    def toggleNonlinear(self, sender, args): self.userNonLinearGeometry = self.chkNonlinear.Checked
    def togglePostProcessingMode(self, sender, args): self.userRunOnlyPostProcessingMode = self.chkPostProcessingMode.Checked
    def toggleVonMises(self, sender, args): self.userCreateVonMisesResults = self.chkVonMises.Checked
    def togglePrinc(self, sender, args): self.userCreatePrincStressResults = self.chkPrinc.Checked
    def toggleDirStress(self, sender, args): self.userCreateDirectionalStressResults = self.chkDirStress.Checked
    def toggleTotalDef(self, sender, args): self.userCreateTotalDeformationResults = self.chkTotalDef.Checked
    def toggleDirDef(self, sender, args): self.userCreateDirectionalDeformationResults = self.chkDirDef.Checked

    def onOK(self, sender, args):
        try:
            self.userNumberOfLoadSteps = int(self.txtLoadSteps.Text)
            self.userNumberOfSubStepsInitial = int(self.txtSubInit.Text)
            self.userNumberOfSubStepsMin = int(self.txtSubMin.Text)
            self.userNumberOfSubStepsMax = int(self.txtSubMax.Text)
        except:
            MessageBox.Show("Bitte nur ganze Zahlen in Zahlenfeldern eingeben.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        self.DialogResult = DialogResult.OK
        self.Close()

    def onCancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


class LogWindow(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Script-Ausgabe"
        self.Width = 400
        self.Height = 850

        self.textbox = TextBox()
        self.textbox.Multiline = True
        self.textbox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        self.textbox.ReadOnly = True
        self.textbox.Dock = DockStyle.Fill

        self.Controls.Add(self.textbox)

        # Wichtig: Event registrieren
        self.FormClosing += self.onClosing

    def write(self, text):
        clean = text.rstrip("\r\n")
        try:
            self.textbox.AppendText(clean + "\r\n")
        except:
            pass  # Falls schon geschlossen nicht crashen

    def flush(self):
        pass

    def onClosing(self, sender, args):
        # stdout/stderr wiederherstellen
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

        # Textbox sicher entkoppeln, damit kein AppendText passiert
        self.textbox = None

def printDivision():
    print("----------------------------------------------")

################################################################
# Main
################################################################

#------------------------------
# Initital Settings and Status
#------------------------------
analysis = ExtAPI.DataModel.AnalysisList[0]
settings = analysis.AnalysisSettings
initialNumberOfSteps  = settings.NumberOfSteps
initialLargeDeflection= settings.LargeDeflection

finishedSolutions  = False
for analysis in ExtAPI.DataModel.AnalysisList:
    # Check for finished solutions
    if analysis.Solution.Status == SolutionStatusType.Done:
        finishedSolutions = True
        break

#------------------------------
# Open GUI
#------------------------------
form = ElbcoreStandardizerForm(initialNumberOfSteps, initialLargeDeflection)

if form.ShowDialog() == DialogResult.OK:

    # Log Window
    log = LogWindow()
    log.Show()
    log.TopMost = True

    sys.stdout = log
    sys.stderr = log

    # GUI Outputs
    userQuadraticMesh = form.userQuadraticMesh
    userDeactivateAutomaticConnections = form.userDeactivateAutomaticConnections
    userCreateWeakSprings = form.userCreateWeakSprings
    userNumberOfLoadSteps = form.userNumberOfLoadSteps
    userWriteOnlyLastTimeStep = form.userWriteOnlyLastTimeStep
    userDirectSolver = form.userDirectSolver
    userWriteNodalForces = form.userWriteNodalForces
    userNonLinearGeometry = form.userNonLinearGeometry
    userNumberOfSubStepsInitial = form.userNumberOfSubStepsInitial
    userNumberOfSubStepsMin = form.userNumberOfSubStepsMin
    userNumberOfSubStepsMax = form.userNumberOfSubStepsMax
    userRunOnlyPostProcessingMode = form.userRunOnlyPostProcessingMode
    userCreateVonMisesResults = form.userCreateVonMisesResults
    userCreatePrincStressResults = form.userCreatePrincStressResults
    userCreateDirectionalStressResults = form.userCreateDirectionalStressResults
    userCreateTotalDeformationResults = form.userCreateTotalDeformationResults
    userCreateDirectionalDeformationResults = form.userCreateDirectionalDeformationResults

    if finishedSolutions and not userRunOnlyPostProcessingMode:
        warn = MessageBox.Show(
            "At least one solution is solved.\n"
            "These results will be lost if the script continues.\n\n"
            "Do you want to proceed?",
            "Warning – Existing Results Detected",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        )

        if warn == DialogResult.No:
            print("Operation cancelled by user – existing results will not be overwritten.")
            # Restore standard output
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
            # Close log window
            log.Close()
            # Stop script execution
            raise SystemExit


    # --------------------------
    # RUN SCRIPT
    # -------------------------
    printDivision()
    print("ELBCORE STANDARDIZER v0-1")


    printDivision()
    print("pre processor:")


    if not userRunOnlyPostProcessingMode:
        # Mesh
        meshSettings = Model.Mesh
        if userQuadraticMesh and meshSettings.ElementOrder != ElementOrder.Quadratic:
            meshSettings.ElementOrder = ElementOrder.Quadratic
            print("mesh set to quadratic")
        elif not userQuadraticMesh and meshSettings.ElementOrder == ElementOrder.Quadratic:
            meshSettings.ElementOrder = ElementOrder.Linear
            print("mesh set to linear")


    # Connections
    connectionSettings = Model.Connections
    if userDeactivateAutomaticConnections:
        try:
            connectionSettings.GenerateAutomaticConnectionOnRefresh = AutomaticOrManual.Manual
            print("automatic connection generation has been deactivated")
        except:
            print("automatic connection generation could not be deactivated since no connection object exists")
            


    for analysisIndex, analysis in enumerate(ExtAPI.DataModel.AnalysisList):
        printDivision()
        print("solver settings for analysis: " +analysis.Name)

        #------------------------------
        # Solver Settings
        #------------------------------
        settings = analysis.AnalysisSettings

        if not userRunOnlyPostProcessingMode:
            # Weak springs    
            if userCreateWeakSprings and settings.WeakSprings != WeakSpringsType.On:
                settings.WeakSprings = WeakSpringsType.On
                print("weak spring have been added")
            elif not userCreateWeakSprings and settings.WeakSprings != WeakSpringsType.Off:
                settings.WeakSprings = WeakSpringsType.Off
                print("weak spring have been removed")

            # Direct Solver    
            if userDirectSolver and settings.SolverType != SolverType.Direct:
                settings.SolverType = SolverType.Direct
                print("solver type was switched to: direct")
            if not userDirectSolver and settings.SolverType != SolverType.ProgramControlled:
                settings.SolverType = SolverType.ProgramControlled
                print("solver type was switched to: programm controlled")

            # Store nodal forces   
            if userWriteNodalForces and settings.NodalForces != OutputControlsNodalForcesType.Yes:
                settings.NodalForces = OutputControlsNodalForcesType.Yes
                print("nodal forces will be stored")
            elif not userWriteNodalForces and settings.NodalForces == OutputControlsNodalForcesType.Yes:
                settings.NodalForces = OutputControlsNodalForcesType.No
                print("nodal forces will not be stored")

            # Set number of Load Steps 
            if userNumberOfLoadSteps != settings.NumberOfSteps:
                settings.NumberOfSteps = int(userNumberOfLoadSteps)
                print("number of load steps set to " +str(settings.NumberOfSteps))


            # Non Linear Geometry  
            settings.LargeDeflection = userNonLinearGeometry

            for index in range(1,len(analysis.StepsEndTime)+1):
                printDivision()
                print("set solver setting for load step: " +str(index))

                if userWriteOnlyLastTimeStep:
                    settings.SetStoreResultsAt(index, TimePointsOptions.LastTimePoints)
                    print("\t -> store only last result")
                else:
                    settings.SetStoreResultsAt(index, TimePointsOptions.AllTimePoints)
                    print("\t -> store all sub steps")


                # Non linear Controls
                if settings.LargeDeflection:
                    print("\t -> NLGEOM On with sub step settings: (INITIAL:"+ str(userNumberOfSubStepsInitial) + " ;MIN:" + str(userNumberOfSubStepsMin) + " ;MAX:" + str(userNumberOfSubStepsMax)+str(")"))
                    settings.SetAutomaticTimeStepping(index, AutomaticTimeStepping.On)
                    settings.SetInitialSubsteps(index, userNumberOfSubStepsInitial)
                    settings.SetMinimumSubsteps(index, userNumberOfSubStepsMin)
                    settings.SetMaximumSubsteps(index, userNumberOfSubStepsMax)

                else:
                    print("\t -> NLGEOM Off")
                    settings.SetAutomaticTimeStepping(index, AutomaticTimeStepping.Off)
                    settings.SetNumberOfSubSteps(index, 1)
            

        solution = analysis.Solution
        printDivision()
        print("post processing:")
        #------------------------------
        # Stress Results
        #------------------------------
        listGroupStressResults = []

        if userCreateVonMisesResults:
            print("\t -> create von Mises stress results")

            listVonMisesResultObjects = []
            for currEndTime in analysis.StepsEndTime:
                newSolution = solution.AddEquivalentStress()
                newSolution.Name = "Equivalent Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listVonMisesResultObjects.append(newSolution)
            groupVonMisesResultObject =  Tree.Group(listVonMisesResultObjects)
            groupVonMisesResultObject.Name = "Equivalent Stresses"

            listGroupStressResults.append(groupVonMisesResultObject)
            
        if userCreatePrincStressResults:
            print("\t -> create principle stresses results")
            listStress1ResultObjects = []
            listStress2ResultObjects = []
            listStress3ResultObjects = []
            
            for currEndTime in analysis.StepsEndTime:
                newSolution = solution.AddMaximumPrincipalStress()
                newSolution.Name = "Max Principle Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listStress1ResultObjects.append(newSolution)
                
                newSolution = solution.AddMiddlePrincipalStress()
                newSolution.Name = "Mid Principle Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listStress2ResultObjects.append(newSolution)
                
                newSolution = solution.AddMinimumPrincipalStress()
                newSolution.Name = "Min Principle Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listStress3ResultObjects.append(newSolution)
                
            groupStess1ResultObject =  Tree.Group(listStress1ResultObjects)
            groupStess1ResultObject.Name = "Max Principle Stress"
            
            groupStess2ResultObject =  Tree.Group(listStress2ResultObjects)
            groupStess2ResultObject.Name = "Mid Principle Stress"

            groupStess3ResultObject =  Tree.Group(listStress3ResultObjects)
            groupStess3ResultObject.Name = "Min Principle Stress"

            listGroupStressResults.append(groupStess1ResultObject)
            listGroupStressResults.append(groupStess2ResultObject)
            listGroupStressResults.append(groupStess3ResultObject)

        if userCreateDirectionalStressResults:
            print("\t -> create directional stresses results")
            listStressXResultObjects = []
            listStressYResultObjects = []
            listStressZResultObjects = []
            
            for currEndTime in analysis.StepsEndTime:
                newSolution = solution.AddNormalStress()
                newSolution.NormalOrientation = NormalOrientationType.XAxis
                newSolution.Name = "x Normal Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listStressXResultObjects.append(newSolution)
                
                newSolution = solution.AddNormalStress()
                newSolution.NormalOrientation = NormalOrientationType.YAxis
                newSolution.Name = "y Normal Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listStressYResultObjects.append(newSolution)

                newSolution = solution.AddNormalStress()
                newSolution.NormalOrientation = NormalOrientationType.ZAxis
                newSolution.Name = "z Normal Stress LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listStressZResultObjects.append(newSolution)
                
            groupStessXResultObject =  Tree.Group(listStressXResultObjects)
            groupStessXResultObject.Name = "x Normal Stress"
            
            groupStessYResultObject =  Tree.Group(listStressYResultObjects)
            groupStessYResultObject.Name = "y Normal Stress"

            groupStessZResultObject =  Tree.Group(listStressZResultObjects)
            groupStessZResultObject.Name = "z Normal Stress"

            listGroupStressResults.append(groupStessXResultObject)
            listGroupStressResults.append(groupStessYResultObject)
            listGroupStressResults.append(groupStessZResultObject)


        if len(listGroupStressResults) > 1:
            groupAllStressGroups = Tree.Group(listGroupStressResults)
            groupAllStressGroups.Name = "Stress Results"


        #------------------------------
        # Deformation Results
        #------------------------------
        listGroupDeformationResults = []

        if userCreateTotalDeformationResults:
            print("\t -> create total deformations results")

            listTotalDeformationResultObjects = []
            
            for currEndTime in analysis.StepsEndTime:
                newSolution = solution.AddTotalDeformation()
                newSolution.Name = "Total Deformation LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listTotalDeformationResultObjects.append(newSolution)
            groupTotalDeformationResultObject =  Tree.Group(listTotalDeformationResultObjects)
            groupTotalDeformationResultObject.Name = "Total Deformations"

            listGroupDeformationResults.append(groupTotalDeformationResultObject)
            
        if userCreateDirectionalDeformationResults:
            print("\t -> create directional deformations results")

            listDispXResultObjects = []
            listDispYResultObjects = []
            listDispZResultObjects = []
            
            for currEndTime in analysis.StepsEndTime:
                newSolution = solution.AddDirectionalDeformation()
                newSolution.NormalOrientation = NormalOrientationType.XAxis
                newSolution.Name = "x Deformation LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listDispXResultObjects.append(newSolution)
                
                newSolution = solution.AddDirectionalDeformation()
                newSolution.NormalOrientation = NormalOrientationType.YAxis
                newSolution.Name = "y Deformation LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listDispYResultObjects.append(newSolution)
                
                newSolution = solution.AddDirectionalDeformation()
                newSolution.NormalOrientation = NormalOrientationType.ZAxis
                newSolution.Name = "z Deformation LC "+str(currEndTime)
                newSolution.DisplayTime = Quantity(currEndTime, "sec")
                listDispZResultObjects.append(newSolution)

            groupDispXResultObjects =  Tree.Group(listDispXResultObjects)
            groupDispXResultObjects.Name = "x Deformations"
            
            groupDispYResultObjects =  Tree.Group(listDispYResultObjects)
            groupDispYResultObjects.Name = "y Deformations"
            
            groupDispZResultObjects =  Tree.Group(listDispZResultObjects)
            groupDispZResultObjects.Name = "z Deformations"

            listGroupDeformationResults.append(groupDispXResultObjects)
            listGroupDeformationResults.append(groupDispYResultObjects)
            listGroupDeformationResults.append(groupDispZResultObjects)

        if len(listGroupDeformationResults) > 1:
            groupAllDeformationGroups = Tree.Group(listGroupDeformationResults)
            groupAllDeformationGroups.Name = "Deformation Results"