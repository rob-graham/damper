Option Explicit

Private Const BAR_TO_PA As Double = 100000#
Private Const MM2_TO_M2 As Double = 0.000001
Private Const CM2_TO_M2 As Double = 0.0001
Private Const MIN_CAVITATION_BAR As Double = 0.3
Private Const MAX_ITER As Long = 50
Private Const TOL_RESIDUAL As Double = 1E-9
Private Const FD_STEP As Double = 1E-5

' === Public hydraulic utilities ===

Public Function DP_Orifice(flow_m3s As Double, rho As Double, Cd As Double, area_mm2 As Double) As Double
    Dim area_m2 As Double
    area_m2 = area_mm2 * MM2_TO_M2
    If area_m2 <= 0# Or rho <= 0# Or Cd <= 0# Then
        DP_Orifice = 0#
        Exit Function
    End If

    Dim dp_pa As Double
    dp_pa = 0.5 * rho * (flow_m3s / (Cd * area_m2)) ^ 2
    DP_Orifice = dp_pa / BAR_TO_PA
End Function

Public Function DP_Poiseuille(flow_m3s As Double, mu_PaS As Double, length_m As Double, diameter_m As Double) As Double
    If mu_PaS <= 0# Or length_m <= 0# Or diameter_m <= 0# Then
        DP_Poiseuille = 0#
        Exit Function
    End If

    Dim dp_pa As Double
    dp_pa = (128# * mu_PaS * length_m * flow_m3s) / (WorksheetFunction.Pi() * (diameter_m ^ 4))
    DP_Poiseuille = dp_pa / BAR_TO_PA
End Function

Public Function DP_Blend(flow_m3s As Double, rho As Double, mu_PaS As Double, area_mm2 As Double, Cd As Double, Optional length_m As Double = 0#) As Double
    Dim area_m2 As Double
    area_m2 = area_mm2 * MM2_TO_M2
    If area_m2 <= 0# Or rho <= 0# Or Cd <= 0# Then
        DP_Blend = 0#
        Exit Function
    End If

    Dim diameter_m As Double
    diameter_m = Sqr(4# * area_m2 / WorksheetFunction.Pi())
    If length_m <= 0# Then
        length_m = 3# * diameter_m
    End If

    Dim v_mps As Double
    v_mps = flow_m3s / area_m2

    Dim reynolds As Double
    If mu_PaS > 0# Then
        reynolds = rho * v_mps * diameter_m / mu_PaS
    End If

    Dim laminar_dp As Double
    laminar_dp = DP_Poiseuille(flow_m3s, mu_PaS, length_m, diameter_m)

    Dim orifice_dp As Double
    orifice_dp = DP_Orifice(flow_m3s, rho, Cd, area_mm2)

    Dim weight As Double
    weight = Application.Min(1#, Application.Max(0#, (reynolds - 1500#) / (3000# - 1500#)))

    DP_Blend = (1# - weight) * laminar_dp + weight * orifice_dp
End Function

Public Function ShimLift(deltaP_bar As Double, area_mm2 As Double, stackRate_Nmm As Double, Optional preload_N As Double = 0#) As Double
    Dim area_m2 As Double
    area_m2 = area_mm2 * MM2_TO_M2
    Dim force_N As Double
    force_N = deltaP_bar * BAR_TO_PA * area_m2
    If force_N <= preload_N Or stackRate_Nmm <= 0# Then
        ShimLift = 0#
    Else
        ShimLift = (force_N - preload_N) / stackRate_Nmm
    End If
End Function

' === Solver entry points ===

Public Function SolveDamper(direction As String, v As Double, x As Double, topoName As String) As Variant
    Dim cfg As Object
    Set cfg = BuildConfiguration(direction, topoName)
    If cfg Is Nothing Then Exit Function

    Dim nodes As Object
    Set nodes = LoadNodePressures()

    Dim rho As Double
    rho = GetScalarName("OilDensity")

    Dim viscosity_cst As Double
    viscosity_cst = InterpolateViscosity(GetScalarName("Temperature"))
    Dim mu As Double
    mu = viscosity_cst * 1E-6 * rho

    Dim Ap_m2 As Double
    Ap_m2 = GetScalarName("Ap") * CM2_TO_M2
    Dim Ar_m2 As Double
    Ar_m2 = GetScalarName("Ar") * CM2_TO_M2

    Dim dirSign As Double
    If StrComp(direction, "Rebound", vbTextCompare) = 0 Then
        dirSign = -1#
    Else
        dirSign = 1#
    End If

    Dim fixedNodes As Object
    Set fixedNodes = CreateObject("Scripting.Dictionary")
    fixedNodes.CompareMode = vbTextCompare
    fixedNodes("Shaft") = GetScalarName("ShaftPressure")
    fixedNodes("Reservoir") = GetScalarName("ReservoirPressure")

    Dim unknownNames() As String
    Dim unknownValues() As Double
    Dim countUnknown As Long
    Dim nodeKey As Variant
    For Each nodeKey In nodes.Keys
        If Not fixedNodes.Exists(nodeKey) Then
            countUnknown = countUnknown + 1
            If countUnknown = 1 Then
                ReDim unknownNames(1 To 1)
                ReDim unknownValues(1 To 1)
            Else
                ReDim Preserve unknownNames(1 To countUnknown)
                ReDim Preserve unknownValues(1 To countUnknown)
            End If
            unknownNames(countUnknown) = CStr(nodeKey)
            unknownValues(countUnknown) = CDbl(nodes(nodeKey))
        End If
    Next nodeKey

    If countUnknown = 0 Then
        SolveDamper = nodes
        Exit Function
    End If

    Dim iter As Long
    Dim converged As Boolean
    Dim evaluation As Object
    For iter = 1 To MAX_ITER
        Set evaluation = EvaluateNetwork(unknownNames, unknownValues, nodes, fixedNodes, cfg, rho, mu, dirSign, v, Ap_m2, Ar_m2)
        Dim residual() As Double
        residual = evaluation("Residual")
        Dim maxResidual As Double
        maxResidual = MaxAbs(residual)
        If maxResidual < TOL_RESIDUAL Then
            converged = True
            Exit For
        End If

        Dim jac() As Double
        jac = NumericalJacobian(unknownNames, unknownValues, nodes, fixedNodes, cfg, rho, mu, dirSign, v, Ap_m2, Ar_m2)

        Dim delta() As Double
        delta = SolveLinearSystem(jac, residual)
        Dim i As Long
        For i = 1 To countUnknown
            unknownValues(i) = unknownValues(i) - delta(i)
            If unknownValues(i) < MIN_CAVITATION_BAR Then
                unknownValues(i) = MIN_CAVITATION_BAR
            End If
        Next i
    Next iter

    If Not converged Then
        Set evaluation = EvaluateNetwork(unknownNames, unknownValues, nodes, fixedNodes, cfg, rho, mu, dirSign, v, Ap_m2, Ar_m2)
    End If

    Dim pressures As Object
    Set pressures = evaluation("Pressures")

    UpdateNodePressures pressures

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("NodePressures") = pressures

    Dim flows As Object
    Set flows = evaluation("Flows")
    result("ElementFlows") = flows

    Dim pA As Double
    Dim pB As Double
    Dim pR As Double
    If pressures.Exists("ChamberA") Then pA = pressures("ChamberA")
    If pressures.Exists("ChamberB") Then pB = pressures("ChamberB")
    If pressures.Exists("Reservoir") Then pR = pressures("Reservoir")

    Dim force_N As Double
    force_N = (pA - pB) * BAR_TO_PA * Ap_m2 + (pB - pR) * BAR_TO_PA * Ar_m2
    result("Force") = force_N

    Dim deltaP_bar As Double
    deltaP_bar = pA - pB
    result("DeltaP") = deltaP_bar

    Dim losses_bar As Double
    losses_bar = ComputeLosses(cfg, pressures)
    result("Losses") = losses_bar

    Dim cavMargin As Double
    cavMargin = MinValue(pressures) - MIN_CAVITATION_BAR
    result("CavitationMargin") = cavMargin

    Dim bleedFraction As Double
    bleedFraction = ComputeBleedFraction(cfg, flows)
    result("BleedFraction") = bleedFraction

    Dim stackRate As Double
    stackRate = LookupShimStackRate(GetScalarText("ShimStack"))
    Dim compElement As Object
    Set compElement = GetElementByType(cfg, "Compression")
    Dim shimLift_mm As Double
    If Not compElement Is Nothing Then
        Dim compDeltaP As Double
        compDeltaP = pressures(compElement("FromNode")) - pressures(compElement("ToNode"))
        shimLift_mm = ShimLift(compDeltaP, compElement("Area_mm2"), stackRate)
    End If
    result("ShimLift") = shimLift_mm

    SolveDamper = result
End Function

Public Sub RunSweep()
    Dim settings As Object
    Set settings = ReadSolverSettings()

    Dim vMin As Double: vMin = settings("MinVelocity")
    Dim vMax As Double: vMax = settings("MaxVelocity")
    Dim stepSize As Double: stepSize = settings("Step")
    Dim direction As String: direction = settings("Direction")
    Dim topology As String: topology = settings("Topology")
    Dim travel As Double: travel = GetScalarName("TravelPosition")

    If stepSize <= 0# Then Err.Raise vbObjectError + 1, , "Step size must be positive."

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp

    Dim resultsRange As Range
    Set resultsRange = ThisWorkbook.Names("SolverResults").RefersToRange
    Dim dataRange As Range
    Set dataRange = resultsRange.Offset(1, 0).Resize(resultsRange.Rows.Count - 1, resultsRange.Columns.Count)
    dataRange.ClearContents

    Dim v As Double
    Dim rowIndex As Long
    rowIndex = 0
    Dim targetCount As Long
    If vMax >= vMin Then
        targetCount = CLng(Application.RoundDown((vMax - vMin) / stepSize, 0)) + 1
    Else
        targetCount = CLng(Application.RoundDown((vMin - vMax) / stepSize, 0)) + 1
    End If
    If targetCount < 1 Then targetCount = 1
    If targetCount > dataRange.Rows.Count Then targetCount = dataRange.Rows.Count

    Dim i As Long
    For i = 0 To targetCount - 1
        If vMax >= vMin Then
            v = vMin + i * stepSize
            If v > vMax + 1E-9 Then Exit For
        Else
            v = vMin - i * stepSize
            If v < vMax - 1E-9 Then Exit For
        End If

        Dim result As Variant
        result = SolveDamper(direction, Abs(v), travel, topology)
        If IsObject(result) Then
            WriteResultRow resultsRange, rowIndex + 2, v, result
            rowIndex = rowIndex + 1
        End If
    Next i

    If rowIndex > 0 Then
        UpdateSummaryMetrics resultsRange, rowIndex
    Else
        ClearSummaryMetrics
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "RunSweep error: " & Err.Description, vbExclamation
    End If
End Sub

Public Sub SinglePoint()
    Dim settings As Object
    Set settings = ReadSolverSettings()

    Dim velocity As Double
    velocity = GetScalarName("VelocityTarget")
    Dim result As Variant
    result = SolveDamper(settings("Direction"), Abs(velocity), GetScalarName("TravelPosition"), settings("Topology"))
    If Not IsObject(result) Then Exit Sub

    Dim resultsRange As Range
    Set resultsRange = ThisWorkbook.Names("SolverResults").RefersToRange
    resultsRange.Offset(1, 0).Resize(resultsRange.Rows.Count - 1, resultsRange.Columns.Count).ClearContents
    WriteResultRow resultsRange, 2, velocity, result
    UpdateSummaryMetrics resultsRange, 1
End Sub

Public Sub ExportCSV()
    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename(InitialFileName:="damper_sweep.csv", FileFilter:="CSV Files (*.csv), *.csv")
    If VarType(filePath) = vbBoolean And filePath = False Then Exit Sub

    Dim resultsRange As Range
    Set resultsRange = ThisWorkbook.Names("SolverResults").RefersToRange
    Dim dataRange As Range
    Set dataRange = resultsRange.Resize(resultsRange.Rows.Count, resultsRange.Columns.Count)

    Dim rowsToExport As Long
    rowsToExport = CountPopulatedRows(dataRange)
    If rowsToExport <= 1 Then
        MsgBox "No sweep results available to export.", vbInformation
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stream As Object
    Set stream = fso.CreateTextFile(filePath, True, False)

    Dim r As Long, c As Long
    For r = 1 To rowsToExport
        Dim values() As String
        ReDim values(1 To dataRange.Columns.Count)
        For c = 1 To dataRange.Columns.Count
            values(c) = Format$(dataRange.Cells(r, c).Value, "0.############")
        Next c
        stream.WriteLine Join(values, ",")
    Next r
    stream.Close
    MsgBox "Sweep exported to " & CStr(filePath), vbInformation
End Sub

' === Helper routines ===

Private Function BuildConfiguration(direction As String, topoName As String) As Object
    Dim elements As Object
    Set elements = LoadElementDefinitions()

    Dim topology As Object
    Set topology = CreateObject("Scripting.Dictionary")
    topology.CompareMode = vbTextCompare

    Dim rows As Variant
    rows = NamedRangeToArray("TopologyMap")
    Dim r As Long
    If IsArray(rows) Then
        For r = 2 To UBound(rows, 1)
            If LenB(rows(r, 1)) = 0 Then GoTo MaybeAddVirtual
            Dim elemId As String
            elemId = CStr(rows(r, 1))
            Dim dirText As String
            dirText = CStr(rows(r, 4))
            If ShouldIncludeBranch(dirText, direction) Then
                If elements.Exists(elemId) Then
                    Dim branch As Object
                    Set branch = CloneDictionary(elements(elemId))
                    branch("ElementID") = elemId
                    branch("FromNode") = CStr(rows(r, 2))
                    branch("ToNode") = CStr(rows(r, 3))
                    topology(elemId) = branch
                End If
            End If
MaybeAddVirtual:
        Next r
    End If

    AddTopologyExtensions topology, topoName, direction

    If topology.Count = 0 Then
        MsgBox "No active branches for topology '" & topoName & "' in direction '" & direction & "'.", vbExclamation
        Set topology = Nothing
    End If
    Set BuildConfiguration = topology
End Function

Private Function ShouldIncludeBranch(branchDirection As String, requestDirection As String) As Boolean
    Dim branchUpper As String
    branchUpper = UCase$(branchDirection)
    Select Case branchUpper
        Case "BIDIRECTIONAL", "BOTH", "COMMON", "BLEED"
            ShouldIncludeBranch = True
        Case Else
            ShouldIncludeBranch = (StrComp(branchDirection, requestDirection, vbTextCompare) = 0)
    End Select
End Function

Private Sub AddTopologyExtensions(topology As Object, topoName As String, direction As String)
    If topology.Exists("E3") Then
        Dim bleedBranch As Object
        Set bleedBranch = topology("E3")
        Dim bleedArea As Double
        bleedArea = LookupBleedArea(GetScalarName("ClickSetting"))
        If bleedArea > 0# Then
            bleedBranch("Area_mm2") = bleedArea
        End If
    End If

    If StrComp(topoName, "Shock_RemoteRes", vbTextCompare) = 0 Then
        If Not topology.Exists("RR1") Then
            Dim rr As Object
            Set rr = CreateObject("Scripting.Dictionary")
            rr.CompareMode = vbTextCompare
            rr("ElementID") = "RR1"
            rr("Type") = "Remote"
            rr("Cd") = 0.55
            rr("Area_mm2") = 0.45
            rr("FromNode") = IIf(StrComp(direction, "Compression", vbTextCompare) = 0, "ChamberA", "Reservoir")
            rr("ToNode") = IIf(StrComp(direction, "Compression", vbTextCompare) = 0, "Reservoir", "ChamberA")
            topology("RR1") = rr
        End If
    End If
End Sub

Private Function LoadElementDefinitions() As Object
    Dim data As Variant
    data = NamedRangeToArray("ElementDefinitions")
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If IsArray(data) Then
        Dim r As Long
        For r = 2 To UBound(data, 1)
            If LenB(data(r, 1)) = 0 Then GoTo ContinueLoop
            Dim entry As Object
            Set entry = CreateObject("Scripting.Dictionary")
            entry.CompareMode = vbTextCompare
            entry("ElementID") = CStr(data(r, 1))
            entry("Type") = CStr(data(r, 2))
            entry("Cd") = CDbl(data(r, 3))
            entry("Area_mm2") = CDbl(data(r, 4))
            dict(entry("ElementID")) = entry
ContinueLoop:
        Next r
    End If
    Set LoadElementDefinitions = dict
End Function

Private Function LoadNodePressures() As Object
    Dim data As Variant
    data = NamedRangeToArray("NodePressures")
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If IsArray(data) Then
        Dim r As Long
        For r = 2 To UBound(data, 1)
            If LenB(data(r, 1)) = 0 Then GoTo ContinueLoop
            dict(CStr(data(r, 1))) = CDbl(data(r, 2))
ContinueLoop:
        Next r
    End If
    Set LoadNodePressures = dict
End Function

Private Function EvaluateNetwork(unknownNames() As String, unknownValues() As Double, nodes As Object, fixedNodes As Object, topology As Object, rho As Double, mu As Double, dirSign As Double, velocity As Double, Ap_m2 As Double, Ar_m2 As Double) As Object
    Dim pressures As Object
    Set pressures = CreateObject("Scripting.Dictionary")
    pressures.CompareMode = vbTextCompare

    Dim key As Variant
    For Each key In nodes.Keys
        pressures(key) = nodes(key)
    Next key

    Dim i As Long
    For i = LBound(unknownNames) To UBound(unknownNames)
        pressures(unknownNames(i)) = unknownValues(i)
    Next i

    For Each key In fixedNodes.Keys
        pressures(key) = fixedNodes(key)
    Next key

    Dim netFlow As Object
    Set netFlow = CreateObject("Scripting.Dictionary")
    netFlow.CompareMode = vbTextCompare

    Dim flows As Object
    Set flows = CreateObject("Scripting.Dictionary")
    flows.CompareMode = vbTextCompare

    Dim elemKey As Variant
    For Each elemKey In topology.Keys
        Dim branch As Object
        Set branch = topology(elemKey)
        Dim fromNode As String
        fromNode = branch("FromNode")
        Dim toNode As String
        toNode = branch("ToNode")
        Dim deltaP As Double
        deltaP = pressures(fromNode) - pressures(toNode)
        Dim flowValue As Double
        flowValue = ComputeElementFlow(branch, deltaP, rho, mu)
        flows(elemKey) = flowValue
        AddFlow netFlow, fromNode, -flowValue
        AddFlow netFlow, toNode, flowValue
    Next elemKey

    Dim qA As Double
    Dim qB As Double
    qA = -dirSign * velocity * Ap_m2
    qB = dirSign * velocity * (Ap_m2 - Ar_m2)
    AddFlow netFlow, "ChamberA", qA
    AddFlow netFlow, "ChamberB", qB

    Dim residual() As Double
    ReDim residual(LBound(unknownNames) To UBound(unknownNames))
    For i = LBound(unknownNames) To UBound(unknownNames)
        residual(i) = netFlow(unknownNames(i))
    Next i

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Residual") = residual
    result("Pressures") = pressures
    result("Flows") = flows
    result("NetFlow") = netFlow
    Set EvaluateNetwork = result
End Function

Private Function NumericalJacobian(unknownNames() As String, unknownValues() As Double, nodes As Object, fixedNodes As Object, topology As Object, rho As Double, mu As Double, dirSign As Double, velocity As Double, Ap_m2 As Double, Ar_m2 As Double) As Double()
    Dim baseEval As Object
    Set baseEval = EvaluateNetwork(unknownNames, unknownValues, nodes, fixedNodes, topology, rho, mu, dirSign, velocity, Ap_m2, Ar_m2)
    Dim baseResidual() As Double
    baseResidual = baseEval("Residual")

    Dim n As Long
    n = UBound(unknownNames) - LBound(unknownNames) + 1
    Dim jac() As Double
    ReDim jac(1 To n, 1 To n)

    Dim i As Long, j As Long
    For j = 1 To n
        Dim perturbed() As Double
        perturbed = unknownValues
        perturbed(j) = perturbed(j) + FD_STEP
        Dim evalPlus As Object
        Set evalPlus = EvaluateNetwork(unknownNames, perturbed, nodes, fixedNodes, topology, rho, mu, dirSign, velocity, Ap_m2, Ar_m2)
        Dim resPlus() As Double
        resPlus = evalPlus("Residual")

        perturbed(j) = perturbed(j) - 2# * FD_STEP
        Dim evalMinus As Object
        Set evalMinus = EvaluateNetwork(unknownNames, perturbed, nodes, fixedNodes, topology, rho, mu, dirSign, velocity, Ap_m2, Ar_m2)
        Dim resMinus() As Double
        resMinus = evalMinus("Residual")

        For i = 1 To n
            jac(i, j) = (resPlus(i) - resMinus(i)) / (2# * FD_STEP)
        Next i
    Next j
    NumericalJacobian = jac
End Function

Private Function SolveLinearSystem(matrix() As Double, rhs() As Double) As Double()
    Dim n As Long
    n = UBound(rhs)
    Dim a() As Double
    Dim b() As Double
    ReDim a(1 To n, 1 To n)
    ReDim b(1 To n)
    Dim i As Long, j As Long
    For i = 1 To n
        b(i) = rhs(i)
        For j = 1 To n
            a(i, j) = matrix(i, j)
        Next j
    Next i

    For i = 1 To n
        Dim pivot As Double
        pivot = a(i, i)
        If Abs(pivot) < 1E-12 Then pivot = 1E-12
        Dim invPivot As Double
        invPivot = 1# / pivot
        For j = i To n
            a(i, j) = a(i, j) * invPivot
        Next j
        b(i) = b(i) * invPivot

        Dim k As Long
        For k = 1 To n
            If k <> i Then
                Dim factor As Double
                factor = a(k, i)
                If factor <> 0# Then
                    For j = i To n
                        a(k, j) = a(k, j) - factor * a(i, j)
                    Next j
                    b(k) = b(k) - factor * b(i)
                End If
            End If
        Next k
    Next i
    SolveLinearSystem = b
End Function

Private Function ComputeElementFlow(branch As Object, deltaP_bar As Double, rho As Double, mu As Double) As Double
    Dim area_mm2 As Double
    area_mm2 = branch("Area_mm2")
    Dim Cd As Double
    Cd = branch("Cd")
    Dim area_m2 As Double
    area_m2 = area_mm2 * MM2_TO_M2

    Select Case UCase$(branch("Type"))
        Case "BLEED", "REMOTE"
            ComputeElementFlow = SolveFlowFromDP(deltaP_bar, rho, mu, area_mm2, Cd)
        Case Else
            ComputeElementFlow = FlowOrifice(deltaP_bar, Cd, area_m2, rho)
    End Select
End Function

Private Function FlowOrifice(deltaP_bar As Double, Cd As Double, area_m2 As Double, rho As Double) As Double
    If area_m2 <= 0# Or Cd <= 0# Or rho <= 0# Then Exit Function
    Dim dp_pa As Double
    dp_pa = deltaP_bar * BAR_TO_PA
    If Abs(dp_pa) < 1E-9 Then Exit Function
    Dim q As Double
    q = Cd * area_m2 * Sqr(2# * Abs(dp_pa) / rho)
    If dp_pa < 0# Then q = -q
    FlowOrifice = q
End Function

Private Function SolveFlowFromDP(deltaP_bar As Double, rho As Double, mu As Double, area_mm2 As Double, Cd As Double) As Double
    Dim target As Double
    target = deltaP_bar
    Dim q As Double
    Dim guessSign As Double
    guessSign = IIf(deltaP_bar >= 0#, 1#, -1#)
    q = guessSign * 1E-6

    Dim iter As Long
    For iter = 1 To 25
        Dim f As Double
        f = DP_Blend(q, rho, mu, area_mm2, Cd) - target
        If Abs(f) < 1E-9 Then Exit For
        Dim df As Double
        df = (DP_Blend(q + 1E-6, rho, mu, area_mm2, Cd) - DP_Blend(q - 1E-6, rho, mu, area_mm2, Cd)) / (2E-6)
        If Abs(df) < 1E-12 Then Exit For
        q = q - f / df
    Next iter
    SolveFlowFromDP = q
End Function

Private Function ComputeLosses(topology As Object, pressures As Object) As Double
    Dim total As Double
    Dim elemKey As Variant
    For Each elemKey In topology.Keys
        Dim branch As Object
        Set branch = topology(elemKey)
        Dim dp As Double
        dp = pressures(branch("FromNode")) - pressures(branch("ToNode"))
        total = total + Abs(dp)
    Next elemKey
    ComputeLosses = Application.Max(0#, total)
End Function

Private Function ComputeBleedFraction(topology As Object, flows As Object) As Double
    Dim total As Double
    Dim bleed As Double
    Dim key As Variant
    For Each key In flows.Keys
        total = total + Abs(flows(key))
        If topology.Exists(key) Then
            If UCase$(topology(key)("Type")) = "BLEED" Then
                bleed = bleed + Abs(flows(key))
            End If
        End If
    Next key
    If total > 0# Then
        ComputeBleedFraction = bleed / total
    Else
        ComputeBleedFraction = 0#
    End If
End Function

Private Function GetElementByType(topology As Object, elemType As String) As Object
    Dim key As Variant
    For Each key In topology.Keys
        If StrComp(topology(key)("Type"), elemType, vbTextCompare) = 0 Then
            Set GetElementByType = topology(key)
            Exit Function
        End If
    Next key
End Function

Private Sub UpdateNodePressures(pressures As Object)
    Dim rng As Range
    Set rng = ThisWorkbook.Names("NodePressures").RefersToRange
    Dim r As Long
    For r = 2 To rng.Rows.Count
        Dim nodeName As String
        nodeName = CStr(rng.Cells(r, 1).Value)
        If LenB(nodeName) = 0 Then Exit For
        If pressures.Exists(nodeName) Then
            rng.Cells(r, 2).Value = pressures(nodeName)
        End If
    Next r
End Sub

Private Function GetScalarName(name As String) As Double
    GetScalarName = CDbl(ThisWorkbook.Names(name).RefersToRange.Value)
End Function

Private Function GetScalarText(name As String) As String
    GetScalarText = CStr(ThisWorkbook.Names(name).RefersToRange.Value)
End Function

Private Function InterpolateViscosity(temp_C As Double) As Double
    Dim table As Variant
    table = NamedRangeToArray("ViscosityLookup")
    If Not IsArray(table) Then
        InterpolateViscosity = 100#
        Exit Function
    End If

    Dim lastTemp As Double
    Dim lastVisc As Double
    Dim r As Long
    For r = 2 To UBound(table, 1)
        If LenB(table(r, 1)) = 0 Then Exit For
        Dim t As Double
        t = CDbl(table(r, 1))
        Dim visc As Double
        visc = CDbl(table(r, 2))
        If temp_C <= t Then
            If r = 2 Then
                InterpolateViscosity = visc
            Else
                Dim frac As Double
                frac = (temp_C - lastTemp) / (t - lastTemp)
                InterpolateViscosity = lastVisc + frac * (visc - lastVisc)
            End If
            Exit Function
        End If
        lastTemp = t
        lastVisc = visc
    Next r
    InterpolateViscosity = lastVisc
End Function

Private Function LookupBleedArea(clickSetting As Double) As Double
    Dim table As Variant
    table = NamedRangeToArray("ClickAreaLookup")
    If Not IsArray(table) Then Exit Function
    Dim r As Long
    For r = 2 To UBound(table, 1)
        If LenB(table(r, 1)) = 0 Then Exit For
        If CDbl(table(r, 1)) = clickSetting Then
            LookupBleedArea = CDbl(table(r, 2))
            Exit Function
        End If
    Next r
End Function

Private Function LookupShimStackRate(code As String) As Double
    Dim table As Variant
    table = NamedRangeToArray("ShimStackCatalog")
    If Not IsArray(table) Then Exit Function
    Dim r As Long
    For r = 2 To UBound(table, 1)
        If LenB(table(r, 3)) = 0 Then Exit For
        If StrComp(CStr(table(r, 3)), code, vbTextCompare) = 0 Then
            LookupShimStackRate = CDbl(table(r, 5))
            Exit Function
        End If
    Next r
    LookupShimStackRate = 50#
End Function

Private Function NamedRangeToArray(name As String) As Variant
    NamedRangeToArray = ThisWorkbook.Names(name).RefersToRange.Value
End Function

Private Function CloneDictionary(source As Object) As Object
    Dim clone As Object
    Set clone = CreateObject("Scripting.Dictionary")
    clone.CompareMode = vbTextCompare
    Dim key As Variant
    For Each key In source.Keys
        clone(key) = source(key)
    Next key
    Set CloneDictionary = clone
End Function

Private Function MaxAbs(values() As Double) As Double
    Dim i As Long
    For i = LBound(values) To UBound(values)
        MaxAbs = Application.Max(MaxAbs, Abs(values(i)))
    Next i
End Function

Private Function MinValue(dict As Object) As Double
    Dim key As Variant
    MinValue = 1E+12
    For Each key In dict.Keys
        MinValue = Application.Min(MinValue, CDbl(dict(key)))
    Next key
End Function

Private Function ReadSolverSettings() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    dict("MinVelocity") = 0#
    dict("MaxVelocity") = 0.6
    dict("Step") = 0.1
    dict("Direction") = "Compression"
    dict("Topology") = "Fork_OpenCartridge"

    Dim data As Variant
    data = NamedRangeToArray("SolverSettings")
    If IsArray(data) Then
        Dim r As Long
        For r = 1 To UBound(data, 1)
            Dim key As String
            key = CStr(data(r, 1))
            If LenB(key) = 0 Then Exit For
            dict(key) = data(r, 2)
        Next r
    End If
    Set ReadSolverSettings = dict
End Function

Private Sub WriteResultRow(resultsRange As Range, rowNumber As Long, velocity As Double, result As Variant)
    resultsRange.Cells(rowNumber, 1).Value = velocity
    resultsRange.Cells(rowNumber, 2).Value = result("Force")
    resultsRange.Cells(rowNumber, 3).Value = result("DeltaP")
    resultsRange.Cells(rowNumber, 4).Value = result("Losses")
    resultsRange.Cells(rowNumber, 5).Value = result("CavitationMargin")
    resultsRange.Cells(rowNumber, 6).Value = result("BleedFraction")
End Sub

Private Sub UpdateSummaryMetrics(resultsRange As Range, rowCount As Long)
    Dim forceRange As Range
    Set forceRange = resultsRange.Offset(1, 1).Resize(rowCount, 1)
    Dim dpRange As Range
    Set dpRange = resultsRange.Offset(1, 2).Resize(rowCount, 1)
    Dim lossRange As Range
    Set lossRange = resultsRange.Offset(1, 3).Resize(rowCount, 1)
    Dim cavRange As Range
    Set cavRange = resultsRange.Offset(1, 4).Resize(rowCount, 1)
    Dim bleedRange As Range
    Set bleedRange = resultsRange.Offset(1, 5).Resize(rowCount, 1)

    Dim metricsRange As Range
    Set metricsRange = ThisWorkbook.Worksheets("Solver").Range("H8:I12")
    metricsRange.Cells(1, 2).Value = WorksheetFunction.Max(forceRange)
    metricsRange.Cells(2, 2).Value = WorksheetFunction.Average(dpRange)
    metricsRange.Cells(3, 2).Value = WorksheetFunction.Average(bleedRange)
    metricsRange.Cells(4, 2).Value = WorksheetFunction.Min(cavRange)
    metricsRange.Cells(5, 2).Value = WorksheetFunction.Max(lossRange)
End Sub

Private Sub ClearSummaryMetrics()
    ThisWorkbook.Worksheets("Solver").Range("H8:I12").Columns(2).ClearContents
End Sub

Private Function CountPopulatedRows(rng As Range) As Long
    Dim r As Long
    For r = rng.Rows.Count To 1 Step -1
        If Application.CountA(rng.Rows(r)) > 0 Then
            CountPopulatedRows = r
            Exit Function
        End If
    Next r
    CountPopulatedRows = 0
End Function

Private Sub AddFlow(bucket As Object, node As String, value As Double)
    If Not bucket.Exists(node) Then
        bucket(node) = value
    Else
        bucket(node) = bucket(node) + value
    End If
End Sub
 
EOF
)