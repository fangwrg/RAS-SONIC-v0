Sub Controller()
    'Making background excel window static
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Define Variables
    Dim ProjectPath, ProjectName, projectFilePath, StartTime, EndTime, geomFilePath, uflowFilePath, planFilePath, planFileOutputPath, PlanTitle, WSEcsvOutputPath, planTimeOutputPath, PairedTimes() As String
    Dim cellNoForOutput As Long
    Dim NoOfSteps, monthVal, runLimit, runCycle As Integer
    Dim learn As Boolean
    Dim RC As New RAS631.HECRASController
    
    'Reading Project Inputs
    ProjectPath = cell("Workflow", 2, 1)
    ProjectName = cell("Workflow", 2, 2)
    projectFilePath = ProjectPath & "\" & ProjectName & ".prj"
    cellNoForOutput = cell("Workflow", 2, 9)
    StartTime = cell("Workflow", 2, 7)
    EndTime = cell("Workflow", 2, 8)
    
    'Learn params
    learn = True
    runLimit = 100
    runCycle = 0
    
    'Start Sequence of calls for program run
restartWholeSequence:
    Call RC.Project_Open(projectFilePath)
    Call RC.ShowRas
    
    'Storing RAS and controller file paths
    geomFilePath = RC.CurrentGeomFile
    Call cell("Workflow", 2, 4, geomFilePath)
    uflowFilePath = RC.CurrentUnSteadyFile
    Call cell("Workflow", 2, 5, uflowFilePath)
    PlanTitle = cell("Workflow", 2, 3)
    Call RC.Plan_SetCurrent(PlanTitle)
    planFilePath = RC.CurrentPlanFile
    Call cell("Workflow", 2, 6, planFilePath)
    planFileOutputPath = planFilePath & ".hdf"
    WSEcsvOutputPath = ThisWorkbook.Path & "\" & "\wse.csv"
    planTimeOutputPath = ThisWorkbook.Path & "\" & "planTime.txt"
    ' Reset state by deleting old output file if it exists
    deleteOldFile (WSEcsvOutputPath)
    
    ' Split time step input by month
    NoOfSteps = generateMonthlyTime(cell("Workflow", 3, 7), cell("Workflow", 3, 8), PairedTimes, monthVal)
    ' The function below assigns roughness values to specific cells in excel DB, spatially and temporally varied (each month) based on only ONE factor
    Call readSeasonalParamsAndPopulateSeasonalRoughness
    
    'Main Program Execution, runs plans by seperating them into time steps
    For i = 0 To NoOfSteps
        RC.QuitRas
        Call changePlanTime(planFilePath, PairedTimes(i))
        If monthVal = 13 Then
            monthVal = 1
        End If
        'Write roughness values (temporally varied) to geometry file
        Call replaceRoughnessValueInGeometryFile(geomFilePath, monthVal, 166147, 71)
        monthVal = monthVal + 1
        'Assign the latest restart file to the unsteady flow file reference, or assign the first file if start of simulation
        Call resetRestartFile(ProjectPath, ProjectPath & "\Backup\firstFile.rst", i <= 0)
        
restartCurrentPlan:
        Call RC.Project_Open(projectFilePath)
        Call RC.ShowRas
        Call RC.Plan_SetCurrent(PlanTitle)
        Dim messageReturned() As String
        Call RC.Compute_CurrentPlan(1, messageReturned, True)
        ' Check if plan ran correctly, this step can be limited to some run cycles however it's not required since unless there is a fundamental issue the second or third time running the plan let's it work without issues
        If Not (getOutputMatches(planFileOutputPath, cellNoForOutput, planTimeOutputPath, True, PairedTimes(i))) Then GoTo restartCurrentPlan
        ' Read output WSE at a specific cell
        Call getWSEfromOutputHdf(planFileOutputPath, cellNoForOutput, WSEcsvOutputPath, True)
        ' Save read WSE to excel DB
        Call writeWSEfromObtainedCsv(WSEcsvOutputPath)
        ThisWorkbook.Save
    Next
    Call RC.QuitRas
    If learn Then
        ' Start Learning Algorithm and rerun entire program
        If learningAlgo(runLimit, runCycle) Then GoTo restartWholeSequence
    End If
End Sub

Function learningAlgo(ByRef runLimit As Integer, ByRef runCycle As Integer) As Boolean
    Dim averageWSEObserved, averageWSESimulated, avgDiff, ratioAvgDiff, thresholdValue, previousParameter, newParameter, someFactor As Double
    Dim thresholdReached, runLimitMet, restart As Boolean
    runLimitMet = False
    thresholdReached = False
    restart = False
    Application.Calculate
    
    ' get avgerage difference or any index parameter like NSE, RMSE, etc.
        averageWSEObserved = cell("OutputWSE", 11, 1) 'Stores the average observed WSE value
        averageWSESimulated = cell("OutputWSE", 11, 2) 'Stores the average simulated WSE value
        avgDiff = averageWSESimulated - averageWSEObserved
        ratioAvgDiff = Abs(avgDiff / averageWSEObserved)
        thresholdValue = 0.00001
    
    ' get Geom Value
        'Stores One parameter that controls the mean of roughness values
        'Can be replaced by one roughness value, channel roughness value, or a distribution of spatially varied roughness by user's choice
        previousParameter = cell("SeasonalRoughness", 31, 13)
    
    ' a crude example is given below, how to change this can be configured by user, ML approaches are added here
        someFactor = 10
        newParameter = previousParameter + ratioAvgDiff * previousParameter * someFactor 'if avgDiff > 0, increase roughness
        Call cell("SeasonalRoughness", 31, 13, newParameter)
        Application.Calculate
    
    runCycle = runCycle + 1
    If runCycle >= runLimit Then runLimitMet = True
    If ratioAvgDiff <= thresholdValue Then thresholdReached = True
    If Not (thresholdReached Or runLimitMet) Then restart = True
    learningAlgo = restart
End Function

Function resetRestartFile(ByVal ProjectPath As String, ByVal firstRestartFilePath As String, ByVal reset As Boolean)
        If reset Then
            deleteOldFile (ProjectPath & "\LakeLivingston.p03.rst")
            Call FileCopy(firstRestartFilePath, ProjectPath & "\LakeLivingston.p03.rst")
        Else
            Dim newRSTfile As String
            newRSTfile = ProjectPath & "\" & newestRstFile(ProjectPath)
            Dim newFileName As String
            newFileName = ProjectPath & "\" & "LakeLivingston.p03.rst"
            If newRSTfile <> newFileName Then
                deleteOldFile (newFileName)
                Name newRSTfile As newFileName
            End If
        End If
End Function

Function newestRstFile(ByVal filePath As String) As String
    Dim FileName As String
    Dim MostRecentFile As String
    Dim MostRecentDate As Date
    Dim FileSpec As String
    FileSpec = "*.rst"
    Directory = filePath & "\"
    FileName = Dir(Directory & FileSpec)
    If FileName <> "" Then
        MostRecentFile = FileName
        MostRecentDate = FileDateTime(Directory & FileName)
        Do While FileName <> ""
            If FileDateTime(Directory & FileName) > MostRecentDate Then
                 MostRecentFile = FileName
                 MostRecentDate = FileDateTime(Directory & FileName)
            End If
            FileName = Dir
        Loop
    End If
    newestRstFile = MostRecentFile
End Function

Function deleteOldFile(ByVal filePath As String)
    On Error Resume Next
        Kill (filePath)
    Resume Next
End Function

Function generateMonthlyTime(ByVal StartTime As Double, ByVal EndTime As Double, ByRef PairedTimes, ByRef FirstMonth) As Integer
    Dim AWB As WorksheetFunction
    Set AWB = Application.WorksheetFunction
    Dim startDay, startMonth, startYear, startHr, startMin As Integer
    Dim endDay, endMonth, endYear, endHr, endMin As Integer
    startDay = day(StartTime)
    startMonth = month(StartTime)
    startYear = year(StartTime)
    startHr = hour(StartTime)
    startMin = Minute(StartTime)
    endDay = day(EndTime)
    endMonth = month(EndTime)
    endYear = year(EndTime)
    endHr = hour(EndTime)
    endMin = Minute(EndTime)
    
    Dim MsgBoxText As String
    Dim BeginningToEnd(0 To 1000) As String
    BeginningToEnd(0) = convertDateToHecRasGeomFormat(startDay, startMonth, startYear, startHr, startMin)
    'MsgBoxText = BeginningToEnd(0)
    step = StartTime
    arrIndex = 1
    stepMonth = startMonth
    While (step - 1.1) < EndTime
        If step >= EndTime Then
            BeginningToEnd(arrIndex) = convertDateToHecRasGeomFormat(endDay, endMonth, endYear, endHr, endMin)
            'MsgBoxText = MsgBoxText + ", " + BeginningToEnd(arrIndex)
            arrIndex = arrIndex + 1
            GoTo exitwhile
        ElseIf stepMonth <> month(step) Then
            BeginningToEnd(arrIndex) = convertDateToHecRasGeomFormat(day(step), month(step), year(step), hour(step), Minute(step))
            'MsgBoxText = MsgBoxText + ", " + BeginningToEnd(arrIndex)
            stepMonth = stepMonth + 1
            If stepMonth = 13 Then
                stepMonth = 1
            End If
            arrIndex = arrIndex + 1
        End If
        step = step + 1
    Wend
exitwhile:
    Dim pairsOfMonthlyTime(0 To 1001) As String
    For i = 0 To (arrIndex - 2)
        pairsOfMonthlyTime(i) = BeginningToEnd(i) & "," & BeginningToEnd(i + 1)
        'MsgBoxText = MsgBoxText + Chr(10) + pairsOfMonthlyTime(i)
    Next
    'MsgBox MsgBoxText
    PairedTimes = pairsOfMonthlyTime
    generateMonthlyTime = arrIndex - 2
    FirstMonth = startMonth
End Function

Function convertDateToHecRasGeomFormat(ByVal day As Integer, ByVal month As Integer, ByVal year As Integer, ByVal hour As Integer, ByVal mins As Integer)
    convertDateToHecRasGeomFormat = Format(day, "00") & MonthName(month, True) & Format(year, "0000") & "," & Format(hour, "00") & ":" & Format(mins, "00")
End Function

Function changePlanTime(ByVal filePath As String, planTimePair As String)
    Dim lineToWrite As String
        lineToWrite = "Simulation Date=" & planTimePair
    Call readwriteLineToString(filePath, 4, lineToWrite)
End Function

Function writeWSEfromObtainedCsv(ByVal filePath As String)
    Dim csvWB As Workbook
    Set csvWB = Workbooks.Open(filePath)
    Dim csvWS As Worksheet
    Set csvWS = csvWB.Sheets(1)
    'MsgBox filePath
    csvWS.Range("A1:B1048575").Copy
    ThisWorkbook.Sheets("OutputWSE").Range("A2:B1048576").PasteSpecial xlPasteAll
    csvWB.Close
End Function

Function getWSEfromOutputHdf(ByVal filePath As String, ByVal cellNo As Long, ByVal csvLocation As String, ByVal break As Boolean)
    Dim consoleCommand As String
        consoleCommand = Chr(34) & Application.ThisWorkbook.Path & "\" & "writeWSE.exe" & Chr(34) & " " & Chr(34) & filePath & Chr(34) & " " & cellNo & " " & Chr(34) & csvLocation & Chr(34)
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim windowStyle As Integer: windowStyle = 1
    Call wsh.Run(consoleCommand, windowStyle, break)
    'Call Shell(consoleCommand, vbMinimizedNoFocus)
End Function

Function getOutputMatches(ByVal filePath As String, ByVal cellNo As Long, ByVal txtLocation As String, ByVal break As Boolean, ByVal currentRunTime As String) As Boolean
    Dim consoleCommand As String
        consoleCommand = Chr(34) & Application.ThisWorkbook.Path & "\" & "getPlanTime.exe" & Chr(34) & " " & Chr(34) & filePath & Chr(34) & " " & cellNo & " " & Chr(34) & txtLocation & Chr(34)
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim windowStyle As Integer: windowStyle = 1
    Call wsh.Run(consoleCommand, windowStyle, break)
    getOutputMatches = currentRunTime = readFileToString(txtLocation)
    'Call Shell(consoleCommand, vbMinimizedNoFocus)
End Function

Function readGeometryFile(ByVal filePath As String, ByVal lineStart As Long, ByVal noOfLines As Long)
    For i = 0 To (noOfLines - 1)
        Call cell("SeasonalRoughness", 1, 2 + i, readwriteLineToString(filePath, lineStart + i))
    Next
End Function

Function readSeasonalParamsAndPopulateSeasonalRoughness()
    For y = 8 To 19
        For x = 25 To 30
            Dim params As Double
            params = cell("SeasonalRoughness", x, y)
            Call cell("SeasonalRoughness", x, 5, params)
        Next
        Dim month As Integer
        month = y - 7
        Sheets("SeasonalRoughness").Calculate
        For Z = 2 To 72
            Dim roughness As Double
            roughness = CDbl(cell("SeasonalRoughness", 20, Z))
            Call cell("SeasonalRoughness", 3 + month, Z, roughness)
        Next
    Next
End Function

Function replaceRoughnessValueInGeometryFile(ByVal filePath As String, ByVal month As Long, ByVal lineStart As Long, ByVal noOfLines As Long)
    For i = 0 To (noOfLines - 1)
        Call readwriteLineToString(filePath, lineStart + i, cell("SeasonalRoughness", 3 + month, 75 + i))
    Next
End Function

Function populateOutput(ByVal data As String)
    Sheets("Workflow").Cells(10, 2).Value = data
End Function

Function cell(ByVal worksheetName As String, ByVal columnNo As Integer, ByVal rowNo As Integer, Optional ByVal writeVal As String) As String
    If writeVal <> "" Then
        Sheets(worksheetName).Cells(rowNo, columnNo).Value = writeVal
    End If
    cell = Sheets(worksheetName).Cells(rowNo, columnNo).Value
End Function

Function readFileToString(ByVal filePath As String) As String
    Dim strText As String
    Dim FSO As Object 'File System Object
    Dim TSO As Object 'Text System Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TSO = FSO.OpenTextFile(filePath)
    strText = TSO.ReadAll
    TSO.Close
    Set TSO = Nothing
    Set FSO = Nothing
    readFileToString = strText
End Function

Function readwriteLineToString(ByVal filePath As String, ByVal lineNumber As Long, Optional ByVal writeLine As String) As String
    Dim tempFilePath As String
    tempFilePath = filePath & ".temp"
    Dim writeB As Boolean
    If writeLine = "" Then
        writeB = False
    Else
        writeB = True
    End If
    Dim strText As String
    Dim lineNoVar As Long
    lineNoVar = 1
    Dim FSO As Object
    Dim TSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TSO = FSO.OpenTextFile(filePath)
    If writeB Then
        Dim writeStream As Object
        Set writeStream = FSO.CreateTextFile(tempFilePath)
        Do While Not TSO.AtEndOfStream
            If lineNoVar = lineNumber Then
                writeStream.writeLine (writeLine)
                Dim oldValue As String
                oldValue = TSO.ReadLine
                strLine = "Replaced:" & Chr(10) & oldValue & Chr(10) & "With:" & Chr(10) & writeLine
            Else
                writeStream.writeLine (TSO.ReadLine)
            End If
            lineNoVar = lineNoVar + 1
        Loop
        writeStream.Close
        Set writeStream = Nothing
    Else
        Do While Not TSO.AtEndOfStream
            If lineNoVar = lineNumber Then
                strLine = TSO.ReadLine
                Exit Do
            Else
                TSO.ReadLine
            End If
            lineNoVar = lineNoVar + 1
        Loop
    End If
    TSO.Close
    Set TSO = Nothing
    Set FSO = Nothing
    If writeB Then
        Kill (filePath)
        Name tempFilePath As filePath
    End If
    readwriteLineToString = strLine
End Function



