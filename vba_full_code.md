
``` vb
Private Sub Worksheet_Activate()
    Dim name As Shape, age As Shape, gender As Shape, mail As Shape
    Dim portfolio As Shape, Monte_Carlo As Shape, weightData As Shape
    Dim information As Worksheet, portfoliosheet As Worksheet, optimizationSheet As Worksheet
    Dim ws As Worksheet, lastRow As Long, portfolioData As String, label As String
    Dim cell As Range, rng As Range, i As Long, decimalPlaces As Integer
    Dim chartObj As ChartObject
    Dim mean As Shape, var As Shape, Std As Shape, Sharpe As Shape

    ' Set desired decimal places
    decimalPlaces = 3

    Set ws = ThisWorkbook.Worksheets("風險厭惡結果單")
    
    ' Clear all charts on the worksheet
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    ' Initialize shapes for key metrics
    Set mean = ws.Shapes("平均數")      ' Average
    Set var = ws.Shapes("變異數")       ' Variance
    Set Std = ws.Shapes("標準差")       ' Standard Deviation
    Set Sharpe = ws.Shapes("夏普率")    ' Sharpe Ratio

    ' Reset shapes' text
    mean.TextFrame.Characters.Text = ""
    var.TextFrame.Characters.Text = ""
    Std.TextFrame.Characters.Text = ""
    Sharpe.TextFrame.Characters.Text = ""

    ' Set worksheets and shapes for user data
    Set information = ThisWorkbook.Worksheets("使用者資料")
    Set portfoliosheet = ThisWorkbook.Worksheets("雪天型結果")
    Set optimizationSheet = ThisWorkbook.Worksheets("幕後結果_最佳化")
    Set name = ActiveSheet.Shapes("暱稱")
    Set age = ActiveSheet.Shapes("年齡")
    Set gender = ActiveSheet.Shapes("性別")
    Set mail = ActiveSheet.Shapes("郵件")
    Set portfolio = ActiveSheet.Shapes("portfolio")
    Set Monte_Carlo = ActiveSheet.Shapes("蒙地卡羅")
    Set weightData = ActiveSheet.Shapes("權重")

    ' Format numeric values in range H33:Q33
    Set rng = ws.Range("H33:Q33")
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            cell.Value = Format(cell.Value, "0." & String(decimalPlaces, "0"))
        End If
    Next cell

    ' Retrieve last row from "使用者資料" and update shapes with user info
    lastRow = information.Cells(information.Rows.Count, "A").End(xlUp).Row
    name.TextFrame.Characters.Text = information.Cells(lastRow, 1).Value
    age.TextFrame.Characters.Text = information.Cells(lastRow, 2).Value
    gender.TextFrame.Characters.Text = information.Cells(lastRow, 3).Value
    mail.TextFrame.Characters.Text = information.Cells(lastRow, 4).Value

    ' Retrieve portfolio data from "雪天型結果"
    portfolioData = ""
    For i = 2 To 7
        If i = 7 Then
            portfolioData = portfolioData & portfoliosheet.Cells(i, 1).Value & vbCrLf
        Else
            portfolioData = portfolioData & portfoliosheet.Cells(i, 1).Value & ","
        End If
    Next i
    For i = 2 To 4
        If i = 4 Then
            portfolioData = portfolioData & portfoliosheet.Cells(i, 2).Value & vbCrLf
        Else
            portfolioData = portfolioData & portfoliosheet.Cells(i, 2).Value & ","
        End If
    Next i
    For i = 2 To 2
        portfolioData = portfolioData & portfoliosheet.Cells(i, 3).Value & vbCrLf
    Next i
    portfolio.TextFrame.Characters.Text = portfolioData

    ' Get labels from "幕後結果_最佳化" range AN1:AW1, separated by spaces, and update Monte_Carlo shape
    label = ""
    For Each cell In optimizationSheet.Range("AN1:AW1")
        label = label & cell.Value & "     "
    Next cell
    label = Trim(label)
    Monte_Carlo.TextFrame2.TextRange.Text = label
    With Monte_Carlo.TextFrame2
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' Concatenate values in range H33:Q33 and update weightData shape
    label = ""
    For Each cell In rng
        label = label & cell.Value & "        "
    Next cell
    label = Trim(label)
    weightData.TextFrame.Characters.Text = label
    With weightData.TextFrame
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

```
```vb
Sub ExtractTextFromShape()
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape
    Dim txtRng As Range, usersheet As Worksheet, nextRow As Long, found As Boolean
    Dim shp1Text As String  ' Store text from shp1
    Dim possheet As Worksheet

    Set usersheet = ThisWorkbook.Worksheets("使用者資料")
    Set shp1 = ActiveSheet.Shapes("文字方塊 1")
    Set shp2 = ActiveSheet.Shapes("文字方塊 2")
    Set shp3 = ActiveSheet.Shapes("文字方塊 3")
    Set shp4 = ActiveSheet.Shapes("文字方塊 4")
    Set possheet = ThisWorkbook.Sheets("POS機按鈕")

    ' Get text from shp1 and trim spaces
    shp1Text = Trim(shp1.TextFrame.Characters.Text)

    ' Check if nickname already exists in column A
    found = False
    For Each txtRng In usersheet.Range("A:A").SpecialCells(xlCellTypeConstants)
        If Trim(txtRng.Value) = shp1Text Then
            found = True
            Exit For
        End If
    Next txtRng

    If found Then
        MsgBox "Nickname already exists. Please choose another.", vbExclamation, "Duplicate Nickname"
    Else
        nextRow = usersheet.Cells(usersheet.Rows.Count, "A").End(xlUp).Row + 1
        Set txtRng = usersheet.Range("A" & nextRow)
        txtRng.Value = shp1Text
        txtRng.Offset(0, 1).Value = shp2.TextFrame.Characters.Text
        txtRng.Offset(0, 2).Value = shp3.TextFrame.Characters.Text
        txtRng.Offset(0, 3).Value = shp4.TextFrame.Characters.Text

        ' Clear text boxes
        shp1.TextFrame.Characters.Text = ""
        shp2.TextFrame.Characters.Text = ""
        shp3.TextFrame.Characters.Text = ""
        shp4.TextFrame.Characters.Text = ""
    End If
    possheet.Activate
End Sub
```

```vb
' Declaration of global variables for the quiz module
Dim questions As Worksheet
Dim quizActive As Boolean
Dim N As Long
Dim textblock As Shape
Dim Ablock As Shape
Dim Bblock As Shape
Dim Cblock As Shape
Dim Dblock As Shape
Dim selectedCategory As String

' Start the quiz for a given category (button name)
Sub StartQuiz(categoryName As String)
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Worksheets("POS機")
    
    targetSheet.Activate
    selectedCategory = categoryName
    N = 2
    
    InitializeTextBoxes
    quizActive = True
    ShowQuestion
End Sub

' Display a question from the question bank
Sub ShowQuestion()
    Set questions = ThisWorkbook.Worksheets("題庫")
    Dim maxRow As Long, totalPoints As Integer
        
    maxRow = questions.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' Set totalPoints based on selected category
    Select Case selectedCategory
        Case "退休金": totalPoints = questions.Cells(2, 8).Value
        Case "第一桶金": totalPoints = questions.Cells(7, 8).Value
        Case "財產保值": totalPoints = questions.Cells(12, 8).Value
        Case "教育存款": totalPoints = questions.Cells(17, 8).Value
        Case "買房": totalPoints = questions.Cells(22, 8).Value
        Case "長期財富累積": totalPoints = questions.Cells(27, 8).Value
    End Select
    
    Do While N <= maxRow
        If questions.Cells(N, 1).Value = selectedCategory Then
            textblock.TextFrame.Characters.Text = questions.Cells(N, 2).Value
            Ablock.TextFrame.Characters.Text = questions.Cells(N, 3).Value
            Bblock.TextFrame.Characters.Text = questions.Cells(N, 4).Value
            Cblock.TextFrame.Characters.Text = questions.Cells(N, 5).Value
            Dblock.TextFrame.Characters.Text = questions.Cells(N, 6).Value
            Exit Do
        End If
        N = N + 1
    Loop
    
    If N > maxRow Then
        ' Activate result sheet based on totalPoints
        Select Case totalPoints
            Case 4 To 7: ThisWorkbook.Worksheets("風險厭惡").Activate
            Case 8 To 10: ThisWorkbook.Worksheets("風險中立偏厭惡").Activate
            Case 11 To 14: ThisWorkbook.Worksheets("風險中立").Activate
            Case 15 To 17: ThisWorkbook.Worksheets("風險中立偏愛好").Activate
            Case 18 To 20: ThisWorkbook.Worksheets("風險愛好").Activate
        End Select
        ClearQuestionAndOptions
        quizActive = False
        N = 2
    End If
End Sub

Sub Ablock_Click()
    If quizActive Then
        questions.Cells(N, 7).Value = 1
        N = N + 1
        ShowQuestion
    End If
End Sub

Sub Bblock_Click()
    If quizActive Then
        questions.Cells(N, 7).Value = 2
        N = N + 1
        ShowQuestion
    End If
End Sub

Sub Cblock_Click()
    If quizActive Then
        questions.Cells(N, 7).Value = 3
        N = N + 1
        ShowQuestion
    End If
End Sub

Sub Dblock_Click()
    If quizActive Then
        questions.Cells(N, 7).Value = 4
        N = N + 1
        ShowQuestion
    End If
End Sub

' Clear question text and answer options
Sub ClearQuestionAndOptions()
    textblock.TextFrame.Characters.Text = ""
    Ablock.TextFrame.Characters.Text = ""
    Bblock.TextFrame.Characters.Text = ""
    Cblock.TextFrame.Characters.Text = ""
    Dblock.TextFrame.Characters.Text = ""
End Sub

' Initialize the text boxes (question and answer options)
Sub InitializeTextBoxes()
    Set textblock = ActiveSheet.Shapes("question_block")
    Set Ablock = ActiveSheet.Shapes("A選項")
    Set Bblock = ActiveSheet.Shapes("B選項")
    Set Cblock = ActiveSheet.Shapes("C選項")
    Set Dblock = ActiveSheet.Shapes("D選項")
End Sub

' Button click handlers for each category
Sub 退休金_Click()
    StartQuiz "退休金"
End Sub

Sub 第一桶金_Click()
    StartQuiz "第一桶金"
End Sub

Sub 財產保值_Click()
    StartQuiz "財產保值"
End Sub

Sub 教育存款_Click()
    StartQuiz "教育存款"
End Sub

Sub 買房_Click()
    StartQuiz "買房"
End Sub

Sub 長期財富累積_Click()
    StartQuiz "長期財富累積"
End Sub

```

```vb
' Navigate to the result sheet (assumes result sheet is named as current sheet name + "結果單")
Sub nextpage_click()
    Dim nextpage As String
    nextpage = ActiveSheet.Name & "結果單"
    ThisWorkbook.Worksheets(nextpage).Activate
End Sub

```

```vb
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' Delay function (in seconds)
Function Delay(seconds As Single)
    Sleep seconds * 1000
End Function

Sub Button1_Click()
    Dim pythonExePath As String
    Dim pythonScriptPath As String
    Dim shellProcessID As Long
    Dim fileName As String

    ' Set path to Python executable and script file
    pythonExePath = """C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe"""
    fileName = "main.py"
    pythonScriptPath = ThisWorkbook.Path & "\" & fileName

    ' Execute Python script and get its process ID
    shellProcessID = Shell(pythonExePath & " " & pythonScriptPath, vbNormalFocus)

    ' Wait until the Python script finishes executing
    Do While IsProcessRunning(shellProcessID)
        DoEvents  ' Allow VBA to process other events
    Loop

    CloseSpecificWorkbook
End Sub

Function CloseSpecificWorkbook()
    Dim xlApp As Object
    Dim targetWorkbook As Object

    ' Get the running Excel application and close the specified workbook without saving
    Set xlApp = GetObject(, "Excel.Application")
    For Each targetWorkbook In xlApp.Workbooks
        If targetWorkbook.Name = "taiex_mid100_stock_data.xlsx" Then
            targetWorkbook.Close SaveChanges:=False
            Exit For
        End If
    Next targetWorkbook
End Function

' Check if a process with the given Process ID is still running
Function IsProcessRunning(PID As Long) As Boolean
    Dim objWMIService As Object, colProcessList As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & PID)
    IsProcessRunning = (colProcessList.Count > 0)
End Function

```

```vb
Sub Optimize_All()
    Generate_WeightsandData
    SharpeRatio_formula
End Sub

Sub Generate_WeightsandData()
    Dim ws As Worksheet
    Dim formulaString1 As String, formulaString2 As String, formulaString3 As String
    Dim formulaString4 As String, formulaString5 As String
    Dim cell As Range
    
    ' Set worksheet for optimization
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Copy range B1:K1 to AN1:AW1
    ws.Range("B1:K1").Copy Destination:=ws.Range("AN1")
    Application.CutCopyMode = False
    
    ' Append "W" to cells in range AN1:AW1 if not already present
    For Each cell In ws.Range("AN1:AW1")
        If Right(cell.Value, 1) <> "W" Then cell.Value = cell.Value & "W"
    Next cell
    
    ' Set range AN2:BB2 to zero and clear AN3:BB1001
    ws.Range("AN2:BB2").Value = 0
    ws.Range("AN3:BB1001").ClearContents
    
    ' Select cell AN3
    Sheets("幕後結果_最佳化").Select
    Range("AN3").Select
    
    ' Define formulas for weight calculations
    formulaString1 = "=B1004/SUM($B1004:$K1004)"  ' Weight calculation
    formulaString2 = "=SUM(AN3:AW3)"                ' Total weight
    formulaString3 = "=SUMPRODUCT(AN3:AW3,幕後結果_最佳化!$Z$1005:$AI$1005)"  ' Mean calculation
    formulaString4 = "=(SUMPRODUCT(AN3:AW3,MMULT(AN3:AW3,幕後結果!$AB$9:$AK$18)))*252"  ' Variance calculation
    formulaString5 = "=AZ3^0.5"                       ' Standard deviation
    
    ' Apply weight formula
    ws.Range("AN3").Formula2 = formulaString1
    Range("AN3").Copy
    Range("AO3:AW3").PasteSpecial
    Range("AN3:AW3").Copy
    Range("AN4:AW1001").PasteSpecial
    Application.CutCopyMode = False
    
    ' Apply total weight formula
    ws.Range("AX3").Formula2 = formulaString2
    ws.Range("AX3").AutoFill Destination:=ws.Range("AX3:AX1001"), Type:=xlFillDefault
    ws.Range("AX3:AX1001").FillDown
    
    ' Apply mean formula
    ws.Range("AY3").Formula2 = formulaString3
    ws.Range("AY3").AutoFill Destination:=ws.Range("AY3:AY1001"), Type:=xlFillDefault
    ws.Range("AY3:AY1001").FillDown
    
    ' Apply variance formula
    ws.Range("AZ3").Formula2 = formulaString4
    ws.Range("AZ3").AutoFill Destination:=ws.Range("AZ3:AZ1001"), Type:=xlFillDefault
    ws.Range("AZ3:AZ1001").FillDown
    
    ' Apply standard deviation formula
    ws.Range("BA3").Formula2 = formulaString5
    ws.Range("BA3").AutoFill Destination:=ws.Range("BA3:BA1001"), Type:=xlFillDefault
    ws.Range("BA3:BA1001").FillDown
End Sub

Sub SharpeRatio_formula()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sharpeRange As Range
    Dim meanValue As Double, stdDevValue As Double

    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    lastRow = 1001  ' Assume 1000 rows of data
    ws.Range("BB3").Formula = "=(AY3-$BC$2)/$BA3"
    ws.Range("BB3").AutoFill Destination:=ws.Range("BB3:BB" & lastRow), Type:=xlFillDefault
    ws.Range("BB3:BB" & lastRow).FillDown

    Set sharpeRange = ws.Range("BB3:BB" & lastRow)
    meanValue = Application.WorksheetFunction.Average(sharpeRange)
    stdDevValue = Application.WorksheetFunction.StDev(sharpeRange)
    
    Dim rng As Range
    Set rng = ws.Range("AM1:BB" & lastRow)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(125, 74, 43)  ' Dark brown
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(125, 74, 43)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(125, 74, 43)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(125, 74, 43)
    End With
End Sub

Sub FindMaxAndCopy()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim maxVal As Double, resultRange As Range
    Dim DestSheet As Worksheet

    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Color cells based on the maximum value
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102) ' Dark yellow
            Set resultRange = ws.Range("AN" & cell.Row & ":" & "BB" & cell.Row)
            ' Uncomment the following line to copy resultRange to another sheet if needed
            'DestSheet.Range("T3:AH3").Value = resultRange.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255) ' White (transparent)
        End If
    Next cell
End Sub

```

```vb
Sub 雪天型分析()
    ' Call related subs (assumed to paste stock codes, download data, and run simulation/optimization)
    貼上股票代碼_雪天型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Sheets("風險厭惡結果單")
    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    ' Set the range to check
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Loop through the range and set colors based on maximum value
    For Each cell In rng
        If cell.Value = maxVal Then
            ' If maximum, color from column AN to BB as dark yellow
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value
            
            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ' If not maximum, set cells from column K to T as white (transparent)
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

Sub 陰天型分析()
    貼上股票代碼_陰天型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Sheets("風險中立偏厭惡結果單")
    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    ' Set the range to check
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Loop through the range and set colors based on maximum value
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value
            
            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

Sub 晴天型分析()
    貼上股票代碼_晴天型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Sheets("風險中立結果單")
    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    ' Set the range to check
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Loop through the range and set colors based on maximum value
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value
            
            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

Sub 雷雨型分析()
    貼上股票代碼_雷雨型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Sheets("風險中立偏愛好結果單")
    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    ' Set the range to check
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Loop through the range and set colors based on maximum value
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value
            
            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

Sub 閃電型分析()
    貼上股票代碼_閃電型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Sheets("風險愛好結果單")
    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    ' Set the range to check
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Loop through the range and set colors based on maximum value
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value
            
            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

```

```vb
' ---------------------------
' Snow Day Final Analysis
' ---------------------------
Sub 雪天型最終分析()
    貼上股票代碼_雪天型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Worksheets("風險厭惡結果單")
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    ' Loop through range and set cell colors
    For Each cell In rng
        If cell.Value = maxVal Then
            ' Color cells from AN to BB dark yellow if they contain the max value
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value

            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ' Otherwise, set cells from K to T to white
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

' ---------------------------
' Cloudy Day Final Analysis
' ---------------------------
Sub 陰天型分析()
    貼上股票代碼_陰天型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Worksheets("風險中立偏厭惡結果單")
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value

            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

' ---------------------------
' Sunny Day Final Analysis
' ---------------------------
Sub 晴天型分析()
    貼上股票代碼_晴天型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Worksheets("風險中立結果單")
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value
    maxVal = Application.WorksheetFunction.Max(rng)
    
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value

            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

' ---------------------------
' Thunderstorm Day Final Analysis
' ---------------------------
Sub 雷雨型分析()
    貼上股票代碼_雷雨型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Worksheets("風險中立偏愛好結果單")
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)
    
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value

            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

' ---------------------------
' Lightning Day Final Analysis
' ---------------------------
Sub 閃電型最終分析()
    貼上股票代碼_閃電型
    下載_ALL
    Simulation_and_Optimize_All

    Dim ws As Worksheet, rng As Range, cell As Range
    Dim maxVal As Double, resultRange1 As Range, resultRange2 As Range
    Dim DestSheet As Worksheet

    ' Set target worksheet
    Set DestSheet = ThisWorkbook.Worksheets("風險愛好結果單")
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    Set rng = ws.Range("BB3:BB1001")
    
    ' Find maximum value
    maxVal = Application.WorksheetFunction.Max(rng)
    
    For Each cell In rng
        If cell.Value = maxVal Then
            ws.Range(ws.Cells(cell.Row, "AN"), ws.Cells(cell.Row, "BB")).Interior.Color = RGB(255, 217, 102)
            Set resultRange1 = ws.Range("AN" & cell.Row & ":" & "AW" & cell.Row)
            DestSheet.Range("H33:Q33").Value = resultRange1.Value

            Set resultRange2 = ws.Range("AY" & cell.Row & ":" & "BB" & cell.Row)
            DestSheet.Range("H34:K34").Value = resultRange2.Value
        Else
            ws.Range(ws.Cells(cell.Row, "K"), ws.Cells(cell.Row, "T")).Interior.Color = RGB(255, 255, 255)
        End If
    Next cell
    DestSheet.Activate
End Sub

' ---------------------------
' Create "Risk Aversion Result" Sheet for Snow Day
' ---------------------------
Sub create風險厭惡結果單()
    Dim mean As Shape, var As Shape, Std As Shape, Sharpe As Shape
    Dim rng2 As Range
    Dim cellValue1 As String, cellValue2 As String, cellValue3 As String, cellValue4 As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("風險厭惡結果單")
    
    ' Initialize shapes for key metrics
    Set mean = ActiveSheet.Shapes("平均數")
    Set var = ActiveSheet.Shapes("變異數")
    Set Std = ActiveSheet.Shapes("標準差")
    Set Sharpe = ActiveSheet.Shapes("夏普率")
    
    ' Set the range to process (H34:K34)
    Set rng2 = ws.Range("H34:K34")
    
    ' Retrieve values from cells
    cellValue1 = ws.Range("H34").Value
    cellValue2 = ws.Range("I34").Value
    cellValue3 = ws.Range("J34").Value
    cellValue4 = ws.Range("K34").Value
    
    ' Update shape texts
    mean.TextFrame.Characters.Text = cellValue1
    var.TextFrame.Characters.Text = cellValue2
    Std.TextFrame.Characters.Text = cellValue3
    Sharpe.TextFrame.Characters.Text = cellValue4
End Sub

' ---------------------------
' Generate Pie Chart for Snow Day
' ---------------------------
Sub GeneratePieChart_雪天型()
    Dim ws1 As Worksheet, ws2 As Worksheet, chartObj As ChartObject, chart As Chart
    Dim dataRange As Range, formulaString1 As String, targetRange As Range
    Dim i As Integer

    Set ws1 = ThisWorkbook.Worksheets("風險厭惡結果單")
    Set ws2 = ThisWorkbook.Worksheets("幕後結果_最佳化")

    ' Clear all charts on the worksheet
    For Each chartObj In ws1.ChartObjects
        chartObj.Delete
    Next chartObj

    ' Set data range (H35:Q35) and apply formula for percentage calculation
    Set dataRange = ws1.Range("H35:Q35")
    formulaString1 = "=H$33*幕後結果_最佳化!Z$1005/$H$34"
    Set targetRange = ws1.Range("H35:Q35")
    targetRange.Formula = formulaString1

    ' Create an embedded pie chart
    Set chartObj = ws1.ChartObjects.Add(Left:=450, Width:=375, Top:=550, Height:=245)
    Set chart = chartObj.chart

    With chart
        .ChartType = xlPie
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = "報酬率佔比圖"  ' "Return Percentage Chart"
        .HasLegend = False
        With .SeriesCollection(1)
            .HasDataLabels = True
            .ApplyDataLabels xlDataLabelsShowValue
            For i = 1 To .Points.Count
                .Points(i).DataLabel.Text = ws2.Cells(1, i + 1).Value
            Next i
        End With
    End With
End Sub

' ---------------------------
' Cloudy Day Final Analysis (Alternate)
' ---------------------------
Sub 陰天型最終分析()
    create風險中立偏厭惡結果單
    GeneratePieChart_陰天型
End Sub

Sub create風險中立偏厭惡結果單()
    Dim mean As Shape, var As Shape, Std As Shape, Sharpe As Shape
    Dim rng2 As Range
    Dim cellValue1 As String, cellValue2 As String, cellValue3 As String, cellValue4 As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("風險中立偏厭惡結果單")
    
    ' Initialize shapes for key metrics
    Set mean = ActiveSheet.Shapes("平均數")
    Set var = ActiveSheet.Shapes("變異數")
    Set Std = ActiveSheet.Shapes("標準差")
    Set Sharpe = ActiveSheet.Shapes("夏普率")
    
    ' Set the range (H34:K34) and update shape texts
    Set rng2 = ws.Range("H34:K34")
    cellValue1 = ws.Range("H34").Value
    cellValue2 = ws.Range("I34").Value
    cellValue3 = ws.Range("J34").Value
    cellValue4 = ws.Range("K34").Value
    mean.TextFrame.Characters.Text = cellValue1
    var.TextFrame.Characters.Text = cellValue2
    Std.TextFrame.Characters.Text = cellValue3
    Sharpe.TextFrame.Characters.Text = cellValue4
End Sub

Sub GeneratePieChart_陰天型()
    Dim ws1 As Worksheet, ws2 As Worksheet, chartObj As ChartObject, chart As Chart
    Dim dataRange As Range, formulaString1 As String, targetRange As Range
    Dim i As Integer

    Set ws1 = ThisWorkbook.Worksheets("風險中立偏厭惡結果單")
    Set ws2 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    For Each chartObj In ws1.ChartObjects
        chartObj.Delete
    Next chartObj

    Set dataRange = ws1.Range("H35:Q35")
    ReDim dataRange ' Not strictly necessary; proportions/weights arrays are not used here.
    formulaString1 = "=H$33*幕後結果_最佳化!Z$1005/$H$34"
    Set targetRange = ws1.Range("H35:Q35")
    targetRange.Formula = formulaString1

    Set chartObj = ws1.ChartObjects.Add(Left:=450, Width:=375, Top:=550, Height:=255)
    Set chart = chartObj.chart

    With chart
        .ChartType = xlPie
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = "報酬率佔比圖"
        .HasLegend = False
        With .SeriesCollection(1)
            .HasDataLabels = True
            .ApplyDataLabels xlDataLabelsShowValue
            For i = 1 To .Points.Count
                .Points(i).DataLabel.Text = ws2.Cells(1, i + 1).Value
            Next i
        End With
    End With
End Sub

' ---------------------------
' Sunny Day Final Analysis (Alternate)
' ---------------------------
Sub 晴天型最終分析()
    create風險中立結果單
    GeneratePieChart_晴天型
End Sub

Sub create風險中立結果單()
    Dim mean As Shape, var As Shape, Std As Shape, Sharpe As Shape
    Dim rng2 As Range
    Dim cellValue1 As String, cellValue2 As String, cellValue3 As String, cellValue4 As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("風險中立結果單")
    
    ' Initialize shapes for key metrics
    Set mean = ActiveSheet.Shapes("平均數")
    Set var = ActiveSheet.Shapes("變異數")
    Set Std = ActiveSheet.Shapes("標準差")
    Set Sharpe = ActiveSheet.Shapes("夏普率")
    
    ' Set range (H34:K34) and update shape texts
    Set rng2 = ws.Range("H34:K34")
    cellValue1 = ws.Range("H34").Value
    cellValue2 = ws.Range("I34").Value
    cellValue3 = ws.Range("J34").Value
    cellValue4 = ws.Range("K34").Value
    mean.TextFrame.Characters.Text = cellValue1
    var.TextFrame.Characters.Text = cellValue2
    Std.TextFrame.Characters.Text = cellValue3
    Sharpe.TextFrame.Characters.Text = cellValue4
End Sub

Sub GeneratePieChart_晴天型()
    Dim ws1 As Worksheet, ws2 As Worksheet, chartObj As ChartObject, chart As Chart
    Dim dataRange As Range, formulaString1 As String, targetRange As Range
    Dim i As Integer

    Set ws1 = ThisWorkbook.Worksheets("風險中立結果單")
    Set ws2 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    For Each chartObj In ws1.ChartObjects
        chartObj.Delete
    Next chartObj

    Set dataRange = ws1.Range("H35:Q35")
    formulaString1 = "=H$33*幕後結果_最佳化!Z$1005/$H$34"
    Set targetRange = ws1.Range("H35:Q35")
    targetRange.Formula = formulaString1

    Set chartObj = ws1.ChartObjects.Add(Left:=450, Width:=375, Top:=550, Height:=255)
    Set chart = chartObj.chart

    With chart
        .ChartType = xlPie
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = "報酬率佔比圖"
        .HasLegend = False
        With .SeriesCollection(1)
            .HasDataLabels = True
            .ApplyDataLabels xlDataLabelsShowValue
            For i = 1 To .Points.Count
                .Points(i).DataLabel.Text = ws2.Cells(1, i + 1).Value
            Next i
        End With
    End With
End Sub

' ---------------------------
' Thunderstorm Day Final Analysis (Alternate)
' ---------------------------
Sub 雷雨型最終分析()
    create風險中立偏愛好結果單
    GeneratePieChart_雷雨型
End Sub

Sub create風險中立偏愛好結果單()
    Dim mean As Shape, var As Shape, Std As Shape, Sharpe As Shape
    Dim rng2 As Range
    Dim cellValue1 As String, cellValue2 As String, cellValue3 As String, cellValue4 As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("風險中立偏愛好結果單")
    
    ' Initialize shapes for key metrics
    Set mean = ActiveSheet.Shapes("平均數")
    Set var = ActiveSheet.Shapes("變異數")
    Set Std = ActiveSheet.Shapes("標準差")
    Set Sharpe = ActiveSheet.Shapes("夏普率")
    
    ' Set range (H34:K34) and update shape texts
    Set rng2 = ws.Range("H34:K34")
    cellValue1 = ws.Range("H34").Value
    cellValue2 = ws.Range("I34").Value
    cellValue3 = ws.Range("J34").Value
    cellValue4 = ws.Range("K34").Value
    mean.TextFrame.Characters.Text = cellValue1
    var.TextFrame.Characters.Text = cellValue2
    Std.TextFrame.Characters.Text = cellValue3
    Sharpe.TextFrame.Characters.Text = cellValue4
End Sub

Sub GeneratePieChart_雷雨型()
    Dim ws1 As Worksheet, ws2 As Worksheet, chartObj As ChartObject, chart As Chart
    Dim dataRange As Range, formulaString1 As String, targetRange As Range
    Dim i As Integer

    Set ws1 = ThisWorkbook.Worksheets("風險中立偏愛好結果單")
    Set ws2 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    For Each chartObj In ws1.ChartObjects
        chartObj.Delete
    Next chartObj

    Set dataRange = ws1.Range("H35:Q35")
    formulaString1 = "=H$33*幕後結果_最佳化!Z$1005/$H$34"
    Set targetRange = ws1.Range("H35:Q35")
    targetRange.Formula = formulaString1

    Set chartObj = ws1.ChartObjects.Add(Left:=450, Width:=375, Top:=550, Height:=255)
    Set chart = chartObj.chart

    With chart
        .ChartType = xlPie
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = "報酬率佔比圖"
        .HasLegend = False
        With .SeriesCollection(1)
            .HasDataLabels = True
            .ApplyDataLabels xlDataLabelsShowValue
            For i = 1 To .Points.Count
                .Points(i).DataLabel.Text = ws2.Cells(1, i + 1).Value
            Next i
        End With
    End With
End Sub

' ---------------------------
' Lightning Day Final Analysis (Alternate)
' ---------------------------
Sub 閃電型最終分析()
    create風險愛好結果單
    GeneratePieChart_閃電型
End Sub

Sub create風險愛好結果單()
    Dim mean As Shape, var As Shape, Std As Shape, Sharpe As Shape
    Dim rng2 As Range
    Dim cellValue1 As String, cellValue2 As String, cellValue3 As String, cellValue4 As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("風險愛好結果單")
    
    ' Initialize shapes for key metrics
    Set mean = ActiveSheet.Shapes("平均數")
    Set var = ActiveSheet.Shapes("變異數")
    Set Std = ActiveSheet.Shapes("標準差")
    Set Sharpe = ActiveSheet.Shapes("夏普率")
    
    ' Set range (H34:K34) and update shape texts
    Set rng2 = ws.Range("H34:K34")
    cellValue1 = ws.Range("H34").Value
    cellValue2 = ws.Range("I34").Value
    cellValue3 = ws.Range("J34").Value
    cellValue4 = ws.Range("K34").Value
    mean.TextFrame.Characters.Text = cellValue1
    var.TextFrame.Characters.Text = cellValue2
    Std.TextFrame.Characters.Text = cellValue3
    Sharpe.TextFrame.Characters.Text = cellValue4
End Sub

Sub GeneratePieChart_閃電型()
    Dim ws1 As Worksheet, ws2 As Worksheet, chartObj As ChartObject, chart As Chart
    Dim dataRange As Range, formulaString1 As String, targetRange As Range
    Dim i As Integer

    Set ws1 = ThisWorkbook.Worksheets("風險愛好結果單")
    Set ws2 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    For Each chartObj In ws1.ChartObjects
        chartObj.Delete
    Next chartObj

    Set dataRange = ws1.Range("H35:Q35")
    formulaString1 = "=H$33*幕後結果_最佳化!Z$1005/$H$34"
    Set targetRange = ws1.Range("H35:Q35")
    targetRange.Formula = formulaString1

    Set chartObj = ws1.ChartObjects.Add(Left:=450, Width:=375, Top:=550, Height:=255)
    Set chart = chartObj.chart

    With chart
        .ChartType = xlPie
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = "報酬率佔比圖"
        .HasLegend = False
        With .SeriesCollection(1)
            .HasDataLabels = True
            .ApplyDataLabels xlDataLabelsShowValue
            For i = 1 To .Points.Count
                .Points(i).DataLabel.Text = ws2.Cells(1, i + 1).Value
            Next i
        End With
    End With
End Sub

```

```vb
Sub 矩形圓角3_Click()
    ' Activate the POS button sheet
    Sheets("POS機按鈕").Activate
End Sub

Sub 返回風險中立偏厭人格頁_Click()
    ' Activate the "風險中立偏厭惡" sheet
    Sheets("風險中立偏厭惡").Activate
End Sub

Sub 返回風險厭惡人格介面_Click()
    ' Activate the "風險厭惡" sheet
    Sheets("風險厭惡").Activate
End Sub

Sub 返回風險中立人格頁面_Click()
    ' Activate the "風險中立" sheet
    Sheets("風險中立").Activate
End Sub

Sub 返回風險中立偏愛好人格頁_Click()
    ' Activate the "風險中立偏愛好" sheet
    Sheets("風險中立偏愛好").Activate
End Sub

Sub 返回風險愛好人格頁_Click()
    ' Activate the "風險愛好" sheet
    Sheets("風險愛好").Activate
End Sub

Sub nextpage_click()
    Dim nextpage As String
    ' Construct the name of the result sheet (current sheet name + "結果單")
    nextpage = ActiveSheet.Name & "結果單"
    ThisWorkbook.Worksheets(nextpage).Activate
End Sub

```

```vb
Sub 列印風險厭惡結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape

    ' Get Desktop path
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    ' Define PDF file names and paths
    PDFName1 = "風險厭惡結果單.pdf"
    PDFName2 = "風險厭惡.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    ' Set worksheets
    Set ws1 = Sheets("風險厭惡結果單")
    Set ws2 = Sheets("風險厭惡")
    
    ' Get shapes to hide during export
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險厭惡結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    ' Hide shapes
    shp1.Visible = msoFalse
    shp2.Visible = msoFalse
    shp3.Visible = msoFalse
    shp4.Visible = msoFalse
    shp5.Visible = msoFalse

    ' Export specified range from ws1 as PDF
    With ws1.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws1.Range("E2:S56").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard
    
    ' Export specified range from ws2 as PDF
    With ws2.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    ' Notify user of saved PDFs
    MsgBox "PDF files have been saved to Desktop:" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    ' Restore shapes' visibility
    shp1.Visible = msoTrue
    shp2.Visible = msoTrue
    shp3.Visible = msoTrue
    shp4.Visible = msoTrue
    shp5.Visible = msoTrue
End Sub

Sub 列印風險中立偏厭惡結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape

    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    PDFName1 = "風險中立偏厭惡結果單.pdf"
    PDFName2 = "風險中立偏厭惡.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    Set ws1 = Sheets("風險中立偏厭惡結果單")
    Set ws2 = Sheets("風險中立偏厭惡")
    
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險中偏厭惡結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    shp1.Visible = msoFalse
    shp2.Visible = msoFalse
    shp3.Visible = msoFalse
    shp4.Visible = msoFalse
    shp5.Visible = msoFalse

    With ws1.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws1.Range("E2:S56").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard
    
    With ws2.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    MsgBox "PDF files have been saved to Desktop:" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    shp1.Visible = msoTrue
    shp2.Visible = msoTrue
    shp3.Visible = msoTrue
    shp4.Visible = msoTrue
    shp5.Visible = msoTrue
End Sub

Sub 列印風險中立結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape

    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    PDFName1 = "風險中立結果單.pdf"
    PDFName2 = "風險中立.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    Set ws1 = Sheets("風險中立結果單")
    Set ws2 = Sheets("風險中立")
    
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險中立結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    shp1.Visible = msoFalse
    shp2.Visible = msoFalse
    shp3.Visible = msoFalse
    shp4.Visible = msoFalse
    shp5.Visible = msoFalse

    With ws1.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws1.Range("E2:S55").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard
    
    With ws2.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    MsgBox "PDF files have been saved to Desktop:" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    shp1.Visible = msoTrue
    shp2.Visible = msoTrue
    shp3.Visible = msoTrue
    shp4.Visible = msoTrue
    shp5.Visible = msoTrue
End Sub

Sub 列印風險中立偏愛好結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape

    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    PDFName1 = "風險中立偏愛好結果單.pdf"
    PDFName2 = "風險中立偏愛好.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    Set ws1 = Sheets("風險中立偏愛好結果單")
    Set ws2 = Sheets("風險中立偏愛好")
    
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險中立偏愛好結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    shp1.Visible = msoFalse
    shp2.Visible = msoFalse
    shp3.Visible = msoFalse
    shp4.Visible = msoFalse
    shp5.Visible = msoFalse

    With ws1.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws1.Range("E2:S55").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard
    
    With ws2.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    MsgBox "PDF files have been saved to Desktop:" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    shp1.Visible = msoTrue
    shp2.Visible = msoTrue
    shp3.Visible = msoTrue
    shp4.Visible = msoTrue
    shp5.Visible = msoTrue
End Sub

Sub 列印風險愛好結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape

    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    PDFName1 = "風險厭惡結果單.pdf"
    PDFName2 = "風險厭惡.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    Set ws1 = Sheets("風險厭惡結果單")
    Set ws2 = Sheets("風險厭惡")
    
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險厭惡結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    shp1.Visible = msoFalse
    shp2.Visible = msoFalse
    shp3.Visible = msoFalse
    shp4.Visible = msoFalse
    shp5.Visible = msoFalse

    With ws1.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws1.Range("E2:S55").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard

    With ws2.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    MsgBox "PDF files have been saved to Desktop:" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    shp1.Visible = msoTrue
    shp2.Visible = msoTrue
    shp3.Visible = msoTrue
    shp4.Visible = msoTrue
    shp5.Visible = msoTrue
End Sub
```

```vb
'-------------------------------------------
' Start and Stock Code Paste for Snow Day
'-------------------------------------------
Sub start()
    ' Show the "Select Celebrity" form
    選擇名人.Show
End Sub

Sub 貼上股票代碼_雪天型()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim i As Long, j As Long, resultRow As Long
    Dim StockCode As String
    
    ' Set source and destination worksheets
    Set ws1 = ThisWorkbook.Sheets("雪天型結果")  ' Source sheet
    Set ws2 = ThisWorkbook.Sheets("幕後結果")     ' Destination sheet
    
    resultRow = 2  ' Start pasting at row 2
    
    ' Loop through columns A to C in the source sheet
    For j = 1 To 3
        For i = 2 To ws1.Cells(Rows.Count, j).End(xlUp).Row
            If ws1.Cells(i, j).Value <> "" Then
                StockCode = Replace(ws1.Cells(i, j).Value, ".TW", "") ' Remove ".TW"
                With ws2.Cells(resultRow, 1)
                    .Value = StockCode  ' Paste to column A in destination
                    .Interior.Color = ws1.Cells(i, j).Interior.Color  ' Copy cell color
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                resultRow = resultRow + 1
                If resultRow > 11 Then Exit For
            End If
        Next i
        If resultRow > 11 Then Exit For
    Next j
    
    ' If fewer than 10 rows were pasted, fill the remaining with blanks
    While resultRow <= 11
        With ws2.Cells(resultRow, 1)
            .Value = ""
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        resultRow = resultRow + 1
    Wend
End Sub

Sub 貼上股票代碼_陰天型()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim i As Long, j As Long, resultRow As Long
    Dim StockCode As String
    
    Set ws1 = ThisWorkbook.Sheets("陰天型結果")
    Set ws2 = ThisWorkbook.Sheets("幕後結果")
    
    resultRow = 2
    For j = 1 To 3
        For i = 2 To ws1.Cells(Rows.Count, j).End(xlUp).Row
            If ws1.Cells(i, j).Value <> "" Then
                StockCode = Replace(ws1.Cells(i, j).Value, ".TW", "")
                With ws2.Cells(resultRow, 1)
                    .Value = StockCode
                    .Interior.Color = ws1.Cells(i, j).Interior.Color
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                resultRow = resultRow + 1
                If resultRow > 11 Then Exit For
            End If
        Next i
        If resultRow > 11 Then Exit For
    Next j
    
    While resultRow <= 11
        With ws2.Cells(resultRow, 1)
            .Value = ""
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        resultRow = resultRow + 1
    Wend
End Sub

Sub 貼上股票代碼_晴天型()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim i As Long, j As Long, resultRow As Long
    Dim StockCode As String
    
    Set ws1 = ThisWorkbook.Sheets("晴天型結果")
    Set ws2 = ThisWorkbook.Sheets("幕後結果")
    
    resultRow = 2
    For j = 1 To 3
        For i = 2 To ws1.Cells(Rows.Count, j).End(xlUp).Row
            If ws1.Cells(i, j).Value <> "" Then
                StockCode = Replace(ws1.Cells(i, j).Value, ".TW", "")
                With ws2.Cells(resultRow, 1)
                    .Value = StockCode
                    .Interior.Color = ws1.Cells(i, j).Interior.Color
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                resultRow = resultRow + 1
                If resultRow > 11 Then Exit For
            End If
        Next i
        If resultRow > 11 Then Exit For
    Next j
    
    While resultRow <= 11
        With ws2.Cells(resultRow, 1)
            .Value = ""
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        resultRow = resultRow + 1
    Wend
End Sub

Sub 貼上股票代碼_雷雨型()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim i As Long, j As Long, resultRow As Long
    Dim StockCode As String
    
    Set ws1 = ThisWorkbook.Sheets("雷雨型結果")
    Set ws2 = ThisWorkbook.Sheets("幕後結果")
    
    resultRow = 2
    For j = 1 To 3
        For i = 2 To ws1.Cells(Rows.Count, j).End(xlUp).Row
            If ws1.Cells(i, j).Value <> "" Then
                StockCode = Replace(ws1.Cells(i, j).Value, ".TW", "")
                With ws2.Cells(resultRow, 1)
                    .Value = StockCode
                    .Interior.Color = ws1.Cells(i, j).Interior.Color
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                resultRow = resultRow + 1
                If resultRow > 11 Then Exit For
            End If
        Next i
        If resultRow > 11 Then Exit For
    Next j
    
    While resultRow <= 11
        With ws2.Cells(resultRow, 1)
            .Value = ""
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        resultRow = resultRow + 1
    Wend
End Sub

'-------------------------------------------
' Download and Data Processing Module
'-------------------------------------------
Sub 下載_ALL()
    ' Call various subs to delete old sheets, download data, process data, etc.
    刪除原併表
    download
    Data_Processing
    FillStockCodes
    SelectAndCopyDates
    SelectAndCopyClosingPrices
    LN_ClosingPrices
    Matrix_ALL
End Sub

Sub 刪除原併表()
    Dim ws As Worksheet, sheetName As String
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        sheetName = ws.Name
        ' Delete sheets with numeric names (excluding protected ones)
        If IsNumeric(sheetName) And sheetName <> "medium_risk" And sheetName <> "high_risk" And sheetName <> "stock price" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

Sub download()
    ' Download macro to remove sheets and add queries
    Dim ds As Worksheet
    Application.DisplayAlerts = False
    For Each ds In ThisWorkbook.Worksheets
        If Not ds Is ActiveSheet Then
            If IsNumeric(ds.Name) And Len(ds.Name) > 6 And ds.Name <> "medium_risk" And ds.Name <> "high_risk" And ds.Name <> "stock price" Then
                ds.Delete
            End If
        End If
    Next ds
    Application.DisplayAlerts = True

    Dim i As Integer, j As Integer
    i = 2
    Set originalSheet = ActiveSheet
    Dim query As WorkbookQuery
    For Each query In ActiveWorkbook.Queries
        query.Delete
    Next query

    Dim S As String, Y As String, M As String
    Dim M_1 As String, M_2 As String, M_3 As String
    Dim currentYear As Integer, currentMonth As Integer
    
    currentYear = Year(Date)
    currentMonth = Month(Date)
    
    If currentMonth > 3 Then
        Y = CStr(currentYear - 1911)
        M = CStr(currentMonth - 3)
    Else
        Y = CStr(currentYear - 1912)
        M = CStr(12 + currentMonth - 3)
    End If
    
    ' Calculate month strings for three consecutive months
    M_1 = Format((CInt(M) + 12) Mod 12, "00")
    If M_1 = "00" Then M_1 = "12"
    M_2 = Format((CInt(M) + 1 + 12) Mod 12, "00")
    If M_2 = "00" Then M_2 = "12"
    M_3 = Format((CInt(M) + 2 + 12) Mod 12, "00")
    If M_3 = "00" Then M_3 = "12"

    While Not IsEmpty(Worksheets("幕後結果").Cells(i, 1))
        S = Worksheets("幕後結果").Cells(i, 1)
        For j = 1 To 3
            Select Case j
                Case 1: M = M_3
                Case 2: M = M_2
                Case 3: M = M_1
            End Select
        
            ActiveWorkbook.Queries.Add Name:="Table_" & S & "_" & Y & "_" & M, _
                Formula:= "let" & vbCrLf & _
                          "    Source = Web.Page(Web.Contents(""https://stock.wearn.com/cdata.asp?Year=" & Y & "&month=" & M & "&kind=" & S & """))," & vbCrLf & _
                          "    Data0 = Source{0}[Data]," & vbCrLf & _
                          "    ChangedType = Table.TransformColumnTypes(Data0,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}, {""Column5"", type text}, {""Column6"", type text}})" & vbCrLf & _
                          "in" & vbCrLf & _
                          "    ChangedType"
        
            Dim wsTemp As Worksheet
            Set wsTemp = ActiveWorkbook.Worksheets.Add
            wsTemp.Name = S & Y & M
            
            With wsTemp.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table_" & S & "_" & Y & "_" & M & """;Extended Properties=""""", Destination:=wsTemp.Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [Table_" & S & "_" & Y & "_" & M & "]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "Table_" & S & "_" & Y & "_" & M
                .Refresh BackgroundQuery:=False
            End With
        Next j
        originalSheet.Activate
        i = i + 1
    Wend
End Sub

Sub Data_Processing()
    Dim wb As Workbook, ws As Worksheet, sheetName As String
    Dim i As Long, j As Long, K As Long, N As Long, M As Long, r As Long, c As Long
    Dim CheckIfSheetExists As Boolean
    
    Set wb = ThisWorkbook
    
    ' Collect data from sheets with numeric names and store in DataTemp array
    i = 1
    Dim DataTemp(1 To 100, 3) As Variant
    For Each ws In wb.Worksheets
        sheetName = ws.Name
        If IsNumeric(sheetName) And Len(sheetName) >= 9 And Len(sheetName) <= 11 Then
            With ws
                Select Case Len(sheetName)
                    Case 9: DataTemp(i, 1) = FormatStockCode(Mid(.Cells(2, "A"), 9, 4))
                    Case 10: DataTemp(i, 1) = FormatStockCode(Mid(.Cells(2, "A"), 10, 5))
                    Case 11: DataTemp(i, 1) = FormatStockCode(Mid(.Cells(2, "A"), 11, 6))
                End Select
                DataTemp(i, 2) = .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).Value
                DataTemp(i, 3) = sheetName
                i = i + 1
            End With
        End If
    Next ws
    
    ' Create new sheets for unique stock codes
    For j = 1 To UBound(DataTemp)
        sheetName = DataTemp(j, 1)
        If sheetName <> "" Then
            CheckIfSheetExists = False
            For Each ws In wb.Worksheets
                If sheetName = ws.Name Then
                    CheckIfSheetExists = True
                    Exit For
                End If
            Next ws
            If Not CheckIfSheetExists Then wb.Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
        End If
    Next j
    
    ' Delete old sheets
    Application.DisplayAlerts = False
    For j = 1 To UBound(DataTemp)
        sheetName = DataTemp(j, 3)
        If sheetName <> "" Then wb.Sheets(sheetName).Delete
    Next j
    Application.DisplayAlerts = True
    
    ' Paste collected data into new sheets
    For j = 1 To UBound(DataTemp)
        sheetName = DataTemp(j, 1)
        If sheetName <> "" Then
            N = UBound(DataTemp(j, 2), 1)
            M = UBound(DataTemp(j, 2), 2)
            With wb.Sheets(sheetName)
                If WorksheetFunction.CountA(.UsedRange) = 0 Then
                    .Range(.Cells(1, 1), .Cells(N, M)).Value = DataTemp(j, 2)
                Else
                    r = .UsedRange.Rows.Count
                    .Range(.Cells(r + 1, 1), .Cells(N + r, M)).Value = DataTemp(j, 2)
                End If
            End With
        End If
    Next j
    
    ' Find unique stock codes
    i = 1
    Dim UniqueSheet(1 To 100) As String
    For j = 1 To UBound(DataTemp)
        CheckIfSheetExists = False
        For K = 1 To UBound(UniqueSheet)
            If UniqueSheet(K) = DataTemp(j, 1) Then
                CheckIfSheetExists = True
                Exit For
            End If
        Next K
        If Not CheckIfSheetExists Then
            UniqueSheet(i) = DataTemp(j, 1)
            i = i + 1
        End If
    Next j
    
    ' Remove duplicates and sort dates for each unique sheet
    For j = 1 To UBound(UniqueSheet)
        sheetName = UniqueSheet(j)
        If sheetName <> "" Then
            With wb.Sheets(sheetName)
                r = .UsedRange.Rows.Count
                c = .UsedRange.Columns.Count
                .Range(.Cells(1, 1), .Cells(r, c)).RemoveDuplicates Columns:=Array(1, c)
                .Rows(2).Delete
                .Rows(1).Delete
                .Activate
                .Range(.Cells(1, 1), .Cells(r, c)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
            End With
        End If
    Next j
End Sub

Sub FillStockCodes()
    Dim ws As Worksheet, i As Integer
    Set ws = ThisWorkbook.Sheets("幕後結果")
    ws.Range("D10:M10").ClearContents
    For i = 2 To 11
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            ws.Cells(1, i + 2).Value = ws.Cells(i, 1).Value
        End If
    Next i
End Sub

Sub SelectAndCopyDates()
    Dim wsSource As Worksheet, wsDestination As Worksheet, lastRow As Long, dateRange As Range, ws As Worksheet, found As Boolean
    found = False
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            Set wsSource = ws
            found = True
            Exit For
        End If
    Next ws
    Set wsDestination = ThisWorkbook.Worksheets("幕後結果")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    wsDestination.Range("C2:C" & wsDestination.Cells(wsDestination.Rows.Count, "C").End(xlUp).Row).ClearContents
    Set dateRange = wsSource.Range("A2:A" & lastRow)
    dateRange.Copy
    wsDestination.Range("C2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

Sub SelectAndCopyClosingPrices()
    Dim wsDestination As Worksheet, ws As Worksheet, lastRow As Long, ClosingPricesRange As Range, i As Long
    Set wsDestination = ThisWorkbook.Worksheets("幕後結果")
    wsDestination.Range("D2:M11").ClearContents
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            If i > 10 Then Exit For
            lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
            Set ClosingPricesRange = ws.Range("E2:E" & lastRow)
            ClosingPricesRange.Copy
            wsDestination.Range("D2").Offset(0, i - 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            i = i + 1
        End If
    Next ws
End Sub

Sub LN_ClosingPrices()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("幕後結果")
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ws.Range("D1:M1").Copy Destination:=ws.Range("P1")
    Application.CutCopyMode = False
    ws.Range("D1:M1").Copy Destination:=ws.Range("AB1")
    Application.CutCopyMode = False
    ws.Range("P2:Y2").Value = 0
    ws.Range("P3:Y" & lastRow).Formula = "=LN(D3/D2)"
    Application.CutCopyMode = False
End Sub

Sub Datas_LN_ClosingPrices()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("幕後結果")
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ws.Range("D1:M1").Copy Destination:=ws.Range("AB1")
    Application.CutCopyMode = False
    ws.Range("AB2").Formula = "=AVERAGE(P3:P" & lastRow & ")"
    ws.Range("AB2").AutoFill Destination:=ws.Range("AB2:AK2"), Type:=xlFillDefault
    ws.Range("AB3").Formula = "=VAR(P3:P" & lastRow & ")"
    ws.Range("AB3").AutoFill Destination:=ws.Range("AB3:AK3"), Type:=xlFillDefault
    ws.Range("AB4").Formula = "=STDEV(P3:P" & lastRow & ")"
    ws.Range("AB4").AutoFill Destination:=ws.Range("AB4:AK4"), Type:=xlFillDefault
    ws.Range("AB5").Formula = "=STDEV(P3:P" & lastRow & ")*SQRT(4)"
    ws.Range("AB5").AutoFill Destination:=ws.Range("AB5:AK5"), Type:=xlFillDefault
End Sub

Function FormatStockCode(ByVal code As String) As String
    Dim formattedCode As String
    If Len(code) = 4 Then
        formattedCode = code
    ElseIf Len(code) = 5 Then
        formattedCode = "0" & code
    ElseIf Len(code) = 6 Then
        formattedCode = "00" & code
    End If
    FormatStockCode = formattedCode
End Function
            
Function exists_or_not(Stock As String) As Boolean
    On Error Resume Next
    exists_or_not = Not Worksheets(Stock) Is Nothing
    On Error GoTo 0
End Function

```

```vb
Sub Matrix_ALL()
    Variance_Matrix
    Correlation_Matrix
    CholeskyDecompose_Matrix
End Sub

Function VarCovar(rng As Range, target As Range) As Variant
    Dim i As Integer, j As Integer, numcols As Integer
    numcols = rng.Columns.Count
    Dim matrix() As Variant
    ReDim matrix(1 To numcols, 1 To numcols)
    For i = 1 To numcols
        For j = 1 To numcols
            matrix(i, j) = Application.WorksheetFunction.Covar(rng.Columns(i), rng.Columns(j))
        Next j
    Next i
    VarCovar = matrix
End Function

Sub Variance_Matrix()
    Dim ws As Worksheet, lastRow As Long, formulaString As String
    Set ws = ThisWorkbook.Worksheets("幕後結果")
    
    ' Get the last non-empty row in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Copy stock codes from D1:M1 to AB8 (horizontal) and to AA9 (transposed)
    ws.Range("D1:M1").Copy Destination:=ws.Range("AB8")
    Application.CutCopyMode = False
    ws.Range("D1:M1").Copy
    ws.Range("AA9").PasteSpecial Paste:=xlPasteAll, Transpose:=True
    Application.CutCopyMode = False
    
    ' Set formula to calculate the variance-covariance matrix
    formulaString = "=VarCovar(P3:Y" & lastRow & ",AB10:AK22)"
    ws.Range("AB9").Formula2 = formulaString
    
    ' Add thick borders around the result range
    Dim rng As Range
    Set rng = ws.Range("AA8:AK19")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
End Sub

Function Cor_Matrix(rng As Range, target As Range) As Variant
    Dim i As Integer, j As Integer, numcols As Integer
    numcols = rng.Columns.Count
    Dim matrix() As Variant
    ReDim matrix(1 To numcols, 1 To numcols)
    For i = 1 To numcols
        For j = 1 To numcols
            matrix(i, j) = Application.WorksheetFunction.Correl(rng.Columns(i), rng.Columns(j))
        Next j
    Next i
    Cor_Matrix = matrix
End Function

Sub Correlation_Matrix()
    Dim ws As Worksheet, lastRow As Long, formulaString As String
    Set ws = ThisWorkbook.Worksheets("幕後結果")
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Copy stock codes from D1:M1 to AB22 (horizontal) and AA23 (transposed)
    ws.Range("D1:M1").Copy Destination:=ws.Range("AB22")
    Application.CutCopyMode = False
    ws.Range("D1:M1").Copy
    ws.Range("AA23").PasteSpecial Paste:=xlPasteAll, Transpose:=True
    Application.CutCopyMode = False
    
    ' Set formula to calculate the correlation matrix
    formulaString = "=Cor_Matrix(P3:Y" & lastRow & ",AB24:AK35)"
    ws.Range("AB23").Formula2 = formulaString
    
    Dim rng As Range
    Set rng = ws.Range("AA22:AK33")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
End Sub

Function CholeskyDecompose(matrix As Range) As Variant
    Dim A, LTM() As Double, S As Double
    Dim i As Long, j As Long, K As Long, N As Long, M As Long

    A = matrix
    N = matrix.Rows.Count
    M = matrix.Columns.Count
    If N <> M Then
        CholeskyDecompose = "Non-invertible matrix"
        Exit Function
    End If

    ReDim LTM(1 To N, 1 To N)
    For j = 1 To N
        S = 0
        For K = 1 To j - 1
            S = S + LTM(j, K) ^ 2
        Next K
        LTM(j, j) = A(j, j) - S
        If LTM(j, j) <= 0 Then Exit For
        LTM(j, j) = Sqr(LTM(j, j))
        For i = j + 1 To N
            S = 0
            For K = 1 To j - 1
                S = S + LTM(i, K) * LTM(j, K)
            Next K
            LTM(i, j) = (A(i, j) - S) / LTM(j, j)
        Next i
    Next j
    CholeskyDecompose = LTM
End Function

Sub CholeskyDecompose_Matrix()
    Dim ws As Worksheet, lastRow As Long, formulaString As String
    Set ws = ThisWorkbook.Worksheets("幕後結果")
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Copy stock codes from D1:M1 to AB36 (horizontal) and AA37 (transposed)
    ws.Range("D1:M1").Copy Destination:=ws.Range("AB36")
    Application.CutCopyMode = False
    ws.Range("D1:M1").Copy
    ws.Range("AA37").PasteSpecial Paste:=xlPasteAll, Transpose:=True
    Application.CutCopyMode = False
    
    ' Set formula to calculate the Cholesky decomposition
    formulaString = "=CholeskyDecompose(AB23#)"
    ws.Range("AB37").Formula2 = formulaString
    
    Dim rng As Range
    Set rng = ws.Range("AA36:AK47")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
End Sub

```

```vb
'===================================================
' Simulation and Optimization Module
'===================================================
Sub Simulation_and_Optimize_All()
    Random1
    Random_Normal
    EpsilonTable
    SimulateReturn
    SimulateReturn_Mean
    Random2
    Optimize_All
End Sub

'----------------------------
' Generate Random Array 1
'----------------------------
Sub Random1()
    Dim ws2 As Worksheet, ws3 As Worksheet
    Dim formulaString As String
    Dim rng1 As Range

    ' Set worksheets
    Set ws2 = ThisWorkbook.Worksheets("幕後結果")
    Set ws3 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Copy stock codes from D1:M1 from ws2 to ws3, cell B1
    ws2.Range("D1:M1").Copy Destination:=ws3.Range("B1")
    Application.CutCopyMode = False
    
    ' Clear previous random data
    ws3.Range("B2:K1001").ClearContents
    
    ' Insert random array formula in B2 with dimensions 1000 rows x 10 columns
    ws3.Range("B2").Formula2 = "=RANDARRAY(1000,10)"
    
    ' Add thick blue borders around the range A1:K1001
    Dim rng As Range
    Set rng = ws3.Range("A1:K1001")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    
    ' Convert formulas to static values
    Set rng1 = ws3.Range("B1003:K2002")
    rng1.Value = rng1.Value
End Sub

'----------------------------
' Generate Random Normal Array
'----------------------------
Sub Random_Normal()
    Dim ws2 As Worksheet, ws3 As Worksheet
    Dim formulaString As String
    Dim cell As Range

    Set ws2 = ThisWorkbook.Worksheets("幕後結果")
    Set ws3 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Clear previous data in N2:W1001
    ws3.Range("N2:W1001").ClearContents
    
    ' Copy stock codes from D1:M1 to N1 in ws3
    ws2.Range("D1:M1").Copy Destination:=ws3.Range("N1")
    Application.CutCopyMode = False
    
    ' Append " ~N(0,1)" to each cell in N1:W1 if not already appended
    For Each cell In ws3.Range("N1:W1")
        If Right(cell.Value, 6) <> " ~N(0,1)" Then
            cell.Value = cell.Value & " ~N(0,1)"
        End If
    Next cell
    
    Sheets("幕後結果_最佳化").Select
    Range("N2").Select
    
    ' Insert normal inverse function formula in N2
    formulaString = "=NORM.INV(B2,0,1)"
    ws3.Range("N2").Formula2 = formulaString
    
    ' Copy the formula rightwards and downwards
    ws3.Range("N2").Copy
    ws3.Range("O2:W2").PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    ws3.Range("N2:W2").Copy
    ws3.Range("N3:W1001").PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    
    ' Convert both random arrays to static values
    ws3.Range("B2:K1001").Copy
    ws3.Range("B2:K1001").PasteSpecial Paste:=xlPasteValues
    ws3.Range("N2:W1001").Copy
    ws3.Range("N2:W1001").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Add thick orange borders around the range M1:W1001
    Dim rng As Range
    Set rng = ws3.Range("M1:W1001")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(198, 89, 17)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(198, 89, 17)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(198, 89, 17)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(198, 89, 17)
    End With
End Sub

'----------------------------
' Build Epsilon Table (Linking Cholesky Matrix & Normal Random)
'----------------------------
Sub EpsilonTable()
    Dim ws2 As Worksheet, ws3 As Worksheet
    Dim formulaString As String
    Dim cell As Range

    Set ws2 = ThisWorkbook.Worksheets("幕後結果")
    Set ws3 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ws3.Range("N1003:W2002").ClearContents
    
    ' Copy stock codes from D1:M1 to N1002 (horizontal) in ws3
    ws2.Range("D1:M1").Copy Destination:=ws3.Range("N1002")
    Application.CutCopyMode = False
    
    ' Append "ε" to each cell in N1002:W1002 if not already appended
    For Each cell In ws3.Range("N1002:W1002")
        If Right(cell.Value, 1) <> "ε" Then
            cell.Value = cell.Value & "ε"
        End If
    Next cell
    
    ' Set formula to create epsilon values via a matrix multiplication (transposed)
    formulaString = "=(TRANSPOSE(MMULT(幕後結果!$AB$37:$AK$46, TRANSPOSE(N2:W2))))"
    ws3.Range("N1003").Formula2 = formulaString
    ws3.Range("N1003").AutoFill Destination:=ws3.Range("N1003:W2002"), Type:=xlFillDefault
    ws3.Range("N1003:W2002").FillDown
    
    ' Add thick yellow borders around the range M1002:W2002
    Dim rng As Range
    Set rng = ws3.Range("M1002:W2002")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(191, 143, 0)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(191, 143, 0)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(191, 143, 0)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(191, 143, 0)
    End With
End Sub

'----------------------------
' Simulate Annualized Return (3-Month Data)
'----------------------------
Sub SimulateReturn()
    Dim ws2 As Worksheet, ws3 As Worksheet
    Dim formulaString1 As String, formulaString2 As String, formulaString3 As String
    Dim formulaString4 As String, formulaString5 As String, formulaString6 As String
    Dim formulaString7 As String, formulaString8 As String, formulaString9 As String, formulaString10 As String
    
    Set ws2 = ThisWorkbook.Worksheets("幕後結果")
    Set ws3 = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Clear previous simulated return data
    ws3.Range("Z2:AI1001").ClearContents
    
    ' Copy stock codes from D1:M1 to Z1 in ws3
    ws2.Range("D1:M1").Copy Destination:=ws3.Range("Z1")
    Application.CutCopyMode = False
    
    ' Append "R" to each cell in Z1:AI1 if not already appended
    Dim cell As Range
    For Each cell In ws3.Range("Z1:AI1")
        If Right(cell.Value, 1) <> "R" Then
            cell.Value = cell.Value & "R"
        End If
    Next cell
    
    ' Define formulas for simulated return (annualized based on 3-month data)
    formulaString1 = "=(幕後結果!$AB$2+幕後結果!$AB$4*幕後結果_最佳化!$N1003)*12/3"
    formulaString2 = "=(幕後結果!$AC$2+幕後結果!$AC$4*幕後結果_最佳化!$O1003)*12/3"
    formulaString3 = "=(幕後結果!$AD$2+幕後結果!$AD$4*幕後結果_最佳化!$P1003)*12/3"
    formulaString4 = "=(幕後結果!$AE$2+幕後結果!$AE$4*幕後結果_最佳化!$Q1003)*12/3"
    formulaString5 = "=(幕後結果!$AF$2+幕後結果!$AF$4*幕後結果_最佳化!$R1003)*12/3"
    formulaString6 = "=(幕後結果!$AG$2+幕後結果!$AG$4*幕後結果_最佳化!$S1003)*12/3"
    formulaString7 = "=(幕後結果!$AH$2+幕後結果!$AH$4*幕後結果_最佳化!$T1003)*12/3"
    formulaString8 = "=(幕後結果!$AI$2+幕後結果!$AI$4*幕後結果_最佳化!$U1003)*12/3"
    formulaString9 = "=(幕後結果!$AJ$2+幕後結果!$AJ$4*幕後結果_最佳化!$V1003)*12/3"
    formulaString10 = "=(幕後結果!$AK$2+幕後結果!$AK$4*幕後結果_最佳化!$W1003)*12/3"
    
    ' Apply formulas to cells Z2:AI2 and autofill down to row 1001
    ws3.Range("Z2").Formula2 = formulaString1
    ws3.Range("AA2").Formula2 = formulaString2
    ws3.Range("AB2").Formula2 = formulaString3
    ws3.Range("AC2").Formula2 = formulaString4
    ws3.Range("AD2").Formula2 = formulaString5
    ws3.Range("AE2").Formula2 = formulaString6
    ws3.Range("AF2").Formula2 = formulaString7
    ws3.Range("AG2").Formula2 = formulaString8
    ws3.Range("AH2").Formula2 = formulaString9
    ws3.Range("AI2").Formula2 = formulaString10
    
    ws3.Range("Z2").AutoFill Destination:=ws3.Range("Z2:Z1001"), Type:=xlFillDefault
    ws3.Range("AA2").AutoFill Destination:=ws3.Range("AA2:AA1001"), Type:=xlFillDefault
    ws3.Range("AB2").AutoFill Destination:=ws3.Range("AB2:AB1001"), Type:=xlFillDefault
    ws3.Range("AC2").AutoFill Destination:=ws3.Range("AC2:AC1001"), Type:=xlFillDefault
    ws3.Range("AD2").AutoFill Destination:=ws3.Range("AD2:AD1001"), Type:=xlFillDefault
    ws3.Range("AE2").AutoFill Destination:=ws3.Range("AE2:AE1001"), Type:=xlFillDefault
    ws3.Range("AF2").AutoFill Destination:=ws3.Range("AF2:AF1001"), Type:=xlFillDefault
    ws3.Range("AG2").AutoFill Destination:=ws3.Range("AG2:AG1001"), Type:=xlFillDefault
    ws3.Range("AH2").AutoFill Destination:=ws3.Range("AH2:AH1001"), Type:=xlFillDefault
    ws3.Range("AI2").AutoFill Destination:=ws3.Range("AI2:AI1001"), Type:=xlFillDefault
    ws3.Range("Z2:AI2").Value = 0
    
    ' Add thick purple borders around the range Y1:AI1001
    Dim rng As Range
    Set rng = ws3.Range("Y1:AI1001")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
End Sub

Sub Random2()
    Dim ws As Worksheet
    Dim formulaString As String
    Dim rng1 As Range
    
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Copy stock codes from Z1:AI1 to B1002
    ws.Range("Z1:AI1").Copy Destination:=ws.Range("B1002")
    Application.CutCopyMode = False
    
    ' Set range B1003:K1003 to 0
    ws.Range("B1003:K1003").Value = 0
    ws.Range("B1004:K2002").ClearContents
    
    Sheets("幕後結果_最佳化").Select
    Range("B1004").Select
    
    formulaString = "=RANDARRAY(999,10)"
    ws.Range("B1004").Formula2 = formulaString
    
    Dim rng As Range
    Set rng = ws.Range("A1002:K2002")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    
    Set rng1 = ws.Range("B1003:K2002")
    rng1.Value = rng1.Value
End Sub

'----------------------------
' Optimize and Calculate Mean Return
'----------------------------
Sub SimulateReturn_Mean()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Copy stock codes from Z1:AI1 to Z1004
    ws.Range("Z1:AI1").Copy Destination:=ws.Range("Z1004")
    Application.CutCopyMode = False
    
    ' Calculate the average return for each stock (across rows 2 to 1001) and autofill from Z1005 to AI1005
    ws.Range("Z1005").Formula = "=AVERAGE(Z2:Z1001)"
    ws.Range("Z1005").AutoFill Destination:=ws.Range("Z1005:AI1005"), Type:=xlFillDefault
    
    ' Add thick purple borders around Y1004:AI1005
    Dim rng As Range
    Set rng = ws.Range("Y1004:AI1005")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(112, 48, 160)
    End With
End Sub


Sub Random2()  ' This macro generates a random array of numbers (0 to 1) independent of input values.
    Dim ws As Worksheet
    Dim formulaString As String
    Dim cell As Range
    Dim rng1 As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    
    ' Copy stock codes from Z1:AI1 to cell B1002
    ws.Range("Z1:AI1").Copy Destination:=ws.Range("B1002")
    Application.CutCopyMode = False
    
    ' Set the range B1003:K1003 to 0
    ws.Range("B1003:K1003").Value = 0
    
    ' Clear previous values in B1004:K2002
    ws.Range("B1004:K2002").ClearContents
    
    Sheets("幕後結果_最佳化").Select
    Range("B1004").Select
    
    ' Define the formula string (without the @ symbol)
    formulaString = "=RANDARRAY(999,10)"
    
    ' Set the formula in cell B1004 using .Formula2
    ws.Range("B1004").Formula2 = formulaString
    
    ' Add thick blue borders around the range A1002:K2002
    Dim rng As Range
    Set rng = ws.Range("A1002:K2002")
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(47, 117, 181)
    End With
    
    ' Define the range to convert formulas to values
    Set rng1 = ws.Range("B1003:K2002")
    rng1.Value = rng1.Value
End Sub

```
```vb
'=== Main Command Button Handler ===
Private Sub CommandButton1_Click()
    ' Get the name of the currently active worksheet as the personality type
    Dim personalityType As String
    personalityType = ActiveSheet.Name
    
    ' Call the corresponding sub based on the personality type
    Select Case personalityType
        Case "風險厭惡"
            Call 雪天型
        Case "風險中立偏厭惡"
            Call 陰天型
        Case "風險中立"
            Call 晴天型
        Case "風險中立偏愛好"
            Call 雷雨型
        Case "風險愛好"
            Call 閃電型
        Case Else
            MsgBox "Unrecognized type: " & personalityType
    End Select
End Sub

'=== Sub for "雪天型" (Snow Type) ===
Sub 雪天型()
    Dim wsResult As Worksheet
    Dim wsLowRisk As Worksheet
    Dim wsMediumRisk As Worksheet
    Dim wsHighRisk As Worksheet
    Dim showWs As Worksheet
    Dim celebrity As String
    Dim choiceCol As Integer
    
    ' Get celebrity selection from the user form
    If 選擇名人.OptionButton1.Value Then
        celebrity = "Buffet"
    ElseIf 選擇名人.OptionButton2.Value Then
        celebrity = "Graham"
    ElseIf 選擇名人.OptionButton3.Value Then
        celebrity = "O'Shaughnessy"
    ElseIf 選擇名人.OptionButton4.Value Then
        celebrity = "Murphy Score"
    ElseIf 選擇名人.OptionButton5.Value Then
        Call 雪天型隨機
        On Error Resume Next
        Set wsResult = ThisWorkbook.Sheets("雪天型結果")
        On Error GoTo 0
        If wsResult Is Nothing Then
            Set wsResult = ThisWorkbook.Sheets.Add
            wsResult.Name = "雪天型結果"
        End If
        wsResult.Range("A2:C8").ClearContents
        ThisWorkbook.Sheets("雪天型隨機").Range("A2:C8").Copy Destination:=wsResult.Range("A2")
        wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
        wsResult.Range("A2:A7").Interior.Color = RGB(221, 235, 247)
        wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
        wsResult.Range("B2:B4").Interior.Color = RGB(255, 242, 204)
        wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
        wsResult.Range("C2").Interior.Color = RGB(252, 228, 214)
        Set showWs = ThisWorkbook.Sheets("雪天型結果")
        Unload 名人介紹
        showWs.Activate
        Exit Sub
    Else
        MsgBox "Please select a valid celebrity on the user form."
        Exit Sub
    End If

    ' Determine the score column based on the celebrity
    Select Case celebrity
        Case "Buffet": choiceCol = 19
        Case "Graham": choiceCol = 20
        Case "O'Shaughnessy": choiceCol = 21
        Case "Murphy Score": choiceCol = 22
    End Select

    Set wsLowRisk = ThisWorkbook.Sheets("low_risk")
    Set wsMediumRisk = ThisWorkbook.Sheets("medium_risk")
    Set wsHighRisk = ThisWorkbook.Sheets("high_risk")
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("雪天型結果")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "雪天型結果"
    Else
        wsResult.Cells.Clear
    End If
    
    wsResult.Cells(1, 1).Value = "Low Risk - Top 6"
    wsResult.Cells(1, 2).Value = "Medium Risk - Top 3"
    wsResult.Cells(1, 3).Value = "High Risk - Top 1"
    
    ShowTopNResults wsLowRisk, wsResult, choiceCol, 1, 6
    ShowTopNResults wsMediumRisk, wsResult, choiceCol, 2, 3
    ShowTopNResults wsHighRisk, wsResult, choiceCol, 3, 1
    
    wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
    wsResult.Range("A2:A7").Interior.Color = RGB(221, 235, 247)
    wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
    wsResult.Range("B2:B4").Interior.Color = RGB(255, 242, 204)
    wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
    wsResult.Range("C2").Interior.Color = RGB(252, 228, 214)
    
    Set showWs = ThisWorkbook.Sheets("雪天型結果")
    Unload 名人介紹
    showWs.Activate
End Sub

'=== Sub for "陰天型" (Cloudy Type) ===
Sub 陰天型()
    Dim wsResult As Worksheet
    Dim wsLowRisk As Worksheet
    Dim wsMediumRisk As Worksheet
    Dim wsHighRisk As Worksheet
    Dim showWs As Worksheet
    Dim celebrity As String
    Dim choiceCol As Integer
    
    If 選擇名人.OptionButton1.Value Then
        celebrity = "Buffet"
    ElseIf 選擇名人.OptionButton2.Value Then
        celebrity = "Graham"
    ElseIf 選擇名人.OptionButton3.Value Then
        celebrity = "O'Shaughnessy"
    ElseIf 選擇名人.OptionButton4.Value Then
        celebrity = "Murphy Score"
    ElseIf 選擇名人.OptionButton5.Value Then
        Call 陰天型隨機
        On Error Resume Next
        Set wsResult = ThisWorkbook.Sheets("陰天型結果")
        On Error GoTo 0
        If wsResult Is Nothing Then
            Set wsResult = ThisWorkbook.Sheets.Add
            wsResult.Name = "陰天型結果"
        End If
        wsResult.Range("A2:C8").ClearContents
        ThisWorkbook.Sheets("陰天型隨機").Range("A2:C8").Copy Destination:=wsResult.Range("A2")
        wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
        wsResult.Range("A2:A6").Interior.Color = RGB(221, 235, 247)
        wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
        wsResult.Range("B2:B5").Interior.Color = RGB(255, 242, 204)
        wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
        wsResult.Range("C2").Interior.Color = RGB(252, 228, 214)
        Set showWs = ThisWorkbook.Sheets("陰天型結果")
        Unload 名人介紹
        showWs.Activate
        Exit Sub
    Else
        MsgBox "Please select a valid celebrity on the user form."
        Exit Sub
    End If

    Select Case celebrity
        Case "Buffet": choiceCol = 19
        Case "Graham": choiceCol = 20
        Case "O'Shaughnessy": choiceCol = 21
        Case "Murphy Score": choiceCol = 22
    End Select

    Set wsLowRisk = ThisWorkbook.Sheets("low_risk")
    Set wsMediumRisk = ThisWorkbook.Sheets("medium_risk")
    Set wsHighRisk = ThisWorkbook.Sheets("high_risk")
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("陰天型結果")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "陰天型結果"
    Else
        wsResult.Cells.Clear
    End If
    
    wsResult.Cells(1, 1).Value = "Low Risk - Top 5"
    wsResult.Cells(1, 2).Value = "Medium Risk - Top 4"
    wsResult.Cells(1, 3).Value = "High Risk - Top 1"
    
    ShowTopNResults wsLowRisk, wsResult, choiceCol, 1, 5
    ShowTopNResults wsMediumRisk, wsResult, choiceCol, 2, 4
    ShowTopNResults wsHighRisk, wsResult, choiceCol, 3, 1
    
    wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
    wsResult.Range("A2:A6").Interior.Color = RGB(221, 235, 247)
    wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
    wsResult.Range("B2:B5").Interior.Color = RGB(255, 242, 204)
    wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
    wsResult.Range("C2").Interior.Color = RGB(252, 228, 214)
    
    Set showWs = ThisWorkbook.Sheets("陰天型結果")
    Unload 名人介紹
    showWs.Activate
End Sub

'=== Sub for "晴天型" (Sunny Type) ===
Sub 晴天型()
    Dim wsResult As Worksheet
    Dim wsLowRisk As Worksheet
    Dim wsMediumRisk As Worksheet
    Dim wsHighRisk As Worksheet
    Dim showWs As Worksheet
    Dim celebrity As String
    Dim choiceCol As Integer
    
    If 選擇名人.OptionButton1.Value Then
        celebrity = "Buffet"
    ElseIf 選擇名人.OptionButton2.Value Then
        celebrity = "Graham"
    ElseIf 選擇名人.OptionButton3.Value Then
        celebrity = "O'Shaughnessy"
    ElseIf 選擇名人.OptionButton4.Value Then
        celebrity = "Murphy Score"
    ElseIf 選擇名人.OptionButton5.Value Then
        Call 晴天型隨機
        On Error Resume Next
        Set wsResult = ThisWorkbook.Sheets("晴天型結果")
        On Error GoTo 0
        If wsResult Is Nothing Then
            Set wsResult = ThisWorkbook.Sheets.Add
            wsResult.Name = "晴天型結果"
        End If
        wsResult.Range("A2:C8").ClearContents
        ThisWorkbook.Sheets("晴天型隨機").Range("A2:C8").Copy Destination:=wsResult.Range("A2")
        wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
        wsResult.Range("A2:A5").Interior.Color = RGB(221, 235, 247)
        wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
        wsResult.Range("B2:B5").Interior.Color = RGB(255, 242, 204)
        wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
        wsResult.Range("C2:C3").Interior.Color = RGB(252, 228, 214)
        Set showWs = ThisWorkbook.Sheets("晴天型結果")
        Unload 名人介紹
        showWs.Activate
        Exit Sub
    Else
        MsgBox "Please select a valid celebrity on the user form."
        Exit Sub
    End If

    Select Case celebrity
        Case "Buffet": choiceCol = 19
        Case "Graham": choiceCol = 20
        Case "O'Shaughnessy": choiceCol = 21
        Case "Murphy Score": choiceCol = 22
    End Select

    Set wsLowRisk = ThisWorkbook.Sheets("low_risk")
    Set wsMediumRisk = ThisWorkbook.Sheets("medium_risk")
    Set wsHighRisk = ThisWorkbook.Sheets("high_risk")
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("晴天型結果")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "晴天型結果"
    Else
        wsResult.Cells.Clear
    End If
    
    wsResult.Cells(1, 1).Value = "Low Risk - Top 4"
    wsResult.Cells(1, 2).Value = "Medium Risk - Top 4"
    wsResult.Cells(1, 3).Value = "High Risk - Top 2"
    
    ShowTopNResults wsLowRisk, wsResult, choiceCol, 1, 4
    ShowTopNResults wsMediumRisk, wsResult, choiceCol, 2, 4
    ShowTopNResults wsHighRisk, wsResult, choiceCol, 3, 2
    
    wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
    wsResult.Range("A2:A5").Interior.Color = RGB(221, 235, 247)
    wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
    wsResult.Range("B2:B5").Interior.Color = RGB(255, 242, 204)
    wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
    wsResult.Range("C2:C3").Interior.Color = RGB(252, 228, 214)
    
    Set showWs = ThisWorkbook.Sheets("晴天型結果")
    Unload 名人介紹
    showWs.Activate
End Sub

'=== Sub for "雷雨型" (Thunderstorm Type) ===
Sub 雷雨型()
    Dim wsResult As Worksheet
    Dim wsLowRisk As Worksheet
    Dim wsMediumRisk As Worksheet
    Dim wsHighRisk As Worksheet
    Dim showWs As Worksheet
    Dim celebrity As String
    Dim choiceCol As Integer
    
    If 選擇名人.OptionButton1.Value Then
        celebrity = "Buffet"
    ElseIf 選擇名人.OptionButton2.Value Then
        celebrity = "Graham"
    ElseIf 選擇名人.OptionButton3.Value Then
        celebrity = "O'Shaughnessy"
    ElseIf 選擇名人.OptionButton4.Value Then
        celebrity = "Murphy Score"
    ElseIf 選擇名人.OptionButton5.Value Then
        Call 雷雨型隨機
        On Error Resume Next
        Set wsResult = ThisWorkbook.Sheets("雷雨型結果")
        On Error GoTo 0
        If wsResult Is Nothing Then
            Set wsResult = ThisWorkbook.Sheets.Add
            wsResult.Name = "雷雨型結果"
        End If
        wsResult.Range("A2:C8").ClearContents
        ThisWorkbook.Sheets("雷雨型隨機").Range("A2:C8").Copy Destination:=wsResult.Range("A2")
        wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
        wsResult.Range("A2:A4").Interior.Color = RGB(221, 235, 247)
        wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
        wsResult.Range("B2:B5").Interior.Color = RGB(255, 242, 204)
        wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
        wsResult.Range("C2:C4").Interior.Color = RGB(252, 228, 214)
        Set showWs = ThisWorkbook.Sheets("雷雨型結果")
        Unload 名人介紹
        showWs.Activate
        Exit Sub
    Else
        MsgBox "Please select a valid celebrity on the user form."
        Exit Sub
    End If

    Select Case celebrity
        Case "Buffet": choiceCol = 19
        Case "Graham": choiceCol = 20
        Case "O'Shaughnessy": choiceCol = 21
        Case "Murphy Score": choiceCol = 22
    End Select

    Set wsLowRisk = ThisWorkbook.Sheets("low_risk")
    Set wsMediumRisk = ThisWorkbook.Sheets("medium_risk")
    Set wsHighRisk = ThisWorkbook.Sheets("high_risk")
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("雷雨型結果")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "雷雨型結果"
    Else
        wsResult.Cells.Clear
    End If
    
    wsResult.Cells(1, 1).Value = "Low Risk - Top 3"
    wsResult.Cells(1, 2).Value = "Medium Risk - Top 4"
    wsResult.Cells(1, 3).Value = "High Risk - Top 3"
    
    ShowTopNResults wsLowRisk, wsResult, choiceCol, 1, 3
    ShowTopNResults wsMediumRisk, wsResult, choiceCol, 2, 4
    ShowTopNResults wsHighRisk, wsResult, choiceCol, 3, 3
    
    wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
    wsResult.Range("A2:A4").Interior.Color = RGB(221, 235, 247)
    wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
    wsResult.Range("B2:B5").Interior.Color = RGB(255, 242, 204)
    wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
    wsResult.Range("C2:C4").Interior.Color = RGB(252, 228, 214)
    
    Set showWs = ThisWorkbook.Sheets("雷雨型結果")
    Unload 名人介紹
    showWs.Activate
End Sub

'=== Sub for "閃電型" (Lightning Type) ===
Sub 閃電型()
    Dim wsResult As Worksheet
    Dim wsLowRisk As Worksheet
    Dim wsMediumRisk As Worksheet
    Dim wsHighRisk As Worksheet
    Dim showWs As Worksheet
    Dim celebrity As String
    Dim choiceCol As Integer
    
    If 選擇名人.OptionButton1.Value Then
        celebrity = "Buffet"
    ElseIf 選擇名人.OptionButton2.Value Then
        celebrity = "Graham"
    ElseIf 選擇名人.OptionButton3.Value Then
        celebrity = "O'Shaughnessy"
    ElseIf 選擇名人.OptionButton4.Value Then
        celebrity = "Murphy Score"
    ElseIf 選擇名人.OptionButton5.Value Then
        Call 閃電型隨機
        On Error Resume Next
        Set wsResult = ThisWorkbook.Sheets("閃電型結果")
        On Error GoTo 0
        If wsResult Is Nothing Then
            Set wsResult = ThisWorkbook.Sheets.Add
            wsResult.Name = "閃電型結果"
        End If
        wsResult.Range("A2:C8").ClearContents
        ThisWorkbook.Sheets("閃電型隨機").Range("A2:C8").Copy Destination:=wsResult.Range("A2")
        wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
        wsResult.Range("A2:A3").Interior.Color = RGB(221, 235, 247)
        wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
        wsResult.Range("B2:B4").Interior.Color = RGB(255, 242, 204)
        wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
        wsResult.Range("C2:C6").Interior.Color = RGB(252, 228, 214)
        Set showWs = ThisWorkbook.Sheets("閃電型結果")
        Unload 名人介紹
        showWs.Activate
        Exit Sub
    Else
        MsgBox "Please select a valid celebrity on the user form."
        Exit Sub
    End If

    Select Case celebrity
        Case "Buffet": choiceCol = 19
        Case "Graham": choiceCol = 20
        Case "O'Shaughnessy": choiceCol = 21
        Case "Murphy Score": choiceCol = 22
    End Select

    Set wsLowRisk = ThisWorkbook.Sheets("low_risk")
    Set wsMediumRisk = ThisWorkbook.Sheets("medium_risk")
    Set wsHighRisk = ThisWorkbook.Sheets("high_risk")
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("閃電型結果")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "閃電型結果"
    Else
        wsResult.Cells.Clear
    End If
    
    wsResult.Cells(1, 1).Value = "Low Risk - Top 2"
    wsResult.Cells(1, 2).Value = "Medium Risk - Top 3"
    wsResult.Cells(1, 3).Value = "High Risk - Top 5"
    
    ShowTopNResults wsLowRisk, wsResult, choiceCol, 1, 2
    ShowTopNResults wsMediumRisk, wsResult, choiceCol, 2, 3
    ShowTopNResults wsHighRisk, wsResult, choiceCol, 3, 5
    
    wsResult.Range("A1").Interior.Color = RGB(189, 215, 238)
    wsResult.Range("A2:A3").Interior.Color = RGB(221, 235, 247)
    wsResult.Range("B1").Interior.Color = RGB(255, 217, 102)
    wsResult.Range("B2:B4").Interior.Color = RGB(255, 242, 204)
    wsResult.Range("C1").Interior.Color = RGB(244, 176, 132)
    wsResult.Range("C2:C6").Interior.Color = RGB(252, 228, 214)
    
    Set showWs = ThisWorkbook.Sheets("閃電型結果")
    Unload 名人介紹
    showWs.Activate
End Sub

'=== Helper Sub to display top N results ===
Sub ShowTopNResults(wsSource As Worksheet, wsDestination As Worksheet, ScoreColumn As Integer, DestCol As Integer, TopN As Integer)
    Dim rng As Range, cell As Range
    Dim i As Integer, LastRowSrc As Long, LastRowDest As Long
    Dim TopScores As Variant
    Dim StockDict As Object
    Set StockDict = CreateObject("Scripting.Dictionary")
    
    LastRowSrc = wsSource.Cells(wsSource.Rows.Count, ScoreColumn).End(xlUp).Row
    LastRowDest = wsDestination.Cells(wsDestination.Rows.Count, DestCol).End(xlUp).Row + 1
    
    Dim StockScores As Object
    Set StockScores = CreateObject("Scripting.Dictionary")
    
    For Each cell In wsSource.Range(wsSource.Cells(2, ScoreColumn), wsSource.Cells(LastRowSrc, ScoreColumn))
        StockScores.Add cell.Row, cell.Value
    Next cell
    
    TopScores = SortAndTop(StockScores, TopN)
    
    For i = 1 To UBound(TopScores, 1)
        If Not StockDict.exists(wsSource.Cells(TopScores(i, 1), 1).Value) Then
            wsDestination.Cells(LastRowDest, DestCol).Value = wsSource.Cells(TopScores(i, 1), 1).Value
            StockDict.Add wsSource.Cells(TopScores(i, 1), 1).Value, Nothing
            LastRowDest = LastRowDest + 1
        End If
    Next i
End Sub

Function SortAndTop(dict As Object, N As Integer) As Variant
    Dim dictArray() As Variant, sortedArray() As Variant, i As Integer
    ReDim dictArray(1 To dict.Count, 1 To 2)
    i = 1
    Dim key
    For Each key In dict.Keys
        dictArray(i, 1) = key
        dictArray(i, 2) = dict(key)
        i = i + 1
    Next key
    Call QuickSort(dictArray, LBound(dictArray, 1), UBound(dictArray, 1), 2)
    If N > UBound(dictArray, 1) Then N = UBound(dictArray, 1)
    ReDim sortedArray(1 To N, 1 To 2)
    For i = 1 To N
        sortedArray(i, 1) = dictArray(i, 1)
        sortedArray(i, 2) = dictArray(i, 2)
    Next i
    SortAndTop = sortedArray
End Function

Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long, SortColumn As Integer)
    Dim pivot As Variant, tmpSwap1 As Variant, tmpSwap2 As Variant
    Dim tmpLow As Long, tmpHi As Long
    tmpLow = inLow
    tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2, SortColumn)
    
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow, SortColumn) > pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (pivot > vArray(tmpHi, SortColumn) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap1 = vArray(tmpLow, 1)
            tmpSwap2 = vArray(tmpLow, 2)
            vArray(tmpLow, 1) = vArray(tmpHi, 1)
            vArray(tmpLow, 2) = vArray(tmpHi, 2)
            vArray(tmpHi, 1) = tmpSwap1
            vArray(tmpHi, 2) = tmpSwap2
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi, SortColumn
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi, SortColumn
End Sub

Private Sub CommandButton2_Click()
    SelectedCelebrity = ""
    Celebrityintro = ""
    名人介紹.Hide
    選擇名人.Show
End Sub

Sub 雪天型隨機()
    Dim lowRiskRng As Range, mediumRiskRng As Range, highRiskRng As Range
    Dim newSheet As Worksheet, i As Integer, randomIndex As Long, showws As Worksheet
    
    Set lowRiskRng = ThisWorkbook.Sheets("low_risk").Range("A2:A256")
    Set mediumRiskRng = ThisWorkbook.Sheets("medium_risk").Range("A2:A135")
    Set highRiskRng = ThisWorkbook.Sheets("high_risk").Range("A2:A50")
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("雪天型隨機")
    On Error GoTo 0
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "雪天型隨機"
    Else
        newSheet.Cells.Clear
    End If
    
    Randomize
    newSheet.Cells(1, 1).Value = "low_risk Random Selection"
    For i = 1 To 6
        randomIndex = Int(lowRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 1).Value = lowRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 2).Value = "medium_risk Random Selection"
    For i = 1 To 3
        randomIndex = Int(mediumRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 2).Value = mediumRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 3).Value = "high_risk Random Selection"
    randomIndex = Int(highRiskRng.Cells.Count * Rnd) + 1
    newSheet.Cells(2, 3).Value = highRiskRng.Cells(randomIndex).Value
    
    newSheet.Columns("A:C").AutoFit
End Sub

Sub 陰天型隨機()
    Dim lowRiskRng As Range, mediumRiskRng As Range, highRiskRng As Range
    Dim newSheet As Worksheet, i As Integer, randomIndex As Long, showws As Worksheet
    
    Set lowRiskRng = ThisWorkbook.Sheets("low_risk").Range("A2:A272")
    Set mediumRiskRng = ThisWorkbook.Sheets("medium_risk").Range("A2:A135")
    Set highRiskRng = ThisWorkbook.Sheets("high_risk").Range("A2:A50")
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("陰天型隨機")
    On Error GoTo 0
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "陰天型隨機"
    Else
        newSheet.Cells.Clear
    End If
    
    Randomize
    newSheet.Cells(1, 1).Value = "low_risk Random Selection"
    For i = 1 To 5
        randomIndex = Int(lowRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 1).Value = lowRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 2).Value = "medium_risk Random Selection"
    For i = 1 To 4
        randomIndex = Int(mediumRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 2).Value = mediumRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 3).Value = "high_risk Random Selection"
    randomIndex = Int(highRiskRng.Cells.Count * Rnd) + 1
    newSheet.Cells(2, 3).Value = highRiskRng.Cells(randomIndex).Value
    
    newSheet.Columns("A:C").AutoFit
End Sub

Sub 晴天型隨機()
    Dim lowRiskRng As Range, mediumRiskRng As Range, highRiskRng As Range
    Dim newSheet As Worksheet, i As Integer, randomIndex As Long, showws As Worksheet
    
    Set lowRiskRng = ThisWorkbook.Sheets("low_risk").Range("A2:A272")
    Set mediumRiskRng = ThisWorkbook.Sheets("medium_risk").Range("A2:A135")
    Set highRiskRng = ThisWorkbook.Sheets("high_risk").Range("A2:A50")
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("晴天型隨機")
    On Error GoTo 0
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "晴天型隨機"
    Else
        newSheet.Cells.Clear
    End If
    
    Randomize
    newSheet.Cells(1, 1).Value = "low_risk Random Selection"
    For i = 1 To 4
        randomIndex = Int(lowRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 1).Value = lowRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 2).Value = "medium_risk Random Selection"
    For i = 1 To 4
        randomIndex = Int(mediumRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 2).Value = mediumRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 3).Value = "high_risk Random Selection"
    For i = 1 To 2
        randomIndex = Int(highRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 3).Value = highRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Columns("A:C").AutoFit
End Sub

Sub 雷雨型隨機()
    Dim lowRiskRng As Range, mediumRiskRng As Range, highRiskRng As Range
    Dim newSheet As Worksheet, i As Integer, randomIndex As Long, showws As Worksheet
    
    Set lowRiskRng = ThisWorkbook.Sheets("low_risk").Range("A2:A272")
    Set mediumRiskRng = ThisWorkbook.Sheets("medium_risk").Range("A2:A135")
    Set highRiskRng = ThisWorkbook.Sheets("high_risk").Range("A2:A50")
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("雷雨型隨機")
    On Error GoTo 0
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "雷雨型隨機"
    Else
        newSheet.Cells.Clear
    End If
    
    Randomize
    newSheet.Cells(1, 1).Value = "low_risk Random Selection"
    For i = 1 To 3
        randomIndex = Int(lowRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 1).Value = lowRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 2).Value = "medium_risk Random Selection"
    For i = 1 To 4
        randomIndex = Int(mediumRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 2).Value = mediumRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 3).Value = "high_risk Random Selection"
    For i = 1 To 3
        randomIndex = Int(highRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 3).Value = highRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Columns("A:C").AutoFit
End Sub

Sub 閃電型隨機()
    Dim lowRiskRng As Range, mediumRiskRng As Range, highRiskRng As Range
    Dim newSheet As Worksheet, i As Integer, randomIndex As Long, showws As Worksheet
    
    Set lowRiskRng = ThisWorkbook.Sheets("low_risk").Range("A2:A272")
    Set mediumRiskRng = ThisWorkbook.Sheets("medium_risk").Range("A2:A135")
    Set highRiskRng = ThisWorkbook.Sheets("high_risk").Range("A2:A50")
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("閃電型隨機")
    On Error GoTo 0
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "閃電型隨機"
    Else
        newSheet.Cells.Clear
    End If
    
    Randomize
    newSheet.Cells(1, 1).Value = "low_risk Random Selection"
    For i = 1 To 2
        randomIndex = Int(lowRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 1).Value = lowRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 2).Value = "medium_risk Random Selection"
    For i = 1 To 3
        randomIndex = Int(mediumRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 2).Value = mediumRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Cells(1, 3).Value = "high_risk Random Selection"
    For i = 1 To 5
        randomIndex = Int(highRiskRng.Cells.Count * Rnd) + 1
        newSheet.Cells(i + 1, 3).Value = highRiskRng.Cells(randomIndex).Value
    Next i
    
    newSheet.Columns("A:C").AutoFit
End Sub

Private Sub UserForm_Click()
    ' No action needed
End Sub

'=== UserForm Code ===
Dim SelectedCelebrity As String
Dim Celebrityintro As String
Dim celetype As Integer

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    ' No action needed
End Sub

Private Sub TextBox1_Change()
    ' No action needed
End Sub

Private Sub UserForm_Initialize()
    Dim basePath As String
    basePath = ThisWorkbook.Path & "\img\"
    
    Image1.Picture = LoadPicture(basePath & "巴菲特.jpg")
    Image2.Picture = LoadPicture(basePath & "葛拉漢.jpg")
    Image3.Picture = LoadPicture(basePath & "LBJ.jpg")
    Image4.Picture = LoadPicture(basePath & "麥可.jpg")
    Image5.Picture = LoadPicture(basePath & "隨機.jpg")
End Sub

Private Sub OptionButton1_Click()
    ' Set selected celebrity name from the "名人介紹" sheet
    SelectedCelebrity = Worksheets("名人介紹").Range("A2").Value
    Celebrityintro = Worksheets("名人介紹").Range("B2").Value
    celetype = 1
End Sub

Private Sub OptionButton2_Click()
    SelectedCelebrity = Worksheets("名人介紹").Range("A3").Value
    Celebrityintro = Worksheets("名人介紹").Range("B3").Value
    celetype = 2
End Sub

Private Sub OptionButton3_Click()
    SelectedCelebrity = Worksheets("名人介紹").Range("A4").Value
    Celebrityintro = Worksheets("名人介紹").Range("B4").Value
    celetype = 3
End Sub

Private Sub OptionButton4_Click()
    SelectedCelebrity = Worksheets("名人介紹").Range("A5").Value
    Celebrityintro = Worksheets("名人介紹").Range("B5").Value
    celetype = 4
End Sub

Private Sub OptionButton5_Click()
    SelectedCelebrity = Worksheets("名人介紹").Range("A6").Value
    Celebrityintro = Worksheets("名人介紹").Range("B6").Value
    celetype = 5
End Sub

Private Sub CommandButton1_Click()
    ' Hide the celebrity selection form and show the celebrity introduction form
    選擇名人.Hide
    名人介紹.Label1.Caption = SelectedCelebrity
    名人介紹.Label2.Caption = Celebrityintro
    名人介紹.Show
End Sub

```