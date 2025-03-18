# 股的Morning Project – VBA Code Repository

This repository contains the complete VBA code for the 股的Morning Project. Instead of sharing the Excel workbook, all the code has been saved as plain text here for easy review. The code is organized into several sections based on functionality.

---

## Table of Contents

1. [Worksheet Event Code](#worksheet-event-code)
2. [User Interaction Code](#user-interaction-code)
3. [Questionnaire & Navigation Code](#questionnaire--navigation-code)
4. [Process & External Script Integration](#process--external-script-integration)
5. [Simulation and Optimization Code](#simulation-and-optimization-code)
6. [PDF Export & Chart Generation Code](#pdf-export--chart-generation-code)
7. [Additional Navigation Code](#additional-navigation-code)

---

## Worksheet Event Code

This code runs when a specific worksheet is activated. It clears charts, formats cells, retrieves user data, and updates shapes accordingly.

```vb
Private Sub Worksheet_Activate()
    Dim name As Shape
    Dim age As Shape
    Dim gender As Shape
    Dim mail As Shape
    Dim portfolio As Shape
    Dim Monte_Carlo As Shape
    Dim weightData As Shape
    Dim information As Worksheet
    Dim portfoliosheet As Worksheet
    Dim optimizationSheet As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim portfolioData As String
    Dim label As String
    Dim cell As Range
    Dim rng As Range
    Dim i As Long, j As Long
    Dim decimalPlaces As Integer
    Dim chartObj As ChartObject

    ' Set desired decimal places
    decimalPlaces = 3
    
    Dim mean As Shape
    Dim var As Shape
    Dim Std As Shape
    Dim Sharpe As Shape
    
    Set ws = ThisWorkbook.Worksheets("風險厭惡結果單")
    
    ' Clear all charts on the worksheet
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Set shapes for key metrics
    Set mean = ws.Shapes("平均數")
    Set var = ws.Shapes("變異數")
    Set Std = ws.Shapes("標準差")
    Set Sharpe = ws.Shapes("夏普率")
    
    ' Reset shape text
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
    
    ' Process range H33:Q33 for numeric formatting
    Set rng = ws.Range("H33:Q33")
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            cell.Value = Format(cell.Value, "0." & String(decimalPlaces, "0"))
        End If
    Next cell
    
    ' Retrieve and set user data from the last row of "使用者資料"
    lastRow = information.Cells(information.Rows.Count, "A").End(xlUp).Row
    name.TextFrame.Characters.Text = information.Cells(lastRow, 1).Value
    age.TextFrame.Characters.Text = information.Cells(lastRow, 2).Value
    gender.TextFrame.Characters.Text = information.Cells(lastRow, 3).Value
    mail.TextFrame.Characters.Text = information.Cells(lastRow, 4).Value
    
    ' Gather portfolio data from "雪天型結果"
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
    
    ' Set "蒙地卡羅" shape text from optimization sheet range AN1:AW1
    label = ""
    For Each cell In optimizationSheet.Range("AN1:AW1")
        label = label & cell.Value & "     "
    Next cell
    label = Trim(label)
    Monte_Carlo.TextFrame2.TextRange.Text = Trim(label)
    With Monte_Carlo.TextFrame2
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' Set "權重" shape text from range H33:Q33
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
---

## User Interaction Code

Handles extracting text from shapes and ensuring unique user nicknames.

```vb
Sub ExtractTextFromShape()
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape
    Dim txtRng As Range
    Dim usersheet As Worksheet
    Dim nextRow As Long
    Dim found As Boolean
    Dim shp1Text As String
    Dim possheet As Worksheet

    Set usersheet = ThisWorkbook.Sheets("使用者資料")
    Set shp1 = ActiveSheet.Shapes("文字方塊 1")
    Set shp2 = ActiveSheet.Shapes("文字方塊 2")
    Set shp3 = ActiveSheet.Shapes("文字方塊 3")
    Set shp4 = ActiveSheet.Shapes("文字方塊 4")
    Set possheet = ThisWorkbook.Sheets("POS機按鈕")

    shp1Text = Trim(shp1.TextFrame.Characters.Text)

    found = False
    For Each txtRng In usersheet.Range("A:A").SpecialCells(xlCellTypeConstants)
        If Trim(txtRng.Value) = shp1Text Then
            found = True
            Exit For
        End If
    Next txtRng

    If found Then
        MsgBox "此暱稱已存在，請輸入其他暱稱。", vbExclamation, "重複的暱稱"
    Else
        nextRow = usersheet.Cells(usersheet.Rows.Count, "A").End(xlUp).Row + 1
        Set txtRng = usersheet.Range("A" & nextRow)
        txtRng.Value = shp1Text
        txtRng.Offset(0, 1).Value = shp2.TextFrame.Characters.Text
        txtRng.Offset(0, 2).Value = shp3.TextFrame.Characters.Text
        txtRng.Offset(0, 3).Value = shp4.TextFrame.Characters.Text

        shp1.TextFrame.Characters.Text = ""
        shp2.TextFrame.Characters.Text = ""
        shp3.TextFrame.Characters.Text = ""
        shp4.TextFrame.Characters.Text = ""
    End If
    possheet.Activate
End Sub
```
---

## Questionnaire & Navigation Code

This section manages the quiz interface for determining user financial personality.
```vb
Dim questions As Worksheet
Dim 正在答題 As Boolean
Dim N As Long
Dim textblock As Shape, Ablock As Shape, Bblock As Shape, Cblock As Shape, Dblock As Shape
Dim selectedCategory As String

Sub 開始答題(按鈕名稱 As String)
    Dim 目標工作表 As Worksheet
    Set 目標工作表 = ThisWorkbook.Worksheets("POS機")
    
    目標工作表.Activate
    selectedCategory = 按鈕名稱
    N = 2
    
    初始化文字方塊
    正在答題 = True
    出題
End Sub

Sub 出題()
    Set questions = ThisWorkbook.Worksheets("題庫")
    Dim maxRow As Long, totalpoints As Integer
    maxRow = questions.Cells(Rows.Count, 2).End(xlUp).Row
    
    Select Case selectedCategory
        Case "退休金": totalpoints = questions.Cells(2, 8).Value
        Case "第一桶金": totalpoints = questions.Cells(7, 8).Value
        Case "財產保值": totalpoints = questions.Cells(12, 8).Value
        Case "教育存款": totalpoints = questions.Cells(17, 8).Value
        Case "買房": totalpoints = questions.Cells(22, 8).Value
        Case "長期財富累積": totalpoints = questions.Cells(27, 8).Value
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
        Select Case totalpoints
            Case 4 To 7: ThisWorkbook.Worksheets("風險厭惡").Activate
            Case 8 To 10: ThisWorkbook.Worksheets("風險中立偏厭惡").Activate
            Case 11 To 14: ThisWorkbook.Worksheets("風險中立").Activate
            Case 15 To 17: ThisWorkbook.Worksheets("風險中立偏愛好").Activate
            Case 18 To 20: ThisWorkbook.Worksheets("風險愛好").Activate
        End Select
        清空題目與選項
        正在答題 = False
        N = 2
    End If
End Sub

Sub Ablock_Click()
    If 正在答題 Then
        questions.Cells(N, 7).Value = 1
        N = N + 1
        出題
    End If
End Sub

Sub Bblock_Click()
    If 正在答題 Then
        questions.Cells(N, 7).Value = 2
        N = N + 1
        出題
    End If
End Sub

Sub Cblock_Click()
    If 正在答題 Then
        questions.Cells(N, 7).Value = 3
        N = N + 1
        出題
    End If
End Sub

Sub Dblock_Click()
    If 正在答題 Then
        questions.Cells(N, 7).Value = 4
        N = N + 1
        出題
    End If
End Sub

Sub 清空題目與選項()
    textblock.TextFrame.Characters.Text = ""
    Ablock.TextFrame.Characters.Text = ""
    Bblock.TextFrame.Characters.Text = ""
    Cblock.TextFrame.Characters.Text = ""
    Dblock.TextFrame.Characters.Text = ""
End Sub

Sub 初始化文字方塊()
    Set textblock = ActiveSheet.Shapes("question_block")
    Set Ablock = ActiveSheet.Shapes("A選項")
    Set Bblock = ActiveSheet.Shapes("B選項")
    Set Cblock = ActiveSheet.Shapes("C選項")
    Set Dblock = ActiveSheet.Shapes("D選項")
End Sub

Sub 退休金_Click()
    開始答題 "退休金"
End Sub
Sub 第一桶金_Click()
    開始答題 "第一桶金"
End Sub
Sub 財產保值_Click()
    開始答題 "財產保值"
End Sub
Sub 教育存款_Click()
    開始答題 "教育存款"
End Sub
Sub 買房_Click()
    開始答題 "買房"
End Sub
Sub 長期財富累積_Click()
    開始答題 "長期財富累積"
End Sub
```
---
## Additional Navigation Code
Simple subs to navigate between sheets.

```vb
Sub nextpage_click()
    Dim nextpage As String
    nextpage = ActiveSheet.Name & "結果單"
    ThisWorkbook.Worksheets(nextpage).Activate
End Sub
```
---

## Process & External Script Integration

Integrates external Python script execution and process management.

```vb
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Function Delay(seconds As Single)
    Sleep seconds * 1000
End Function

Sub 按鈕1_Click()
    Dim pythonExePath As String
    Dim pythonScriptPath As String
    Dim shellProcessID As Long
    Dim fileName As String

    pythonExePath = """C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe"""
    fileName = "main.py"
    pythonScriptPath = ThisWorkbook.Path & "\" & fileName

    shellProcessID = Shell(pythonExePath & " " & pythonScriptPath, vbNormalFocus)

    Do While IsProcessRunning(shellProcessID)
        DoEvents
    Loop

    CloseSpecificWorkbook
End Sub

Function CloseSpecificWorkbook()
    Dim xlApp As Object
    Dim targetWorkbook As Object
    Set xlApp = GetObject(, "Excel.Application")
    For Each targetWorkbook In xlApp.Workbooks
        If targetWorkbook.Name = "taiex_mid100_stock_data.xlsx" Then
            targetWorkbook.Close SaveChanges:=False
            Exit For
        End If
    Next targetWorkbook
End Function

Function IsProcessRunning(PID As Long) As Boolean
    Dim objWMIService As Object, colProcessList As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & PID)
    IsProcessRunning = (colProcessList.Count > 0)
End Function

```

---

## Simulation and Optimization Code

Generates random weights, calculates performance metrics, and applies Monte Carlo simulation.

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
    
    Set ws = ThisWorkbook.Worksheets("幕後結果_最佳化")
    ws.Range("B1:K1").Copy Destination:=ws.Range("AN1")
    Application.CutCopyMode = False
    
    For Each cell In ws.Range("AN1:AW1")
        If Right(cell.Value, 1) <> "W" Then
            cell.Value = cell.Value & "W"
        End If
    Next cell
    
    ws.Range("AN2:BB2").Value = 0
    ws.Range("AN3:BB1001").ClearContents
    Sheets("幕後結果_最佳化").Select
    Range("AN3").Select
    
    formulaString1 = "=B1004/SUM($B1004:$K1004)"
    formulaString2 = "=SUM(AN3:AW3)"
    formulaString3 = "=SUMPRODUCT(AN3:AW3,幕後結果_最佳化!$Z$1005:$AI$1005)"
    formulaString4 = "=(SUMPRODUCT(AN3:AW3,MMULT(AN3:AW3,幕後結果!$AB$9:$AK$18)))*252"
    formulaString5 = "=AZ3^0.5"
    
    ws.Range("AN3").Formula2 = formulaString1
    ws.Range("AN3").Copy
    ws.Range("AO3:AW3").PasteSpecial
    ws.Range("AN3:AW3").Copy
    ws.Range("AN4:AW1001").PasteSpecial
    Application.CutCopyMode = False
    
    ws.Range("AX3").Formula2 = formulaString2
    ws.Range("AX3").AutoFill Destination:=ws.Range("AX3:AX1001")
    ws.Range("AX3:AX1001").FillDown
    
    ws.Range("AY3").Formula2 = formulaString3
    ws.Range("AY3").AutoFill Destination:=ws.Range("AY3:AY1001")
    ws.Range("AY3:AY1001").FillDown
    
    ws.Range("AZ3").Formula2 = formulaString4
    ws.Range("AZ3").AutoFill Destination:=ws.Range("AZ3:AZ1001")
    ws.Range("AZ3:AZ1001").FillDown
    
    ws.Range("BA3").Formula2 = formulaString5
    ws.Range("BA3").AutoFill Destination:=ws.Range("BA3:BA1001")
    ws.Range("BA3:BA1001").FillDown
End Sub

Sub SharpeRatio_formula()
    Dim ws As Worksheet
    Dim lastRow As Long, meanValue As Double, stdDevValue As Double
    Dim sharpeRange As Range, cell As Range

    Set ws = ThisWorkbook.Sheets("幕後結果_最佳化")
    lastRow = 1001
    ws.Range("BB3").Formula = "=(AY3-$BC$2)/$BA3"
    ws.Range("BB3").AutoFill Destination:=ws.Range("BB3:BB" & lastRow)
    ws.Range("BB3:BB" & lastRow).FillDown

    Set sharpeRange = ws.Range("BB3:BB" & lastRow)
    meanValue = Application.WorksheetFunction.Average(sharpeRange)
    stdDevValue = Application.WorksheetFunction.StDev(sharpeRange)
    
    Dim rng As Range
    Set rng = ws.Range("AM1:BB" & lastRow)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(125, 74, 43)
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
```
---

## PDF Export & Chart Generation Code
Handles generating pie charts and exporting worksheet ranges as PDF files.

```vb
Sub 列印風險厭惡結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet

    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    PDFName1 = "風險厭惡結果單.pdf"
    PDFName2 = "風險厭惡.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    Set ws1 = Sheets("風險厭惡結果單")
    Set ws2 = Sheets("風險厭惡")
    
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險厭惡結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    shp1.Visible = msoFalse: shp2.Visible = msoFalse: shp3.Visible = msoFalse
    shp4.Visible = msoFalse: shp5.Visible = msoFalse

    With ws1.PageSetup
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
    ws1.Range("E2:S56").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard
    
    With ws2.PageSetup
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    MsgBox "PDF檔案已儲存到桌面：" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    shp1.Visible = msoTrue: shp2.Visible = msoTrue: shp3.Visible = msoTrue
    shp4.Visible = msoTrue: shp5.Visible = msoTrue
End Sub

Sub GeneratePieChart_雪天型()
    Dim ws1 As Worksheet, ws2 As Worksheet, chartObj As ChartObject, chart As Chart
    Dim dataRange As Range, formulaString1 As String, targetRange As Range, i As Integer

    Set ws1 = ThisWorkbook.Worksheets("風險厭惡結果單")
    Set ws2 = ThisWorkbook.Worksheets("幕後結果_最佳化")

    For Each chartObj In ws1.ChartObjects
        chartObj.Delete
    Next chartObj

    Set dataRange = ws1.Range("H35:Q35")
    formulaString1 = "=H$33*幕後結果_最佳化!Z$1005/$H$34"
    Set targetRange = ws1.Range("H35:Q35")
    targetRange.Formula = formulaString1

    Set chartObj = ws1.ChartObjects.Add(Left:=450, Width:=375, Top:=550, Height:=245)
    Set chart = chartObj.Chart

    With chart
        .ChartType = xlPie
        .SetSourceData Source:=dataRange
        .HasTitle = True: .ChartTitle.Text = "報酬率佔比圖"
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

Sub 列印風險中立結果()
    Dim FilePath1 As String, FilePath2 As String
    Dim PDFName1 As String, PDFName2 As String
    Dim DesktopPath As String
    Dim ws1 As Worksheet, ws2 As Worksheet

    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    PDFName1 = "風險中立結果單.pdf"
    PDFName2 = "風險中立.pdf"
    FilePath1 = DesktopPath & "\" & PDFName1
    FilePath2 = DesktopPath & "\" & PDFName2

    Set ws1 = Sheets("風險中立結果單")
    Set ws2 = Sheets("風險中立")
    
    Dim shp1 As Shape, shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape
    Set shp1 = ws1.Shapes("反回人格頁")
    Set shp2 = ws1.Shapes("列印風險中立結果")
    Set shp3 = ws1.Shapes("返回圖")
    Set shp4 = ws2.Shapes("返回重測")
    Set shp5 = ws2.Shapes("nextpage")
    
    shp1.Visible = msoFalse: shp2.Visible = msoFalse: shp3.Visible = msoFalse
    shp4.Visible = msoFalse: shp5.Visible = msoFalse

    With ws1.PageSetup
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
    ws1.Range("E2:S55").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath1, Quality:=xlQualityStandard
    
    With ws2.PageSetup
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
    ws2.Range("J2:W51").ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath2, Quality:=xlQualityStandard

    MsgBox "PDF檔案已儲存到桌面：" & vbCrLf & FilePath1 & vbCrLf & FilePath2
    
    shp1.Visible = msoTrue: shp2.Visible = msoTrue: shp3.Visible = msoTrue
    shp4.Visible = msoTrue: shp5.Visible = msoTrue
End Sub
```
---

## Additional Navigation Code
Simple subs for sheet navigation.

```vb
Sub 矩形圓角3_Click()
    Sheets("POS機按鈕").Activate
End Sub

Sub 返回風險中立偏厭人格頁_Click()
    Sheets("風險中立偏厭惡").Activate
End Sub

Sub 返回風險厭惡人格介面_Click()
    Sheets("風險厭惡").Activate
End Sub

Sub 返回風險中立人格頁面_Click()
    Sheets("風險中立").Activate
End Sub

Sub 返回風險中立偏愛好人格頁_Click()
    Sheets("風險中立偏愛好").Activate
End Sub

Sub 返回風險愛好人格頁_Click()
    Sheets("風險愛好").Activate
End Sub

```
---
## Note
This README provides a summary of the key modules and functionalities. For a full view of the complete 3000+ lines of code, please refer to the file [Full VBA Code](vba_full_code.md)
 in this repository.