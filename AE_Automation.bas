Attribute VB_Name = "Module1"
Option Explicit
Dim importOK As Boolean
Sub SetupWorkbook()
    Dim sheetNames As Variant
    Dim i As Long
    Dim ws As Worksheet

    sheetNames = Array("Control", "Raw", "Clean", "Summary")

    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetNames(i))
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = sheetNames(i)
        End If
    Next i

    MsgBox "Workbook ready: Control, Raw, Clean, Summary", vbInformation
End Sub

Sub ImportRawData()
    Dim filePath As Variant
    Dim srcWb As Workbook
    Dim rawWs As Worksheet
    importOK = False
    On Error GoTo ErrHandler

    filePath = Application.GetOpenFilename( _
        FileFilter:="Data files (*.csv;*.xls;*.xlsx),*.csv;*.xls;*.xlsx", _
        Title:="Select the NHS A&E data file")

    If VarType(filePath) = vbBoolean Then
        MsgBox "No file selected. Import cancelled.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set rawWs = ThisWorkbook.Worksheets("Raw")
    rawWs.Cells.Clear

    Set srcWb = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    srcWb.Worksheets(1).UsedRange.Copy Destination:=rawWs.Range("A1")
    srcWb.Close SaveChanges:=False

    Application.ScreenUpdating = True
    importOK = True
    MsgBox "Imported " & rawWs.UsedRange.Rows.Count - 1 & " rows.", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

Sub CleanData()
    Dim rawWs As Worksheet, cleanWs As Worksheet
    Dim data As Variant, out() As Variant
    Dim lastRow As Long, r As Long, c As Long, n As Long
    Dim totalAtt As Double, totalOver As Double
    Dim t1Att As Double, t1Over As Double
    Dim parts() As String

    On Error GoTo ErrHandler

    Set rawWs = ThisWorkbook.Worksheets("Raw")
    Set cleanWs = ThisWorkbook.Worksheets("Clean")

    ' 1. Check columns E:J are attendances and K:P are over-4hr waits
    For c = 5 To 10
        If Not rawWs.Cells(1, c).Value Like "A&E attendances*" Then
            MsgBox "Unexpected header in column " & c & ": " & rawWs.Cells(1, c).Value, vbCritical
            Exit Sub
        End If
    Next c
    For c = 11 To 16
        If Not rawWs.Cells(1, c).Value Like "Attendances over 4*" Then
            MsgBox "Unexpected header in column " & c & ": " & rawWs.Cells(1, c).Value, vbCritical
            Exit Sub
        End If
    Next c

    ' 2. Read the whole Raw table into memory
    lastRow = rawWs.Cells(rawWs.Rows.Count, "A").End(xlUp).Row
    data = rawWs.Range("A1:P" & lastRow).Value

    ReDim out(1 To lastRow, 1 To 10)
    out(1, 1) = "Period": out(1, 2) = "Org Code"
    out(1, 3) = "Region": out(1, 4) = "Org Name"
    out(1, 5) = "Type 1 Attendances": out(1, 6) = "Total Attendances"
    out(1, 7) = "Type 1 Over 4hrs": out(1, 8) = "Total Over 4hrs"
    out(1, 9) = "Type 1 4hr %": out(1, 10) = "All Types 4hr %"
    n = 1

    ' 3. Process each row
    For r = 2 To UBound(data, 1)
        totalAtt = 0: totalOver = 0
        For c = 5 To 10: totalAtt = totalAtt + Val(data(r, c)): Next c
        For c = 11 To 16: totalOver = totalOver + Val(data(r, c)): Next c

                If totalAtt > 0 And data(r, 1) Like "MSitAE-*-*" Then
            n = n + 1
            parts = Split(data(r, 1), "-")
            out(n, 1) = CDate("1 " & parts(1) & " " & parts(2))
            out(n, 2) = data(r, 2)
            out(n, 3) = data(r, 3)
            out(n, 4) = data(r, 4)

            t1Att = Val(data(r, 5)) + Val(data(r, 8))
            t1Over = Val(data(r, 11)) + Val(data(r, 14))
            out(n, 5) = t1Att
            out(n, 6) = totalAtt
            out(n, 7) = t1Over
            out(n, 8) = totalOver
            If t1Att > 0 Then out(n, 9) = 1 - t1Over / t1Att Else out(n, 9) = ""
            out(n, 10) = 1 - totalOver / totalAtt
        End If
    Next r

    ' 4. Write the result to Clean in one go, then format it
    cleanWs.Cells.Clear
    cleanWs.Range("A1").Resize(n, 10).Value = out
    cleanWs.Range("A2:A" & n).NumberFormat = "mmm yyyy"
    cleanWs.Range("E2:H" & n).NumberFormat = "#,##0"
    cleanWs.Range("I2:J" & n).NumberFormat = "0.0%"
    cleanWs.Rows(1).Font.Bold = True
    cleanWs.Columns("A:J").AutoFit

    MsgBox "Kept " & n - 1 & " rows. Removed " & (lastRow - n) & " rows with no attendances.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Cleaning failed on row " & r & ": " & Err.Description, vbCritical
End Sub

Sub BuildSummary()
    Dim cleanWs As Worksheet, sumWs As Worksheet
    Dim data As Variant, out() As Variant
    Dim dict As Object, key As Variant, totals As Variant
    Dim lastRow As Long, r As Long, i As Long, n As Long
    Dim grand(0 To 3) As Double

    On Error GoTo ErrHandler
    Set cleanWs = ThisWorkbook.Worksheets("Clean")
    Set sumWs = ThisWorkbook.Worksheets("Summary")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = cleanWs.Cells(cleanWs.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Clean sheet is empty. Run CleanData first.", vbExclamation
        Exit Sub
    End If
    data = cleanWs.Range("A1:J" & lastRow).Value

    ' 1. Group by Region (column C) and add up columns E:H
    For r = 2 To UBound(data, 1)
        key = data(r, 3)
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#, 0#)
        totals = dict(key)
        For i = 0 To 3
            totals(i) = totals(i) + data(r, 5 + i)
            grand(i) = grand(i) + data(r, 5 + i)
        Next i
        dict(key) = totals
    Next r

    ' 2. Build the output table
    ReDim out(1 To dict.Count + 2, 1 To 7)
    out(1, 1) = "Region": out(1, 2) = "Type 1 Attendances"
    out(1, 3) = "Total Attendances": out(1, 4) = "Type 1 Over 4hrs"
    out(1, 5) = "Total Over 4hrs": out(1, 6) = "Type 1 4hr %"
    out(1, 7) = "All Types 4hr %"
    n = 1

    For Each key In dict.Keys
        n = n + 1
        totals = dict(key)
        out(n, 1) = Replace(Replace(StrConv(Replace(key, "NHS ENGLAND ", ""), vbProperCase), "Of", "of"), "And", "and")
        For i = 0 To 3: out(n, 2 + i) = totals(i): Next i
        If totals(0) > 0 Then out(n, 6) = 1 - totals(2) / totals(0) Else out(n, 6) = ""
        out(n, 7) = 1 - totals(3) / totals(1)
    Next key

    n = n + 1
    out(n, 1) = "ENGLAND"
    For i = 0 To 3: out(n, 2 + i) = grand(i): Next i
    out(n, 6) = 1 - grand(2) / grand(0)
    out(n, 7) = 1 - grand(3) / grand(1)

    ' 3. Write and format
    sumWs.Cells.Clear
    With sumWs.Range("A1")
        .Value = "A&E 4-hour performance by region - " & Format(data(2, 1), "mmmm yyyy")
        .Font.Bold = True
        .Font.Size = 14
    End With
    sumWs.Range("A3").Resize(n, 7).Value = out
        ' Sort regions best to worst - England row stays at the bottom
    sumWs.Range("A4:G" & n + 1).Sort _
        Key1:=sumWs.Range("G4"), Order1:=xlDescending, Header:=xlNo
    sumWs.Range("B4:E" & n + 2).NumberFormat = "#,##0"
    sumWs.Range("F4:G" & n + 2).NumberFormat = "0.0%"
    sumWs.Rows(3).Font.Bold = True
    sumWs.Rows(n + 2).Font.Bold = True
    sumWs.Range("A3:G" & n + 2).Columns.AutoFit

    MsgBox "Summary built for " & dict.Count & " regions.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Summary failed: " & Err.Description, vbCritical
End Sub

Sub BuildChart()
    Dim sumWs As Worksheet
    Dim lastRow As Long
    Dim co As ChartObject

    Set sumWs = ThisWorkbook.Worksheets("Summary")
    lastRow = sumWs.Cells(sumWs.Rows.Count, "A").End(xlUp).Row

    ' Remove any old chart so re-runs don't stack charts
    Do While sumWs.ChartObjects.Count > 0
        sumWs.ChartObjects(1).Delete
    Loop

        Set co = sumWs.ChartObjects.Add( _
        Left:=sumWs.Range("A14").Left, Top:=sumWs.Range("A14").Top, _
        Width:=620, Height:=320)

    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=Union(sumWs.Range("A3:A" & lastRow), _
                                     sumWs.Range("G3:G" & lastRow))
        .HasTitle = True
        .ChartTitle.Text = "All types 4-hour performance by region"
        .HasLegend = False
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlValue).MinimumScale = 0.5
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.NumberFormat = "0.0%"
        ' Make the ENGLAND bar stand out
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(174, 183, 189)
        .SeriesCollection(1).Points(lastRow - 3).Format.Fill.ForeColor.RGB = RGB(0, 94, 184)
        .Axes(xlCategory).Crosses = xlMaximum
    End With
End Sub
Sub ExportPDF()
    Dim sumWs As Worksheet
    Dim folder As String, pdfPath As String

    Set sumWs = ThisWorkbook.Worksheets("Summary")

    folder = ThisWorkbook.Path
    If Left(folder, 4) = "http" Then folder = Environ("USERPROFILE") & "\Documents"

    pdfPath = folder & "\AE_Summary_" & _
        Format(ThisWorkbook.Worksheets("Clean").Range("A2").Value, "yyyy-mm") & ".pdf"

    With sumWs.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    sumWs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, OpenAfterPublish:=True
End Sub

Sub RunAll()
    ImportRawData
    If Not importOK Then Exit Sub
    CleanData
    BuildSummary
    BuildChart
    ExportPDF
End Sub
