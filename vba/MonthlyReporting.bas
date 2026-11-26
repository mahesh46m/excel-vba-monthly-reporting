Attribute VB_Name = "MonthlyReporting"
' =====================================================================
' MonthlyReporting.bas - monthly reporting automation demo
' Author: Mahesh M
'
' Imports a raw monthly transactions CSV, cleans it (dedupe on txn_id,
' trim labels, normalise mixed date formats), builds a formatted summary
' sheet with KPIs, and applies report formatting.
'
' How to run:
'   1. Open a blank workbook and save it as .xlsm (macro-enabled).
'   2. Alt+F11 -> File -> Import File -> select this .bas file.
'   3. Run BuildMonthlyReport and pick transactions_2026-01.csv.
'
' The demo ships with expected_summary.csv (computed independently in
' Python) so the macro's output can be verified against ground truth.
' All sample data is synthetic.
' =====================================================================
Option Explicit

Private Const SHEET_RAW As String = "Raw"
Private Const SHEET_SUMMARY As String = "Summary"

Public Sub BuildMonthlyReport()
    Dim csvPath As Variant
    csvPath = Application.GetOpenFilename( _
        FileFilter:="CSV files (*.csv),*.csv", _
        Title:="Select the monthly transactions CSV")
    If csvPath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ImportCsv CStr(csvPath)
    Dim clean As Collection
    Set clean = CleanRaw()
    WriteSummary clean
    FormatSummary

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ThisWorkbook.Sheets(SHEET_SUMMARY).Activate
    MsgBox "Monthly report built: " & clean.Count & " clean rows.", vbInformation
End Sub

' Import the CSV into a fresh "Raw" sheet, values only.
Private Sub ImportCsv(ByVal csvPath As String)
    Dim ws As Worksheet
    Set ws = FreshSheet(SHEET_RAW)

    Dim wb As Workbook
    Set wb = Workbooks.Open(csvPath, ReadOnly:=True)
    wb.Sheets(1).UsedRange.Copy
    ws.Range("A1").PasteSpecial xlPasteValues
    wb.Close SaveChanges:=False
    Application.CutCopyMode = False
End Sub

' Remove an existing sheet if present, then add a fresh one.
Private Function FreshSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    Set FreshSheet = ws
End Function

' Clean the raw rows: dedupe on txn_id, trim text, normalise dates to
' yyyy-mm-dd (accepts ISO and m/d/yyyy). Returns a Collection of arrays:
' txn_id, date(ISO string), category, region, rep, amount.
Private Function CleanRaw() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_RAW)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim result As New Collection

    Dim i As Long
    For i = 2 To lastRow   ' row 1 is the header
        Dim txnId As String
        txnId = Trim$(CStr(ws.Cells(i, 1).Value))
        If Len(txnId) = 0 Then GoTo NextRow
        If seen.Exists(txnId) Then GoTo NextRow
        seen.Add txnId, True

        Dim isoDate As String
        isoDate = NormaliseDate(Trim$(CStr(ws.Cells(i, 2).Value)))
        Dim amount As Double
        amount = Val(Replace(CStr(ws.Cells(i, 6).Value), ",", ""))

        result.Add Array( _
            txnId, _
            isoDate, _
            Trim$(CStr(ws.Cells(i, 3).Value)), _
            Trim$(CStr(ws.Cells(i, 4).Value)), _
            Trim$(CStr(ws.Cells(i, 5).Value)), _
            amount)
NextRow:
    Next i
    Set CleanRaw = result
End Function

' Accept "yyyy-mm-dd" or "m/d/yyyy" and return "yyyy-mm-dd".
Private Function NormaliseDate(ByVal raw As String) As String
    If InStr(raw, "/") > 0 Then
        Dim parts() As String
        parts = Split(raw, "/")
        NormaliseDate = Format$(CLng(parts(2)), "0000") & "-" & _
                        Format$(CLng(parts(0)), "00") & "-" & _
                        Format$(CLng(parts(1)), "00")
    Else
        NormaliseDate = raw
    End If
End Function

' Aggregate into a dictionary of dimension|member -> total, then write.
Private Sub WriteSummary(ByVal clean As Collection)
    Dim ws As Worksheet
    Set ws = FreshSheet(SHEET_SUMMARY)

    Dim totals As Object
    Set totals = CreateObject("Scripting.Dictionary")
    Dim grand As Double
    Dim item As Variant
    For Each item In clean
        grand = grand + item(5)
        AddTotal totals, "Category", item(2), item(5)
        AddTotal totals, "Region", item(3), item(5)
    Next item

    With ws
        .Range("A1").Value = "Monthly Sales Report - January 2026 (synthetic data)"
        .Range("A3").Value = "Clean rows"
        .Range("B3").Value = clean.Count
        .Range("A4").Value = "Grand total"
        .Range("B4").Value = grand
        .Range("B4").NumberFormat = "$#,##0.00"
        .Range("A5").Value = "Generated"
        .Range("B5").Value = Format$(Now, "yyyy-mm-dd hh:nn")

        .Range("A7").Value = "Dimension"
        .Range("B7").Value = "Member"
        .Range("C7").Value = "Total amount"

        Dim r As Long
        r = 8
        Dim k As Variant
        For Each k In totals.Keys
            Dim dimName As String, member As String
            dimName = Split(k, "|")(0)
            member = Split(k, "|")(1)
            .Cells(r, 1).Value = dimName
            .Cells(r, 2).Value = member
            .Cells(r, 3).Value = totals(k)
            .Cells(r, 3).NumberFormat = "$#,##0.00"
            r = r + 1
        Next k
    End With
End Sub

Private Sub AddTotal(ByVal totals As Object, ByVal dimName As String, _
                     ByVal member As String, ByVal amount As Double)
    Dim key As String
    key = dimName & "|" & member
    If totals.Exists(key) Then
        totals(key) = totals(key) + amount
    Else
        totals.Add key, amount
    End If
End Sub

' Presentation: title style, bold header row, autofit, freeze panes,
' highlight the top-earning category row.
Private Sub FormatSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SUMMARY)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    With ws
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A7:C7").Font.Bold = True
        .Range("A7:C7").Interior.Color = RGB(31, 56, 100)
        .Range("A7:C7").Font.Color = RGB(255, 255, 255)
        .Columns("A:C").AutoFit
        .Range("A8").Select
        ActiveWindow.FreezePanes = True

        ' Highlight the top category row (within the Category block).
        Dim r As Long, bestRow As Long
        Dim best As Double
        best = -1
        For r = 8 To lastRow
            If .Cells(r, 1).Value = "Category" Then
                If .Cells(r, 3).Value > best Then
                    best = .Cells(r, 3).Value
                    bestRow = r
                End If
            End If
        Next r
        If bestRow > 0 Then
            .Range(.Cells(bestRow, 1), .Cells(bestRow, 3)).Interior.Color = RGB(198, 239, 206)
        End If
    End With
End Sub
