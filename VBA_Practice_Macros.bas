Attribute VB_Name = "PracticeMacros"
Option Explicit

' ============================================================
' Practice macros for Kaiser_Practice_Dataset.xlsx (synthetic data)
' Sheets used: Assumptions, Members, Claims, Pension
' Import: Alt+F11 > File > Import File... > select this .bas
' Save the workbook as .xlsm to keep the macros.
' ============================================================

' ---------- Task 1: Highlight high-risk members ----------
Public Sub HighlightHighRisk()
    Dim ws As Worksheet
    Dim threshold As Double
    Dim lastRow As Long, i As Long, cnt As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("Members")
    threshold = ThisWorkbook.Sheets("Assumptions").Range("B8").Value
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False
    ws.Range("A2:K" & lastRow).Interior.ColorIndex = xlNone   ' reset

    For i = 2 To lastRow
        If ws.Cells(i, 8).Value >= threshold Then              ' column H = Risk Score
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 11)).Interior.Color = RGB(255, 235, 156)
            cnt = cnt + 1
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "High-risk members (HCC >= " & threshold & "): " & cnt
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

' ---------- Task 4: Standardize the Plan column ----------
Public Sub StandardizePlans()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, fixedCount As Long
    Dim raw As String, clean As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("Members")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False
    For i = 2 To lastRow
        raw = CStr(ws.Cells(i, 5).Value)
        Select Case UCase(Trim(raw))
            Case "MEDICARE ADVANTAGE": clean = "Medicare Advantage"
            Case "COMMERCIAL":         clean = "Commercial"
            Case "MEDICAID":           clean = "Medicaid"
            Case Else:                 clean = raw               ' leave unknown values untouched
        End Select
        If clean <> raw Then
            ws.Cells(i, 5).Value = clean
            fixedCount = fixedCount + 1
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Plan names standardized: " & fixedCount & " cells fixed."
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

' ---------- Task 2 + 5 + 8: Summary by Plan (PMPM, risk-adjusted, trend) ----------
Public Sub SummaryByPlan()
    Dim wsM As Worksheet, wsOut As Worksheet
    Dim plans As Variant
    Dim lastRow As Long, k As Long, r As Long
    Dim nMembers As Long
    Dim months As Double, cost As Double, pmpm As Double, avgRisk As Double
    Dim trend As Double

    On Error GoTo ErrHandler
    Call StandardizePlans                       ' SUMIFS needs clean text

    Set wsM = ThisWorkbook.Sheets("Members")
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    trend = ThisWorkbook.Sheets("Assumptions").Range("B7").Value
    plans = Array("Medicare Advantage", "Commercial", "Medicaid")

    Set wsOut = GetFreshSheet("Summary")
    wsOut.Range("A1:H1").Value = Array("Plan", "Members", "Member Months", "Total Cost", _
        "PMPM", "Avg Risk Score", "Risk-Adj PMPM", "Projected Next-Yr Cost")
    wsOut.Range("A1:H1").Font.Bold = True

    r = 2
    For k = LBound(plans) To UBound(plans)
        nMembers = Application.WorksheetFunction.CountIfs(wsM.Range("E2:E" & lastRow), plans(k))
        months = Application.WorksheetFunction.SumIfs(wsM.Range("F2:F" & lastRow), wsM.Range("E2:E" & lastRow), plans(k))
        cost = Application.WorksheetFunction.SumIfs(wsM.Range("I2:I" & lastRow), wsM.Range("E2:E" & lastRow), plans(k))

        If nMembers > 0 Then
            avgRisk = Application.WorksheetFunction.AverageIfs(wsM.Range("H2:H" & lastRow), wsM.Range("E2:E" & lastRow), plans(k))
        Else
            avgRisk = 0
        End If
        If months > 0 Then pmpm = cost / months Else pmpm = 0

        wsOut.Cells(r, 1).Value = plans(k)
        wsOut.Cells(r, 2).Value = nMembers
        wsOut.Cells(r, 3).Value = months
        wsOut.Cells(r, 4).Value = cost
        wsOut.Cells(r, 5).Value = pmpm
        wsOut.Cells(r, 6).Value = avgRisk
        If avgRisk > 0 Then wsOut.Cells(r, 7).Value = pmpm / avgRisk
        wsOut.Cells(r, 8).Value = cost * (1 + trend)
        r = r + 1
    Next k

    wsOut.Range("D2:E" & r).NumberFormat = "$#,##0.00"
    wsOut.Range("G2:H" & r).NumberFormat = "$#,##0.00"
    wsOut.Range("F2:F" & r).NumberFormat = "0.000"
    wsOut.Columns("A:H").AutoFit
    MsgBox "Summary created on sheet 'Summary'."
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

' ---------- Task 5: Cost by Claim Category and by Region ----------
Public Sub CostByCategoryAndRegion()
    Dim wsC As Worksheet, wsM As Worksheet, wsOut As Worksheet
    Dim cats As Variant, regs As Variant
    Dim lastC As Long, lastM As Long, k As Long, r As Long

    On Error GoTo ErrHandler
    Set wsC = ThisWorkbook.Sheets("Claims")
    Set wsM = ThisWorkbook.Sheets("Members")
    lastC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    lastM = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row

    cats = Array("Inpatient", "ER", "Outpatient", "Professional", "Pharmacy")
    regs = Array("Northern California", "Southern California", "Colorado", "Georgia", _
                 "Hawaii", "Mid-Atlantic", "Northwest", "Washington")

    Set wsOut = GetFreshSheet("Cost_Breakdown")
    wsOut.Range("A1:B1").Value = Array("Claim Category", "Total Paid")
    wsOut.Range("D1:E1").Value = Array("Region", "Total Cost")
    wsOut.Range("A1:E1").Font.Bold = True

    For k = LBound(cats) To UBound(cats)
        wsOut.Cells(k + 2, 1).Value = cats(k)
        wsOut.Cells(k + 2, 2).Value = Application.WorksheetFunction.SumIfs( _
            wsC.Range("E2:E" & lastC), wsC.Range("D2:D" & lastC), cats(k))
    Next k

    For k = LBound(regs) To UBound(regs)
        wsOut.Cells(k + 2, 4).Value = regs(k)
        wsOut.Cells(k + 2, 5).Value = Application.WorksheetFunction.SumIfs( _
            wsM.Range("I2:I" & lastM), wsM.Range("D2:D" & lastM), regs(k))
    Next k

    wsOut.Range("B2:B10,E2:E10").NumberFormat = "$#,##0"
    wsOut.Columns("A:E").AutoFit
    MsgBox "Breakdown created on sheet 'Cost_Breakdown'."
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

' ---------- Task 3: Data-quality report ----------
Public Sub DataQualityCheck()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim rawPlan As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("Members")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set wsOut = GetFreshSheet("DQ_Report")
    wsOut.Range("A1:C1").Value = Array("Row", "Member ID", "Issue")
    wsOut.Range("A1:C1").Font.Bold = True
    r = 2

    Application.ScreenUpdating = False
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 4).Value)) = "" Then
            wsOut.Range("A" & r & ":C" & r).Value = Array(i, ws.Cells(i, 1).Value, "Blank Region")
            r = r + 1
        End If

        If Application.WorksheetFunction.CountIf(ws.Range("A2:A" & lastRow), ws.Cells(i, 1).Value) > 1 Then
            wsOut.Range("A" & r & ":C" & r).Value = Array(i, ws.Cells(i, 1).Value, "Duplicate Member ID")
            r = r + 1
        End If

        rawPlan = CStr(ws.Cells(i, 5).Value)
        If StrComp(rawPlan, "Medicare Advantage", vbBinaryCompare) <> 0 And _
           StrComp(rawPlan, "Commercial", vbBinaryCompare) <> 0 And _
           StrComp(rawPlan, "Medicaid", vbBinaryCompare) <> 0 Then
            wsOut.Range("A" & r & ":C" & r).Value = Array(i, ws.Cells(i, 1).Value, "Non-standard Plan text: '" & rawPlan & "'")
            r = r + 1
        End If
    Next i
    Application.ScreenUpdating = True

    wsOut.Columns("A:C").AutoFit
    MsgBox "Data-quality check complete. Issues found: " & (r - 2)
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

' ---------- Task 6: Actives within 5 years of age 65 ----------
Public Sub PensionNearRetirement()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, cnt As Long
    Dim totalPV As Double

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("Pension")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False
    For i = 2 To lastRow
        If ws.Cells(i, 4).Value = "Active" Then
            If ws.Cells(i, 7).Value >= 60 And ws.Cells(i, 7).Value < 65 Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 10)).Interior.Color = RGB(198, 224, 180)
                cnt = cnt + 1
                totalPV = totalPV + ws.Cells(i, 10).Value
            End If
        End If
    Next i
    Application.ScreenUpdating = True

    MsgBox "Active participants age 60-64: " & cnt & vbCrLf & _
           "Total PV of accrued benefit: " & Format(totalPV, "$#,##0")
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

' ---------- Task 7: User-defined function ----------
' Use in a cell:  =AccruedBenefit(H2, F2, Assumptions!$B$4)
Public Function AccruedBenefit(service As Double, fas As Double, multiplier As Double) As Double
    AccruedBenefit = service * multiplier * fas
End Function

' ---------- Helper ----------
Private Function GetFreshSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    Set GetFreshSheet = ws
End Function
