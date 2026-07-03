' ==============================================================
' Jira Weekly Report Generator - VBA Macro v8
'
' Changes from v7:
'   Each Report tab is now split into 3 focused tabs:
'   -Status     : Blocked + Assigned To Do
'   -Unassigned : Unassigned Open Tickets
'   -Cleanup    : Assigned Open Tickets With Missing Fields
'
' Full tab list after running:
'   Overall-Status,   Overall-Unassigned,   Overall-Cleanup
'   Overall-Charts
'   Story-Status,     Story-Unassigned,     Story-Cleanup
'   Story-Charts
'   SubTask-Status,   SubTask-Unassigned,   SubTask-Cleanup
'   SubTask-Charts
'   Epic-Status,      Epic-Unassigned,      Epic-Cleanup
'   Epic-Charts
'   (plus original Report and Charts tabs kept for summary)
'
' CSV Columns (A to M):
'   A: Issue Type       B: Key            C: Summary
'   D: Status           E: Assignee       F: Description
'   G: Component/s      H: Labels         I: Priority
'   J: Requested by Date (Schwab)         K: Time Tracking
'   L: Created          M: Creator
'
' Instructions:
'   1. Press Alt+F11 to open the VBA Editor
'   2. Click Insert > Module
'   3. Open this file in Notepad, Ctrl+A, Ctrl+C
'   4. Click in the VBA white area, Ctrl+V
'   5. Press Alt+F8, select GenerateReport, click Run
' ==============================================================

Sub GenerateReport()

    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    Set wsData = ThisWorkbook.Sheets("Data")

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in the Data sheet.", vbExclamation
        GoTo Cleanup
    End If

    Dim totalTickets As Long
    totalTickets = lastRow - 1

    ' ----------------------------------------------------------
    ' Status names
    ' ----------------------------------------------------------
    Dim sName(7) As String
    sName(0) = "In Progress"
    sName(1) = "QA Review"
    sName(2) = "In Approval"
    sName(3) = "Blocked"
    sName(4) = "To Do"
    sName(5) = "Backlog"
    sName(6) = "Done"
    sName(7) = "Canceled"

    Dim sCnt(7) As Long
    For j = 0 To 7
        sCnt(j) = WorksheetFunction.CountIf( _
                  wsData.Range("D2:D" & lastRow), sName(j))
    Next j

    ' ----------------------------------------------------------
    ' Issue Type setup
    ' ----------------------------------------------------------
    Dim tName(2) As String
    Dim tCnt(2) As Long
    tName(0) = "Story"
    tName(1) = "Epic"
    tName(2) = "Sub-task"

    For j = 0 To 2
        tCnt(j) = WorksheetFunction.CountIf( _
                  wsData.Range("A2:A" & lastRow), tName(j))
    Next j

    ' Per-type status counts (0=Story,1=Epic,2=Sub-task)
    Dim tsCount(2, 7) As Long
    Dim typeIdx As Integer
    For i = 2 To lastRow
        Select Case Trim(wsData.Cells(i, 1).Value)
            Case "Story":    typeIdx = 0
            Case "Epic":     typeIdx = 1
            Case "Sub-task": typeIdx = 2
            Case Else:       typeIdx = -1
        End Select
        If typeIdx >= 0 Then
            For j = 0 To 7
                If Trim(wsData.Cells(i, 4).Value) = sName(j) Then
                    tsCount(typeIdx, j) = tsCount(typeIdx, j) + 1
                End If
            Next j
        End If
    Next i

    ' ----------------------------------------------------------
    ' All tabs to create/clear
    ' ----------------------------------------------------------
    Dim allTabs(15) As String
    ' Story
    allTabs(0)  = "Story-Status"
    allTabs(1)  = "Story-Unassigned"
    allTabs(2)  = "Story-Cleanup"
    allTabs(3)  = "Story-Charts"
    ' SubTask
    allTabs(4)  = "SubTask-Status"
    allTabs(5)  = "SubTask-Unassigned"
    allTabs(6)  = "SubTask-Cleanup"
    allTabs(7)  = "SubTask-Charts"
    ' Epic
    allTabs(8)  = "Epic-Status"
    allTabs(9)  = "Epic-Unassigned"
    allTabs(10) = "Epic-Cleanup"
    allTabs(11) = "Epic-Charts"
    ' Overall
    allTabs(12) = "Overall-Status"
    allTabs(13) = "Overall-Unassigned"
    allTabs(14) = "Overall-Cleanup"
    allTabs(15) = "Overall-Charts"




    Dim t As Integer
    For t = 0 To 15
        If Not SheetExists(allTabs(t)) Then
            Dim newWs As Worksheet
            Set newWs = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newWs.Name = allTabs(t)
        End If
    Next t

    ' Tab colours
    ' Overall


    ThisWorkbook.Sheets("Overall-Status").Tab.Color     = RGB(176, 0, 0)
    ThisWorkbook.Sheets("Overall-Unassigned").Tab.Color = RGB(85, 85, 85)
    ThisWorkbook.Sheets("Overall-Cleanup").Tab.Color    = RGB(123, 63, 0)
    ThisWorkbook.Sheets("Overall-Charts").Tab.Color     = RGB(154, 106, 0)
    ' Story
    ThisWorkbook.Sheets("Story-Status").Tab.Color       = RGB(176, 0, 0)
    ThisWorkbook.Sheets("Story-Unassigned").Tab.Color   = RGB(85, 85, 85)
    ThisWorkbook.Sheets("Story-Cleanup").Tab.Color      = RGB(123, 63, 0)
    ThisWorkbook.Sheets("Story-Charts").Tab.Color       = RGB(26, 94, 154)
    ' SubTask
    ThisWorkbook.Sheets("SubTask-Status").Tab.Color     = RGB(176, 0, 0)
    ThisWorkbook.Sheets("SubTask-Unassigned").Tab.Color = RGB(85, 85, 85)
    ThisWorkbook.Sheets("SubTask-Cleanup").Tab.Color    = RGB(123, 63, 0)
    ThisWorkbook.Sheets("SubTask-Charts").Tab.Color     = RGB(26, 122, 68)
    ' Epic
    ThisWorkbook.Sheets("Epic-Status").Tab.Color        = RGB(176, 0, 0)
    ThisWorkbook.Sheets("Epic-Unassigned").Tab.Color    = RGB(85, 85, 85)
    ThisWorkbook.Sheets("Epic-Cleanup").Tab.Color       = RGB(123, 63, 0)
    ThisWorkbook.Sheets("Epic-Charts").Tab.Color        = RGB(108, 45, 142)

    ' Clear all tabs
    For t = 0 To 15
        Dim wsClear As Worksheet
        Set wsClear = ThisWorkbook.Sheets(allTabs(t))
        wsClear.Cells.Clear
        wsClear.Cells.ClearFormats
        On Error Resume Next
        wsClear.ChartObjects.Delete
        On Error GoTo ErrHandler
    Next t

    ' ----------------------------------------------------------
    ' Pre-collect ALL row indexes once (for all scopes)
    ' Scopes: "" = all, "Story", "Epic", "Sub-task"
    ' ----------------------------------------------------------
    Dim filterList(3) As String
    filterList(0) = "Story"
    filterList(1) = "Sub-task"
    filterList(2) = "Epic"
    filterList(3) = ""

    ' Arrays: first index = scope (0-3), second = row up to 500
    Dim bRowsAll(3, 500)    As Long
    Dim dRowsAll(3, 500)    As Long
    Dim uRowsAll(3, 500)    As Long
    Dim incRowsAll(3, 500)  As Long
    Dim bCntAll(3)          As Long
    Dim dCntAll(3)          As Long
    Dim uCntAll(3)          As Long
    Dim incCntAll(3)        As Long

    Dim sc As Integer
    Dim stVal As String
    Dim asVal As String
    Dim iTyp As String

    For i = 2 To lastRow
        iTyp  = Trim(wsData.Cells(i, 1).Value)
        stVal = Trim(wsData.Cells(i, 4).Value)
        asVal = LCase(Trim(wsData.Cells(i, 5).Value))

        ' Determine which scopes this row belongs to
        Dim inScope(3) As Boolean
        inScope(0) = (iTyp = "Story")
        inScope(1) = (iTyp = "Sub-task")
        inScope(2) = (iTyp = "Epic")
        inScope(3) = True

        For sc = 0 To 3
            If inScope(sc) Then

                ' Blocked
                If stVal = "Blocked" Then
                    bRowsAll(sc, bCntAll(sc)) = i
                    bCntAll(sc) = bCntAll(sc) + 1
                End If

                ' Assigned To Do
                If stVal = "To Do" And asVal <> "unassigned" And asVal <> "" Then
                    dRowsAll(sc, dCntAll(sc)) = i
                    dCntAll(sc) = dCntAll(sc) + 1
                End If

                ' Unassigned open (excl Done & Canceled)
                If (asVal = "unassigned" Or asVal = "") _
                   And stVal <> "Done" And stVal <> "Canceled" Then
                    uRowsAll(sc, uCntAll(sc)) = i
                    uCntAll(sc) = uCntAll(sc) + 1
                End If

                ' Assigned + missing applicable fields (excl Done & Canceled)
                If asVal <> "unassigned" And asVal <> "" _
                   And stVal <> "Done" And stVal <> "Canceled" Then
                    Dim aPri As Boolean
                    Dim aRDt As Boolean
                    aPri = (iTyp = "Story" Or iTyp = "Epic")
                    aRDt = (iTyp = "Story")
                    Dim iMiss As Boolean : iMiss = False
                    If Trim(wsData.Cells(i, 6).Value)  = "" Then iMiss = True
                    If Trim(wsData.Cells(i, 7).Value)  = "" Then iMiss = True
                    If Trim(wsData.Cells(i, 8).Value)  = "" Then iMiss = True
                    If aPri And Trim(wsData.Cells(i, 9).Value)  = "" Then iMiss = True
                    If aRDt And Trim(wsData.Cells(i, 10).Value) = "" Then iMiss = True
                    If iMiss Then
                        incRowsAll(sc, incCntAll(sc)) = i
                        incCntAll(sc) = incCntAll(sc) + 1
                    End If
                End If

            End If
        Next sc
    Next i

    ' ----------------------------------------------------------
    ' Scope totals for display
    ' ----------------------------------------------------------
    Dim scopeTotal(3) As Long
    scopeTotal(0) = tCnt(0)       ' Story
    scopeTotal(1) = tCnt(2)       ' Sub-task
    scopeTotal(2) = tCnt(1)       ' Epic
    scopeTotal(3) = totalTickets  ' Overall

    ' Per-scope status counts
    Dim scopeSCnt(3, 7) As Long
    For j = 0 To 7
        scopeSCnt(0, j) = tsCount(0, j)  ' Story
        scopeSCnt(1, j) = tsCount(2, j)  ' Sub-task
        scopeSCnt(2, j) = tsCount(1, j)  ' Epic
        scopeSCnt(3, j) = sCnt(j)        ' Overall
    Next j

    ' ----------------------------------------------------------
    ' Tab name mapping: Status/Unassigned/Cleanup per scope
    ' ----------------------------------------------------------
    Dim stsTab(3)  As String
    Dim unaTab(3)  As String
    Dim clnTab(3)  As String
    Dim chtTab(3)  As String
    Dim scopeLabel(3) As String

    stsTab(0) = "Story-Status"     : unaTab(0) = "Story-Unassigned"
    clnTab(0) = "Story-Cleanup"    : chtTab(0) = "Story-Charts"
    scopeLabel(0) = "Story"

    stsTab(1) = "SubTask-Status"   : unaTab(1) = "SubTask-Unassigned"
    clnTab(1) = "SubTask-Cleanup"  : chtTab(1) = "SubTask-Charts"
    scopeLabel(1) = "Sub-task"

    stsTab(2) = "Epic-Status"      : unaTab(2) = "Epic-Unassigned"
    clnTab(2) = "Epic-Cleanup"     : chtTab(2) = "Epic-Charts"
    scopeLabel(2) = "Epic"

    stsTab(3) = "Overall-Status"   : unaTab(3) = "Overall-Unassigned"
    clnTab(3) = "Overall-Cleanup"  : chtTab(3) = "Overall-Charts"
    scopeLabel(3) = "All Types"

    ' ==========================================================
    ' BUILD ALL TABS
    ' ==========================================================
    For sc = 0 To 3

        Dim ptSCnt(7) As Long
        For j = 0 To 7
            ptSCnt(j) = scopeSCnt(sc, j)
        Next j
        Dim ptTotal As Long
        ptTotal = scopeTotal(sc)

        ' ---- Summary header used on each sub-tab ----
        Dim scopeTitle As String
        If sc = 0 Then
            scopeTitle = "All Types"
        Else
            scopeTitle = scopeLabel(sc) & " Tickets"
        End If

        ' ======================================================
        ' -STATUS tab: Blocked + Assigned To Do
        ' ======================================================
        Dim wsSts As Worksheet
        Set wsSts = ThisWorkbook.Sheets(stsTab(sc))
        Call SetColWidths(wsSts)

        Dim rs As Long : rs = 1
        Call MergeTitle(wsSts, rs, _
             "JIRA STATUS REPORT  -  " & UCase(scopeTitle), "1B3A6B", 16)
        rs = rs + 1
        Call MergeTitle(wsSts, rs, "Generated: " & Format(Now, "mmmm dd, yyyy"), "555555", 10)
        rs = rs + 1
        Call MergeTitle(wsSts, rs, "Total Tickets in scope: " & ptTotal, "1B3A6B", 11)
        rs = rs + 2

        ' Mini summary
        Call SectionHead(wsSts, rs, "SUMMARY", "2E5FA3") : rs = rs + 1
        Call SummaryLine(wsSts, rs, "Total Tickets Tracked: " & ptTotal, False) : rs = rs + 1
        Call SummaryLine(wsSts, rs, "BLOCKED TICKETS (Immediate Attention Required): " & _
             bCntAll(sc), (bCntAll(sc) > 0)) : rs = rs + 1
        Call SummaryLine(wsSts, rs, "ASSIGNED TO DO TICKETS (Not Yet Started): " & _
             dCntAll(sc), False) : rs = rs + 1
        rs = rs + 1

        ' Section: Blocked
        Call SectionHead(wsSts, rs, _
             "BLOCKED TICKETS  (Immediate Attention Required)", "B00000")
        rs = rs + 1
        If bCntAll(sc) = 0 Then
            Call EmptyMsg(wsSts, rs, "  No blocked tickets.") : rs = rs + 2
        Else
            Call WriteAllHeaders(wsSts, rs, "B00000", True) : rs = rs + 1
            Dim kb As Integer
            For kb = 0 To bCntAll(sc) - 1
                Call WriteAllFields(wsSts, wsData, rs, bRowsAll(sc, kb), kb, _
                     RGB(255, 245, 245), RGB(255, 255, 255), True)
                rs = rs + 1
            Next kb
            rs = rs + 1
        End If

        ' Section: Assigned To Do
        Call SectionHead(wsSts, rs, _
             "ASSIGNED TO DO TICKETS  (Not Yet Started)", "9A6A00")
        rs = rs + 1
        If dCntAll(sc) = 0 Then
            Call EmptyMsg(wsSts, rs, "  No assigned To Do tickets.") : rs = rs + 2
        Else
            Call WriteAllHeaders(wsSts, rs, "9A6A00", True) : rs = rs + 1
            Dim kd As Integer
            For kd = 0 To dCntAll(sc) - 1
                Call WriteAllFields(wsSts, wsData, rs, dRowsAll(sc, kd), kd, _
                     RGB(255, 251, 240), RGB(255, 255, 255), True)
                rs = rs + 1
            Next kd
            rs = rs + 1
        End If

        ' ======================================================
        ' -UNASSIGNED tab: Unassigned Open Tickets
        ' ======================================================
        Dim wsUna As Worksheet
        Set wsUna = ThisWorkbook.Sheets(unaTab(sc))
        Call SetColWidths(wsUna)

        Dim ru As Long : ru = 1
        Call MergeTitle(wsUna, ru, _
             "JIRA UNASSIGNED REPORT  -  " & UCase(scopeTitle), "1B3A6B", 16)
        ru = ru + 1
        Call MergeTitle(wsUna, ru, "Generated: " & Format(Now, "mmmm dd, yyyy"), "555555", 10)
        ru = ru + 1
        Call MergeTitle(wsUna, ru, "Total Tickets in scope: " & ptTotal, "1B3A6B", 11)
        ru = ru + 2

        ' Mini summary
        Call SectionHead(wsUna, ru, "SUMMARY", "2E5FA3") : ru = ru + 1
        Call SummaryLine(wsUna, ru, "Total Tickets Tracked: " & ptTotal, False) : ru = ru + 1
        Call SummaryLine(wsUna, ru, _
             "UNASSIGNED OPEN TICKETS (Need an Owner. Excl Done & Canceled): " & _
             uCntAll(sc), (uCntAll(sc) > 0)) : ru = ru + 1
        ru = ru + 1

        ' Section: Unassigned Open
        Call SectionHead(wsUna, ru, _
             "UNASSIGNED OPEN TICKETS  (Need an Owner. Excluded Done & Canceled tickets)", "5D6D7E")
        ru = ru + 1
        If uCntAll(sc) = 0 Then
            Call EmptyMsg(wsUna, ru, "  All open tickets have an assignee.") : ru = ru + 2
        Else
            Call WriteAllHeaders(wsUna, ru, "5D6D7E", True) : ru = ru + 1
            Dim ku As Integer
            For ku = 0 To uCntAll(sc) - 1
                Call WriteAllFields(wsUna, wsData, ru, uRowsAll(sc, ku), ku, _
                     RGB(248, 249, 250), RGB(255, 255, 255), True)
                ru = ru + 1
            Next ku
            ru = ru + 1
        End If

        ' ======================================================
        ' -CLEANUP tab: Assigned Open Tickets With Missing Fields
        ' ======================================================
        Dim wsCln As Worksheet
        Set wsCln = ThisWorkbook.Sheets(clnTab(sc))
        Call SetColWidths(wsCln)
        wsCln.Columns("O").ColumnWidth = 28

        Dim rc As Long : rc = 1
        Call MergeTitle(wsCln, rc, _
             "JIRA CLEANUP REPORT  -  " & UCase(scopeTitle), "1B3A6B", 16)
        rc = rc + 1
        Call MergeTitle(wsCln, rc, "Generated: " & Format(Now, "mmmm dd, yyyy"), "555555", 10)
        rc = rc + 1
        Call MergeTitle(wsCln, rc, "Total Tickets in scope: " & ptTotal, "1B3A6B", 11)
        rc = rc + 2

        ' Mini summary
        Call SectionHead(wsCln, rc, "SUMMARY", "2E5FA3") : rc = rc + 1
        Call SummaryLine(wsCln, rc, "Total Tickets Tracked: " & ptTotal, False) : rc = rc + 1
        Call SummaryLine(wsCln, rc, _
             "ASSIGNED OPEN TICKETS WITH MISSING REQUIRED FIELDS: " & incCntAll(sc), _
             (incCntAll(sc) > 0)) : rc = rc + 1
        rc = rc + 1

        ' Rule note
        wsCln.Range(wsCln.Cells(rc, 1), wsCln.Cells(rc, 15)).Merge
        wsCln.Cells(rc, 1).Value = _
            "  Rule: Assignee set, Status not Done/Canceled, at least one applicable field blank." & _
            "  Story: Desc+Comp+Labels+Priority+ReqDate" & _
            "  |  Sub-task: Desc+Comp+Labels" & _
            "  |  Epic: Desc+Comp+Labels+Priority"
        wsCln.Cells(rc, 1).Font.Name = "Arial"
        wsCln.Cells(rc, 1).Font.Size = 9
        wsCln.Cells(rc, 1).Font.Italic = True
        wsCln.Cells(rc, 1).Font.Color = RGB(123, 63, 0)
        wsCln.Cells(rc, 1).Interior.Color = RGB(255, 243, 224)
        wsCln.Rows(rc).RowHeight = 17 : rc = rc + 1

        ' Section: Missing Fields
        Call SectionHead(wsCln, rc, _
             "ASSIGNED OPEN TICKETS WITH MISSING REQUIRED FIELDS", "7B3F00")
        rc = rc + 1
        If incCntAll(sc) = 0 Then
            Call EmptyMsg(wsCln, rc, "  All assigned tickets have required fields populated.")
            rc = rc + 2
        Else
            Call WriteAllHeaders(wsCln, rc, "7B3F00", True)
            Call TblHdr(wsCln, rc, 15, "Missing Fields", "7B3F00")
            wsCln.Rows(rc).RowHeight = 17 : rc = rc + 1

            Dim ki As Integer
            For ki = 0 To incCntAll(sc) - 1
                Dim sRow As Long : sRow = incRowsAll(sc, ki)
                Dim rTyp As String: rTyp = Trim(wsData.Cells(sRow, 1).Value)
                Dim apPri As Boolean: apPri = (rTyp = "Story" Or rTyp = "Epic")
                Dim apRDt As Boolean: apRDt = (rTyp = "Story")

                Call WriteFieldsWithTypeHighlight(wsCln, wsData, rc, sRow, ki, _
                     RGB(255, 248, 235), RGB(255, 255, 255), apPri, apRDt)

                Dim miss As String : miss = ""
                If Trim(wsData.Cells(sRow, 6).Value)  = "" Then miss = miss & "Description; "
                If Trim(wsData.Cells(sRow, 7).Value)  = "" Then miss = miss & "Component/s; "
                If Trim(wsData.Cells(sRow, 8).Value)  = "" Then miss = miss & "Labels; "
                If apPri And Trim(wsData.Cells(sRow, 9).Value)  = "" Then miss = miss & "Priority; "
                If apRDt And Trim(wsData.Cells(sRow, 10).Value) = "" Then miss = miss & "Req by Date; "
                If Len(miss) > 2 Then miss = Left(miss, Len(miss) - 2)

                wsCln.Cells(rc, 15).Value = miss
                wsCln.Cells(rc, 15).Font.Name = "Arial"
                wsCln.Cells(rc, 15).Font.Size = 9
                wsCln.Cells(rc, 15).Font.Color = RGB(176, 0, 0)
                wsCln.Cells(rc, 15).Font.Bold = True
                If ki Mod 2 = 0 Then
                    wsCln.Cells(rc, 15).Interior.Color = RGB(255, 248, 235)
                Else
                    wsCln.Cells(rc, 15).Interior.Color = RGB(255, 255, 255)
                End If
                wsCln.Cells(rc, 15).BorderAround xlContinuous, xlThin
                wsCln.Rows(rc).RowHeight = 17 : rc = rc + 1
            Next ki
            rc = rc + 1
        End If

        ' ======================================================
        ' -CHARTS tab
        ' ======================================================
        Dim wsCht As Worksheet
        Set wsCht = ThisWorkbook.Sheets(chtTab(sc))
        Call BuildCharts(wsCht, sName, ptSCnt, scopeTitle)

    Next sc












    ' Navigate to Overall-Status as default view
    ThisWorkbook.Sheets("Story-Status").Activate
    ThisWorkbook.Sheets("Story-Status").Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Report generated successfully!" & Chr(10) & Chr(10) & _
           "Tabs created:" & Chr(10) & _
           "  Overall-Status, Overall-Unassigned, Overall-Cleanup" & Chr(10) & _
           "  Story-Status, Story-Unassigned, Story-Cleanup" & Chr(10) & _
           "  SubTask-Status, SubTask-Unassigned, SubTask-Cleanup" & Chr(10) & _
           "  Epic-Status, Epic-Unassigned, Epic-Cleanup" & Chr(10) & Chr(10) & _
           "Total tickets: " & totalTickets & Chr(10) & _
           "Story: " & tCnt(0) & "  Sub-task: " & tCnt(2) & "  Epic: " & tCnt(1), _
           vbInformation, "Jira Weekly Report v8"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro Error"

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ==============================================================
' BUILD CHARTS
' ==============================================================
Private Sub BuildCharts(wsCht As Worksheet, sName() As String, _
                        sCnt() As Long, chartTitle As String)
    Dim j As Integer
    wsCht.Cells(1, 1).Value = "Status" : wsCht.Cells(1, 2).Value = "Count"
    wsCht.Cells(1, 1).Font.Bold = True : wsCht.Cells(1, 2).Font.Bold = True
    For j = 0 To 7
        wsCht.Cells(j + 2, 1).Value = sName(j)
        wsCht.Cells(j + 2, 2).Value = sCnt(j)
    Next j
    wsCht.Columns("A").ColumnWidth = 18 : wsCht.Columns("B").ColumnWidth = 10

    Dim pieObj As ChartObject
    Set pieObj = wsCht.ChartObjects.Add(10, 10, 380, 280)
    With pieObj.Chart
        .ChartType = xlPie
        .SetSourceData wsCht.Range("A1:B9")
        .HasTitle = True
        .ChartTitle.Text = chartTitle & " - Distribution by Status"
        .ChartTitle.Font.Size = 11 : .ChartTitle.Font.Bold = True
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .PlotArea.Interior.Color = RGB(255, 255, 255)
        .ChartArea.Border.LineStyle = xlNone
    End With

    Dim barObj As ChartObject
    Set barObj = wsCht.ChartObjects.Add(400, 10, 380, 280)
    With barObj.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsCht.Range("A1:B9")
        .HasTitle = True
        .ChartTitle.Text = chartTitle & " - Count by Status"
        .ChartTitle.Font.Size = 11 : .ChartTitle.Font.Bold = True
        .HasLegend = False
        .PlotArea.Interior.Color = RGB(248, 250, 255)
        .ChartArea.Border.LineStyle = xlNone
    End With
End Sub

' ==============================================================
' Set standard column widths on a worksheet
' ==============================================================
Private Sub SetColWidths(ws As Worksheet)
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 13
    ws.Columns("I").ColumnWidth = 11
    ws.Columns("J").ColumnWidth = 13
    ws.Columns("K").ColumnWidth = 11
    ws.Columns("L").ColumnWidth = 13
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 38
End Sub

' ==============================================================
' Check if a sheet exists
' ==============================================================
Private Function SheetExists(shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

' ==============================================================
' Write all 13 column headers + optional Jira Link header
' ==============================================================
Private Sub WriteAllHeaders(ws As Worksheet, r As Long, _
                            hexCol As String, addLink As Boolean)
    Dim hdrs(12) As String
    hdrs(0)="Issue Type": hdrs(1)="Key":       hdrs(2)="Summary"
    hdrs(3)="Status":     hdrs(4)="Assignee":  hdrs(5)="Description"
    hdrs(6)="Component/s":hdrs(7)="Labels":    hdrs(8)="Priority"
    hdrs(9)="Req by Date":hdrs(10)="Time Track":hdrs(11)="Created"
    hdrs(12)="Creator"
    Dim c As Integer
    For c = 0 To 12
        Call TblHdr(ws, r, c + 1, hdrs(c), hexCol)
    Next c
    If addLink Then Call TblHdr(ws, r, 14, "Jira Link", hexCol)
    ws.Rows(r).RowHeight = 17
End Sub

' ==============================================================
' Copy all 13 fields + optional clickable Jira Link col 14
' ==============================================================
Private Sub WriteAllFields(wsRpt As Worksheet, wsData As Worksheet, _
                           rptRow As Long, dataRow As Long, _
                           idx As Integer, evenBg As Long, oddBg As Long, _
                           addLink As Boolean)
    Dim bg As Long
    If idx Mod 2 = 0 Then bg = evenBg Else bg = oddBg
    Dim c As Integer
    For c = 1 To 13
        wsRpt.Cells(rptRow, c).Value = wsData.Cells(dataRow, c).Value
        wsRpt.Cells(rptRow, c).Font.Name = "Arial"
        wsRpt.Cells(rptRow, c).Font.Size = 9
        wsRpt.Cells(rptRow, c).Interior.Color = bg
        wsRpt.Cells(rptRow, c).BorderAround xlContinuous, xlThin
        wsRpt.Cells(rptRow, c).VerticalAlignment = xlVAlignCenter
    Next c
    If addLink Then
        Dim keyVal As String
        keyVal = Trim(wsData.Cells(dataRow, 2).Value)
        Dim url As String
        url = "https://jira.schwab.com/browse/" & keyVal
        Dim lc As Range : Set lc = wsRpt.Cells(rptRow, 14)
        wsRpt.Hyperlinks.Add Anchor:=lc, Address:=url, TextToDisplay:=url
        lc.Font.Name = "Arial" : lc.Font.Size = 9
        lc.Font.Color = RGB(0, 70, 180)
        lc.Interior.Color = bg
        lc.BorderAround xlContinuous, xlThin
        lc.VerticalAlignment = xlVAlignCenter
    End If
    wsRpt.Rows(rptRow).RowHeight = 17
End Sub

' ==============================================================
' Write all 13 fields for Cleanup tab - type-aware yellow
' + Jira Link in col 14
' ==============================================================
Private Sub WriteFieldsWithTypeHighlight(wsRpt As Worksheet, _
                                         wsData As Worksheet, _
                                         rptRow As Long, dataRow As Long, _
                                         idx As Integer, _
                                         evenBg As Long, oddBg As Long, _
                                         applyPri As Boolean, applyRDt As Boolean)
    Dim bg As Long
    If idx Mod 2 = 0 Then bg = evenBg Else bg = oddBg
    Dim c As Integer
    For c = 1 To 13
        wsRpt.Cells(rptRow, c).Value = wsData.Cells(dataRow, c).Value
        wsRpt.Cells(rptRow, c).Font.Name = "Arial"
        wsRpt.Cells(rptRow, c).Font.Size = 9
        wsRpt.Cells(rptRow, c).Interior.Color = bg
        wsRpt.Cells(rptRow, c).BorderAround xlContinuous, xlThin
        wsRpt.Cells(rptRow, c).VerticalAlignment = xlVAlignCenter
        If Trim(wsData.Cells(dataRow, c).Value) = "" Then
            Dim doY As Boolean : doY = False
            If c = 6 Then doY = True
            If c = 7 Then doY = True
            If c = 8 Then doY = True
            If c = 9  And applyPri Then doY = True
            If c = 10 And applyRDt Then doY = True
            If doY Then
                wsRpt.Cells(rptRow, c).Interior.Color = RGB(255, 235, 100)
                wsRpt.Cells(rptRow, c).Font.Color = RGB(176, 0, 0)
            End If
        End If
    Next c
    Dim keyVal As String : keyVal = Trim(wsData.Cells(dataRow, 2).Value)
    Dim url As String    : url = "https://jira.schwab.com/browse/" & keyVal
    Dim lc As Range      : Set lc = wsRpt.Cells(rptRow, 14)
    wsRpt.Hyperlinks.Add Anchor:=lc, Address:=url, TextToDisplay:=url
    lc.Font.Name = "Arial" : lc.Font.Size = 9
    lc.Font.Color = RGB(0, 70, 180)
    lc.Interior.Color = bg
    lc.BorderAround xlContinuous, xlThin
    lc.VerticalAlignment = xlVAlignCenter
    wsRpt.Rows(rptRow).RowHeight = 17
End Sub

' ==============================================================
' Merge across 15 cols - centred title
' ==============================================================
Private Sub MergeTitle(ws As Worksheet, r As Long, txt As String, _
                       hexCol As String, sz As Integer)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Merge
    ws.Cells(r, 1).Value = txt
    ws.Cells(r, 1).Font.Name = "Arial" : ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = sz
    ws.Cells(r, 1).Font.Color = HexRGB(hexCol)
    ws.Cells(r, 1).HorizontalAlignment = xlHAlignCenter
    ws.Cells(r, 1).VerticalAlignment = xlVAlignCenter
    ws.Rows(r).RowHeight = sz + 8
End Sub

' ==============================================================
' Coloured section heading across 15 columns
' ==============================================================
Private Sub SectionHead(ws As Worksheet, r As Long, txt As String, hexCol As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Merge
    ws.Cells(r, 1).Value = txt
    ws.Cells(r, 1).Font.Name = "Arial" : ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Interior.Color = HexRGB(hexCol)
    ws.Cells(r, 1).HorizontalAlignment = xlHAlignLeft
    ws.Cells(r, 1).VerticalAlignment = xlVAlignCenter
    ws.Rows(r).RowHeight = 22
End Sub

' ==============================================================
' Sub-heading for Executive Summary type blocks
' ==============================================================
Private Sub SubHead(ws As Worksheet, r As Long, txt As String, hexCol As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Merge
    ws.Cells(r, 1).Value = "  " & txt
    ws.Cells(r, 1).Font.Name = "Arial" : ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 10
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Interior.Color = HexRGB(hexCol)
    ws.Cells(r, 1).HorizontalAlignment = xlHAlignLeft
    ws.Cells(r, 1).VerticalAlignment = xlVAlignCenter
    ws.Rows(r).RowHeight = 18
End Sub

' ==============================================================
' Summary line with optional red alert colouring
' ==============================================================
Private Sub SummaryLine(ws As Worksheet, r As Long, txt As String, isAlert As Boolean)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Merge
    ws.Cells(r, 1).Value = "  " & txt
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Cells(r, 1).Font.Size = 10
    ws.Cells(r, 1).Interior.Color = RGB(248, 250, 255)
    If isAlert Then
        ws.Cells(r, 1).Font.Color = RGB(176, 0, 0)
        ws.Cells(r, 1).Font.Bold = True
    Else
        ws.Cells(r, 1).Font.Color = RGB(0, 0, 0)
    End If
    ws.Rows(r).RowHeight = 17
End Sub

' ==============================================================
' Merged empty-result message
' ==============================================================
Private Sub EmptyMsg(ws As Worksheet, r As Long, txt As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Merge
    ws.Cells(r, 1).Value = txt
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Cells(r, 1).Font.Color = RGB(26, 122, 68)
    ws.Rows(r).RowHeight = 17
End Sub

' ==============================================================
' Single table header cell
' ==============================================================
Private Sub TblHdr(ws As Worksheet, r As Long, col As Integer, _
                   txt As String, hexCol As String)
    ws.Cells(r, col).Value = txt
    ws.Cells(r, col).Font.Name = "Arial" : ws.Cells(r, col).Font.Bold = True
    ws.Cells(r, col).Font.Size = 9
    ws.Cells(r, col).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, col).Interior.Color = HexRGB(hexCol)
    ws.Cells(r, col).HorizontalAlignment = xlHAlignCenter
    ws.Cells(r, col).VerticalAlignment = xlVAlignCenter
    ws.Cells(r, col).BorderAround xlContinuous, xlThin
End Sub

' ==============================================================
' Convert 6-char hex string to Excel RGB Long
' ==============================================================
Private Function HexRGB(h As String) As Long
    HexRGB = RGB(CLng("&H" & Left(h, 2)), _
                 CLng("&H" & Mid(h, 3, 2)), _
                 CLng("&H" & Right(h, 2)))
End Function
