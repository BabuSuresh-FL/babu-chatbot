Attribute VB_Name = "JiraReport"
' ============================================================
' Jira Weekly Report -- VBA Macro
' How to use:
'   1. Press Alt+F11 to open the VBA editor
'   2. Insert > Module
'   3. Paste this entire code and close the editor
'   4. Click the "Generate Report" button on the Instructions sheet
' ============================================================

Sub GenerateReport()

    Dim wsData    As Worksheet
    Dim wsReport  As Worksheet
    Dim wsCharts  As Worksheet
    Dim lastRow   As Long
    Dim i         As Long
    Dim cell      As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsData   = ThisWorkbook.Sheets("Data")
    Set wsReport = ThisWorkbook.Sheets("Report")
    Set wsCharts = ThisWorkbook.Sheets("Charts")

    ' --- Clear previous report (keep row 1 headers) ---
    wsReport.Rows("2:" & wsReport.Rows.Count).ClearContents
    wsReport.Rows("2:" & wsReport.Rows.Count).ClearFormats
    wsCharts.Cells.Clear

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found in the Data sheet. Please import your Jira CSV first.", vbExclamation
        GoTo Cleanup
    End If

    ' ?? Dimensions ????????????????????????????????????????????????????????????
    ' Data columns: A=Issue Type, B=Key, C=Summary, D=Status, E=Assignee, F=Created

    Dim totalTickets  As Long
    totalTickets = lastRow - 1

    ' ?? Count by Status ???????????????????????????????????????????????????????
    Dim statuses(7)  As String
    Dim sCounts(7)   As Long
    statuses(0) = "In Progress"
    statuses(1) = "QA Review"
    statuses(2) = "In Approval"
    statuses(3) = "Blocked"
    statuses(4) = "To Do"
    statuses(5) = "Backlog"
    statuses(6) = "Done"
    statuses(7) = "Canceled"

    Dim j As Integer
    For j = 0 To 7
        sCounts(j) = WorksheetFunction.CountIf(wsData.Range("D2:D" & lastRow), statuses(j))
    Next j

    ' ?? Count by Issue Type ???????????????????????????????????????????????????
    Dim types(2)   As String
    Dim tCounts(2) As Long
    types(0) = "Story"
    types(1) = "Epic"
    types(2) = "Sub-task"
    For j = 0 To 2
        tCounts(j) = WorksheetFunction.CountIf(wsData.Range("A2:A" & lastRow), types(j))
    Next j

    ' ?? Blocked ticket list ???????????????????????????????????????????????????
    Dim blockedKeys()    As String
    Dim blockedSummary() As String
    Dim blockedAssignee() As String
    Dim blockedType()    As String
    Dim blockedCount     As Long
    blockedCount = 0
    ReDim blockedKeys(totalTickets)
    ReDim blockedSummary(totalTickets)
    ReDim blockedAssignee(totalTickets)
    ReDim blockedType(totalTickets)

    For i = 2 To lastRow
        If Trim(wsData.Cells(i, 4).Value) = "Blocked" Then
            blockedKeys(blockedCount)     = wsData.Cells(i, 2).Value
            blockedSummary(blockedCount)  = wsData.Cells(i, 3).Value
            blockedAssignee(blockedCount) = wsData.Cells(i, 5).Value
            blockedType(blockedCount)     = wsData.Cells(i, 1).Value
            blockedCount = blockedCount + 1
        End If
    Next i

    ' ?? To Do ticket list ????????????????????????????????????????????????????
    Dim todoKeys()     As String
    Dim todoSummary()  As String
    Dim todoAssignee() As String
    Dim todoCount      As Long
    todoCount = 0
    ReDim todoKeys(totalTickets)
    ReDim todoSummary(totalTickets)
    ReDim todoAssignee(totalTickets)

    For i = 2 To lastRow
        If Trim(wsData.Cells(i, 4).Value) = "To Do" Then
            todoKeys(todoCount)     = wsData.Cells(i, 2).Value
            todoSummary(todoCount)  = wsData.Cells(i, 3).Value
            todoAssignee(todoCount) = wsData.Cells(i, 5).Value
            todoCount = todoCount + 1
        End If
    Next i

    ' ?? Unassigned open tickets ???????????????????????????????????????????????
    Dim uaKeys()     As String
    Dim uaSummary()  As String
    Dim uaStatus()   As String
    Dim uaCount      As Long
    uaCount = 0
    ReDim uaKeys(totalTickets)
    ReDim uaSummary(totalTickets)
    ReDim uaStatus(totalTickets)

    For i = 2 To lastRow
        Dim st As String
        st = Trim(wsData.Cells(i, 4).Value)
        If (LCase(Trim(wsData.Cells(i, 5).Value)) = "unassigned" Or _
            Trim(wsData.Cells(i, 5).Value) = "") And _
           st <> "Done" And st <> "Canceled" Then
            uaKeys(uaCount)    = wsData.Cells(i, 2).Value
            uaSummary(uaCount) = wsData.Cells(i, 3).Value
            uaStatus(uaCount)  = st
            uaCount = uaCount + 1
        End If
    Next i

    ' ?? Now write the Report sheet ????????????????????????????????????????????
    Dim r As Long
    r = 1

    ' Helper: merge and write a section title
    ' (inline since VBA subs can't be nested cleanly)

    ' === TITLE ================================================================
    With wsReport
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 20
        .Rows.RowHeight = 18
    End With

    Dim rpt As Worksheet
    Set rpt = wsReport

    ' Title block
    Call WriteTitle(rpt, r, "JIRA WEEKLY STATUS REPORT", "1B3A6B", 16)
    r = r + 1
    Call WriteTitle(rpt, r, "Generated: " & Format(Now, "mmmm dd, yyyy  hh:mm AM/PM"), "555555", 10)
    r = r + 1
    Call WriteTitle(rpt, r, "Total Tickets: " & totalTickets, "1B3A6B", 11)
    r = r + 2

    ' === SECTION 1: EXECUTIVE SUMMARY ========================================
    Call WriteSectionHead(rpt, r, "1.  EXECUTIVE SUMMARY", "1B3A6B")
    r = r + 1

    Dim doneCt    As Long: doneCt    = sCounts(6)
    Dim cancelCt  As Long: cancelCt  = sCounts(7)
    Dim blockedCt As Long: blockedCt = sCounts(3)
    Dim activeCt  As Long
    activeCt = sCounts(0) + sCounts(1) + sCounts(2)
    Dim pendingCt As Long
    pendingCt = sCounts(4) + sCounts(5)
    Dim nonCancel As Long: nonCancel = totalTickets - cancelCt
    Dim compRate  As Double
    If nonCancel > 0 Then compRate = doneCt / nonCancel Else compRate = 0

    Dim summaryLines(7) As String
    summaryLines(0) = "Total tickets tracked:  " & totalTickets
    summaryLines(1) = "Active (In Progress + QA + In Approval):  " & activeCt & _
                      "  (" & Format(IIf(totalTickets > 0, activeCt / totalTickets, 0), "0.0%") & ")"
    summaryLines(2) = "Blocked (needs immediate attention):  " & blockedCt
    summaryLines(3) = "To Do + Backlog (not yet started):  " & pendingCt
    summaryLines(4) = "Done:  " & doneCt & _
                      "  (" & Format(IIf(totalTickets > 0, doneCt / totalTickets, 0), "0.0%") & ")"
    summaryLines(5) = "Canceled:  " & cancelCt
    summaryLines(6) = "Completion rate (Done / non-Canceled):  " & Format(compRate, "0.0%")
    summaryLines(7) = "Unassigned open tickets:  " & uaCount

    Dim k As Integer
    For k = 0 To 7
        Dim bc As Range
        Set bc = rpt.Range(rpt.Cells(r, 1), rpt.Cells(r, 7))
        bc.Merge
        rpt.Cells(r, 1).Value = "  " & Chr(149) & "  " & summaryLines(k)
        rpt.Cells(r, 1).Font.Name = "Arial"
        rpt.Cells(r, 1).Font.Size = 10
        rpt.Cells(r, 1).Font.Color = RGB(0, 0, 0)
        If k = 2 And blockedCt > 0 Then
            rpt.Cells(r, 1).Font.Color = RGB(176, 0, 0)
            rpt.Cells(r, 1).Font.Bold = True
        End If
        If k = 7 And uaCount > 0 Then
            rpt.Cells(r, 1).Font.Color = RGB(154, 106, 0)
        End If
        rpt.Cells(r, 1).Interior.Color = RGB(248, 250, 255)
        r = r + 1
    Next k
    r = r + 1

    ' === SECTION 2: STATUS SCORECARD =========================================
    Call WriteSectionHead(rpt, r, "2.  STATUS SCORECARD", "1B3A6B")
    r = r + 1

    ' Header row
    Dim statusBGColors(7)  As String
    Dim statusFGColors(7)  As String
    statusBGColors(0) = "2E5FA3" : statusFGColors(0) = "D6EAF8"  ' In Progress
    statusBGColors(1) = "6C2D8E" : statusFGColors(1) = "E8DAEF"  ' QA Review
    statusBGColors(2) = "1A7A44" : statusFGColors(2) = "D5F5E3"  ' In Approval
    statusBGColors(3) = "B00000" : statusFGColors(3) = "FADBD8"  ' Blocked
    statusBGColors(4) = "9A6A00" : statusFGColors(4) = "FEF9E7"  ' To Do
    statusBGColors(5) = "555555" : statusFGColors(5) = "F2F3F4"  ' Backlog
    statusBGColors(6) = "1A7A44" : statusFGColors(6) = "D5F5E3"  ' Done
    statusBGColors(7) = "999999" : statusFGColors(7) = "FAFAFA"  ' Canceled

    ' Status name row
    For j = 0 To 7
        Dim sc As Range
        Set sc = rpt.Cells(r, j + 1)
        sc.Value = statuses(j)
        sc.Font.Name = "Arial"
        sc.Font.Bold = True
        sc.Font.Size = 9
        sc.Font.Color = RGB(255, 255, 255)
        sc.Interior.Color = HexToLong(statusBGColors(j))
        sc.Alignment.Horizontal = xlHAlignCenter
        sc.Alignment.Vertical = xlVAlignCenter
        sc.BorderAround xlContinuous, xlThin, , HexToLong(statusBGColors(j))
    Next j
    r = r + 1

    ' Count row
    For j = 0 To 7
        Dim cc As Range
        Set cc = rpt.Cells(r, j + 1)
        cc.Value = sCounts(j)
        cc.Font.Name = "Arial"
        cc.Font.Bold = True
        cc.Font.Size = 16
        cc.Font.Color = HexToLong(statusBGColors(j))
        cc.Interior.Color = HexToLong(statusFGColors(j))
        cc.Alignment.Horizontal = xlHAlignCenter
        cc.Alignment.Vertical = xlVAlignCenter
        cc.BorderAround xlContinuous, xlThin, , HexToLong(statusBGColors(j))
        rpt.Rows(r).RowHeight = 30
    Next j
    r = r + 1

    ' Percentage row
    For j = 0 To 7
        Dim pc As Range
        Set pc = rpt.Cells(r, j + 1)
        If totalTickets > 0 Then
            pc.Value = sCounts(j) / totalTickets
        Else
            pc.Value = 0
        End If
        pc.NumberFormat = "0.0%"
        pc.Font.Name = "Arial"
        pc.Font.Size = 9
        pc.Font.Color = HexToLong(statusBGColors(j))
        pc.Interior.Color = HexToLong(statusFGColors(j))
        pc.Alignment.Horizontal = xlHAlignCenter
        pc.BorderAround xlContinuous, xlThin, , HexToLong(statusBGColors(j))
    Next j
    r = r + 2

    ' === SECTION 3: BY ISSUE TYPE ============================================
    Call WriteSectionHead(rpt, r, "3.  BREAKDOWN BY ISSUE TYPE", "1B3A6B")
    r = r + 1

    Dim typeHdrs(4) As String
    typeHdrs(0) = "Issue Type"
    typeHdrs(1) = "Total"
    typeHdrs(2) = "% of All"
    typeHdrs(3) = "Statuses breakdown (see Data sheet)"
    typeHdrs(4) = ""

    Dim th As Range
    ' Header
    For k = 0 To 2
        Set th = rpt.Cells(r, k + 1)
        th.Value = typeHdrs(k)
        th.Font.Name = "Arial": th.Font.Bold = True: th.Font.Size = 9
        th.Font.Color = RGB(255, 255, 255)
        th.Interior.Color = HexToLong("1B3A6B")
        th.Alignment.Horizontal = xlHAlignCenter
        th.Borders.LineStyle = xlContinuous
    Next k
    r = r + 1

    Dim typeColors(2) As String
    typeColors(0) = "D6EAF8"  ' Story
    typeColors(1) = "E8DAEF"  ' Epic
    typeColors(2) = "D5F5E3"  ' Sub-task

    For j = 0 To 2
        Dim pct As Double
        If totalTickets > 0 Then pct = tCounts(j) / totalTickets Else pct = 0
        Dim tr1 As Range
        Set tr1 = rpt.Cells(r, 1)
        tr1.Value = types(j)
        tr1.Font.Name = "Arial": tr1.Font.Bold = True: tr1.Font.Size = 10
        tr1.Interior.Color = HexToLong(typeColors(j))
        tr1.Borders.LineStyle = xlContinuous

        rpt.Cells(r, 2).Value = tCounts(j)
        rpt.Cells(r, 2).Font.Name = "Arial": rpt.Cells(r, 2).Font.Bold = True
        rpt.Cells(r, 2).Alignment.Horizontal = xlHAlignCenter
        rpt.Cells(r, 2).Interior.Color = HexToLong(typeColors(j))
        rpt.Cells(r, 2).Borders.LineStyle = xlContinuous

        rpt.Cells(r, 3).Value = pct
        rpt.Cells(r, 3).NumberFormat = "0.0%"
        rpt.Cells(r, 3).Alignment.Horizontal = xlHAlignCenter
        rpt.Cells(r, 3).Interior.Color = HexToLong(typeColors(j))
        rpt.Cells(r, 3).Borders.LineStyle = xlContinuous
        r = r + 1
    Next j
    r = r + 1

    ' === SECTION 4: BLOCKED TICKETS ==========================================
    Dim blockedHeadColor As String
    blockedHeadColor = "B00000"
    Call WriteSectionHead(rpt, r, "4.  BLOCKED TICKETS  (Immediate Attention Required)", blockedHeadColor)
    r = r + 1

    If blockedCount = 0 Then
        Dim nb As Range
        Set nb = rpt.Range(rpt.Cells(r, 1), rpt.Cells(r, 7))
        nb.Merge
        rpt.Cells(r, 1).Value = "  No blocked tickets this week."
        rpt.Cells(r, 1).Font.Color = RGB(26, 122, 68)
        rpt.Cells(r, 1).Font.Name = "Arial"
        r = r + 2
    Else
        ' Table headers
        Dim bHdrs(3) As String
        bHdrs(0) = "Ticket": bHdrs(1) = "Type"
        bHdrs(2) = "Assignee": bHdrs(3) = "Summary"
        For k = 0 To 3
            Set th = rpt.Cells(r, k + 1)
            th.Value = bHdrs(k)
            th.Font.Name = "Arial": th.Font.Bold = True: th.Font.Size = 9
            th.Font.Color = RGB(255, 255, 255)
            th.Interior.Color = HexToLong(blockedHeadColor)
            th.Alignment.Horizontal = xlHAlignCenter
            th.Borders.LineStyle = xlContinuous
        Next k
        r = r + 1
        For k = 0 To blockedCount - 1
            Dim rowBG As String
            rowBG = IIf(k Mod 2 = 0, "FFF5F5", "FFFFFF")
            rpt.Cells(r, 1).Value = blockedKeys(k)
            rpt.Cells(r, 1).Font.Bold = True
            rpt.Cells(r, 1).Font.Color = RGB(176, 0, 0)
            rpt.Cells(r, 2).Value = blockedType(k)
            rpt.Cells(r, 3).Value = blockedAssignee(k)
            rpt.Cells(r, 4).Value = blockedSummary(k)
            Dim bCol As Integer
            For bCol = 1 To 4
                rpt.Cells(r, bCol).Font.Name = "Arial"
                rpt.Cells(r, bCol).Font.Size = 9
                rpt.Cells(r, bCol).Interior.Color = HexToLong(rowBG)
                rpt.Cells(r, bCol).Borders.LineStyle = xlContinuous
            Next bCol
            r = r + 1
        Next k
        r = r + 1
    End If

    ' === SECTION 5: TO DO TICKETS ============================================
    Call WriteSectionHead(rpt, r, "5.  TO DO TICKETS  (Not Yet Started)", "9A6A00")
    r = r + 1

    If todoCount = 0 Then
        Dim nt As Range
        Set nt = rpt.Range(rpt.Cells(r, 1), rpt.Cells(r, 7))
        nt.Merge
        rpt.Cells(r, 1).Value = "  No To Do tickets at this time."
        rpt.Cells(r, 1).Font.Color = RGB(26, 122, 68)
        rpt.Cells(r, 1).Font.Name = "Arial"
        r = r + 2
    Else
        Dim tHdrs(2) As String
        tHdrs(0) = "Ticket": tHdrs(1) = "Assignee": tHdrs(2) = "Summary"
        For k = 0 To 2
            Set th = rpt.Cells(r, k + 1)
            th.Value = tHdrs(k)
            th.Font.Name = "Arial": th.Font.Bold = True: th.Font.Size = 9
            th.Font.Color = RGB(255, 255, 255)
            th.Interior.Color = HexToLong("9A6A00")
            th.Alignment.Horizontal = xlHAlignCenter
            th.Borders.LineStyle = xlContinuous
        Next k
        r = r + 1
        For k = 0 To todoCount - 1
            rowBG = IIf(k Mod 2 = 0, "FFFBF0", "FFFFFF")
            rpt.Cells(r, 1).Value = todoKeys(k)
            rpt.Cells(r, 1).Font.Bold = True
            rpt.Cells(r, 1).Font.Color = RGB(154, 106, 0)
            rpt.Cells(r, 2).Value = todoAssignee(k)
            rpt.Cells(r, 3).Value = todoSummary(k)
            Dim tCol As Integer
            For tCol = 1 To 3
                rpt.Cells(r, tCol).Font.Name = "Arial"
                rpt.Cells(r, tCol).Font.Size = 9
                rpt.Cells(r, tCol).Interior.Color = HexToLong(rowBG)
                rpt.Cells(r, tCol).Borders.LineStyle = xlContinuous
            Next tCol
            r = r + 1
        Next k
        r = r + 1
    End If

    ' === SECTION 6: UNASSIGNED ===============================================
    Call WriteSectionHead(rpt, r, "6.  UNASSIGNED OPEN TICKETS  (Need an Owner)", "5D6D7E")
    r = r + 1

    If uaCount = 0 Then
        Dim nu As Range
        Set nu = rpt.Range(rpt.Cells(r, 1), rpt.Cells(r, 7))
        nu.Merge
        rpt.Cells(r, 1).Value = "  All open tickets have an assignee."
        rpt.Cells(r, 1).Font.Color = RGB(26, 122, 68)
        rpt.Cells(r, 1).Font.Name = "Arial"
        r = r + 2
    Else
        Dim uHdrs(2) As String
        uHdrs(0) = "Ticket": uHdrs(1) = "Status": uHdrs(2) = "Summary"
        For k = 0 To 2
            Set th = rpt.Cells(r, k + 1)
            th.Value = uHdrs(k)
            th.Font.Name = "Arial": th.Font.Bold = True: th.Font.Size = 9
            th.Font.Color = RGB(255, 255, 255)
            th.Interior.Color = HexToLong("5D6D7E")
            th.Alignment.Horizontal = xlHAlignCenter
            th.Borders.LineStyle = xlContinuous
        Next k
        r = r + 1
        For k = 0 To uaCount - 1
            rowBG = IIf(k Mod 2 = 0, "F8F9FA", "FFFFFF")
            rpt.Cells(r, 1).Value = uaKeys(k)
            rpt.Cells(r, 1).Font.Bold = True
            rpt.Cells(r, 2).Value = uaStatus(k)
            rpt.Cells(r, 3).Value = uaSummary(k)
            Dim uCol As Integer
            For uCol = 1 To 3
                rpt.Cells(r, uCol).Font.Name = "Arial"
                rpt.Cells(r, uCol).Font.Size = 9
                rpt.Cells(r, uCol).Interior.Color = HexToLong(rowBG)
                rpt.Cells(r, uCol).Borders.LineStyle = xlContinuous
            Next uCol
            r = r + 1
        Next k
        r = r + 1
    End If

    ' ?? Charts sheet: Status pie and bar ??????????????????????????????????????
    Dim chtWs As Worksheet
    Set chtWs = wsCharts

    ' Write data for charts in columns A:B
    chtWs.Cells(1, 1).Value = "Status"
    chtWs.Cells(1, 2).Value = "Count"
    For j = 0 To 7
        chtWs.Cells(j + 2, 1).Value = statuses(j)
        chtWs.Cells(j + 2, 2).Value = sCounts(j)
    Next j
    chtWs.Columns("A").ColumnWidth = 18
    chtWs.Columns("B").ColumnWidth = 10

    ' Pie chart
    Dim pie As ChartObject
    Set pie = chtWs.ChartObjects.Add(Left:=10, Top:=10, Width:=400, Height:=280)
    With pie.Chart
        .ChartType = xlPie
        .SetSourceData chtWs.Range("A1:B9")
        .HasTitle = True
        .ChartTitle.Text = "Ticket Distribution by Status"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .ApplyLayout 6
        .PlotArea.Interior.Color = RGB(255, 255, 255)
        .ChartArea.Border.LineStyle = xlNone
    End With

    ' Bar chart
    Dim bar As ChartObject
    Set bar = chtWs.ChartObjects.Add(Left:=420, Top:=10, Width:=400, Height:=280)
    With bar.Chart
        .ChartType = xlBarClustered
        .SetSourceData chtWs.Range("A1:B9")
        .HasTitle = True
        .ChartTitle.Text = "Ticket Count by Status"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .HasLegend = False
        .Axes(xlValue).HasTitle = False
        .PlotArea.Interior.Color = RGB(248, 250, 255)
        .ChartArea.Border.LineStyle = xlNone
    End With

    ' ?? Final polish ??????????????????????????????????????????????????????????
    wsReport.Activate
    wsReport.Range("A1").Select
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Report generated successfully!  " & totalTickets & " tickets processed.", vbInformation, "Jira Weekly Report"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' ?? Helper: write a merged title row ?????????????????????????????????????????
Private Sub WriteTitle(ws As Worksheet, r As Long, txt As String, _
                       hexColor As String, sz As Integer)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
    rng.Merge
    ws.Cells(r, 1).Value = txt
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = sz
    ws.Cells(r, 1).Font.Color = HexToLong(hexColor)
    ws.Cells(r, 1).Alignment.Horizontal = xlHAlignCenter
    ws.Cells(r, 1).Alignment.Vertical = xlVAlignCenter
    ws.Rows(r).RowHeight = sz + 10
End Sub

' ?? Helper: write a section heading row ??????????????????????????????????????
Private Sub WriteSectionHead(ws As Worksheet, r As Long, txt As String, hexColor As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
    rng.Merge
    ws.Cells(r, 1).Value = txt
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Interior.Color = HexToLong(hexColor)
    ws.Cells(r, 1).Alignment.Horizontal = xlHAlignLeft
    ws.Cells(r, 1).Alignment.Vertical = xlVAlignCenter
    ws.Rows(r).RowHeight = 24
    Dim bdr As Border
    Set bdr = ws.Cells(r, 1).Borders(xlEdgeBottom)
    bdr.LineStyle = xlContinuous
    bdr.Weight = xlMedium
    bdr.Color = HexToLong(hexColor)
End Sub

' ?? Helper: convert hex color string to Long ?????????????????????????????????
Private Function HexToLong(hexStr As String) As Long
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Left(hexStr, 2))
    g = CLng("&H" & Mid(hexStr, 3, 2))
    b = CLng("&H" & Right(hexStr, 2))
    HexToLong = RGB(r, g, b)
End Function