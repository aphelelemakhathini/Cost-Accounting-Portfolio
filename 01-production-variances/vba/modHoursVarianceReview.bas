Attribute VB_Name = "modHoursVarianceReview"
'===============================================================================
' modHoursVarianceReview: planned vs confirmed hours, per operation (Excel VBA)
'
' Stage 2 of the variance process. Paste the operations export (production
' order operations: order, activity, work centre, order type, op qty, setup and
' machine standards and confirmations, dates, material, plant) into "Sheet1"
' and run one of:
'   BuildHoursSummaryWeekly   - review window = last 14 working days (Mon-Sat)
'   BuildHoursSummaryMonthly  - review window = last 30 calendar days
'
' Builds: Comments (one row per order, over-plan lines in red), one tab per
' department, Favourable Variances, a hidden HOURS DATA sheet where every
' number is a formula pointing back at the export row, and an HOURS OVERVIEW
' with a plant slicer, live Top 5 / Top 10 tables (LET/FILTER/SORTBY) and charts.
' Remarks typed into the output sheets survive a re-run.
'===============================================================================
Option Explicit

' =========================================================
' SETTINGS
' A line turns RED when EITHER condition is met:
'   - the difference (confirmed vs planned) is at least the hour threshold
'   - the difference is at least PCT_THRESHOLD of planned hours
' =========================================================
Private Const MACH_HRS_THRESHOLD As Double = 2      ' machine: 2 hrs
Private Const SETUP_HRS_THRESHOLD As Double = 1     ' setup:   1 hr
Private Const PCT_THRESHOLD As Double = 0.5         ' both:    50%

' Review period for the pop-ups / overview, set by the entry macro:
'   weekly  = the last 14 working days (Monday-Saturday, Sundays skipped)
'   monthly = the last 30 calendar days
' ending on the latest date in the export
' that has real operation data (GI / GR / INSP lines are ignored and the
' operation must have planned or confirmed hours).
' Dates are Actual Op. Finish (Actual Start Date if finish is blank).
Private WINDOW_DAYS As Long          ' set by the entry macro
Private USE_CALENDAR As Boolean      ' True = calendar days, False = working days (Mon-Sat)

Private Const OVERVIEW_NAME As String = "HOURS OVERVIEW"
Private Const SUMMARY_NAME As String = "Comments"
Private Const DATA_NAME As String = "HOURS DATA"
' The SAP export sheet (found by name now, because the overview is the first sheet)
Private Const SOURCE_NAME As String = "Sheet1"
' Sheet name for the better-than-planned operations (change it here if you want another name)
Private Const GOOD_NAME As String = "Favourable Variances"
' Favourable Variances: a part is highlighted when it is under plan by MORE than GOOD_HRS_HILITE hours
Private Const GOOD_HRS_HILITE As Double = 2
Private Const PICK_NAME As String = "PLANT FILTER"

' Flag dots on the All machines table (|average difference per op|)
Private Const SETUP_RED_FROM As Double = 0.5
Private Const SETUP_ORANGE_FROM As Double = 0.1      ' below this = yellow
Private Const MACH_RED_FROM As Double = 3
Private Const MACH_ORANGE_FROM As Double = 2
Private Const MACH_YELLOW_FROM As Double = 1         ' below this = no dot

' Department tabs: sheet name | work centers that feed it (Like patterns)
Private Const DEPT_WELD As String = "Welding|WLD*;BASEPLATE*"
Private Const DEPT_PAINT As String = "Painting|PNT*"
Private Const DEPT_MACH As String = "Machining & Spares|MCH*"
Private Const DEPT_ASSY As String = "Assembly & Mounting|ASM*"
Private Const DEPT_OTHER As String = "Other Work Centers|*"   ' whatever is left over

' Special departments: when the plant is SPECIAL_PLANT and the material is one of these,
' the operation goes to the sheet with the name after the = sign (instead of a work-center department)
Private Const SPECIAL_PLANT As String = "1300"   ' service workshop
Private Const SPECIAL_MATERIALS As String = "SV_REPAIR=Repair Customer Pump;SV_CLAIM=Claim Repair;SV_REWIND=Rewind Motor;SV_COMM=Commissioning"

' Orders are listed newest first (by the latest operation date on the order)
Private Const SORT_NEWEST_FIRST As Boolean = True

' ---------------------------------------------------------
' Operations in the review period (filled during the main loop)
' ---------------------------------------------------------
Private wOrder() As String, wItem() As String, wOp() As String, wWC() As String, wType() As String
Private wDate() As Date
Private wPM() As Double, wCM() As Double, wPS() As Double, wCS() As Double
Private wN As Long
Private wHd() As Long        ' HOURS DATA row of each review-period operation

' ---------------------------------------------------------
' Every done operation (feeds the HOURS DATA sheet)
' ---------------------------------------------------------
Private hSrc() As Long, hOrder() As String, hItem() As String, hOp() As String
Private hN As Long
Private hPlant() As String
Private gSrcName As String
Private gFormatNotes As String

' ---------------------------------------------------------
' Every operation with a variance (for the department tabs)
' ---------------------------------------------------------
Private dPlant() As String, dOrder() As String, dItem() As String
Private dOp() As String, dWC() As String
Private dMach() As String, dSetup() As String
Private dStartD() As Date, dFinishD() As Date
Private dMRed() As Boolean, dSRed() As Boolean
Private dDate() As Date
Private dDept() As String
Private dN As Long

' ---------------------------------------------------------
' Operations that finished BETTER than planned (over the same thresholds):
' they live on the Good Variances sheet only
' ---------------------------------------------------------
Private gPlant() As String, gOrder() As String, gItem() As String
Private gOp() As String, gWC() As String
Private gMach() As String, gSetup() As String
Private gStartD() As Date, gFinishD() As Date
Private gDate() As Date
Private gMHi() As Boolean, gSHi() As Boolean
Private gN As Long

' ---------------------------------------------------------
' Remarks typed by users, kept when the sheets are rebuilt
' ---------------------------------------------------------
Private rkKey() As String, rkVal() As String
Private rkN As Long

Sub BuildHoursSummaryWeekly()
    WINDOW_DAYS = 14
    USE_CALENDAR = False
    RunHoursSummary
End Sub

Sub BuildHoursSummaryMonthly()
    WINDOW_DAYS = 30
    USE_CALENDAR = True
    RunHoursSummary
End Sub

Private Sub RunHoursSummary()

    Dim wsSrc As Worksheet, wsOut As Worksheet, ws As Worksheet
    Dim data As Variant
    Dim lastRow As Long, i As Long, n As Long, r As Long, c As Long

    Dim outArr() As Variant
    Dim mFlags() As String, sFlags() As String

    Dim curOrder As String, thisOrder As String, rawOrder As String
    Dim curPlant As String, curItem As String
    Dim wcVal As String, workctr As String, activity As String
    Dim opLabel As String, machineText As String, setupText As String
    Dim opLines As String, startLines As String, finishLines As String
    Dim machineLines As String, setupLines As String
    Dim mFlag As String, sFlag As String
    Dim plannedSetup As Double, confirmedSetup As Double
    Dim plannedMachine As Double, confirmedMachine As Double
    Dim opDate As Date, maxDate As Date
    Dim winStart As Date, winEnd As Date
    Dim isDone As Boolean, canRed As Boolean
    Dim mRedNow As Boolean, sRedNow As Boolean
    Dim mGoodNow As Boolean, sGoodNow As Boolean
    Dim goodMach As String, goodSetup As String, thisM As String, thisS As String
    Dim ordDate() As Date, ordMax As Date

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo CleanUp

    ' -------------------------------------------------------
    ' 1. Read the source sheet (Sheet1) into memory
    ' -------------------------------------------------------
    Set wsSrc = GetSourceSheet()
    If wsSrc Is Nothing Then
        MsgBox "The export sheet '" & SOURCE_NAME & "' was not found.", vbExclamation
        GoTo CleanUp
    End If
    gSrcName = wsSrc.Name
    wsSrc.Visible = xlSheetVisible      ' must be visible while the other sheets are rebuilt
    lastRow = wsSrc.Cells(wsSrc.rows.Count, 3).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found on " & wsSrc.Name, vbExclamation
        GoTo CleanUp
    End If
    data = wsSrc.Range("A1:W" & lastRow).value

    ' Work out the review period: latest date of an operation that was done
    ' (confirmed setup + machine > 0) on an included work center,
    ' then count back WINDOW_DAYS days (working or calendar, inclusive)
    maxDate = 0
    For i = 2 To lastRow
        wcVal = UCase(Trim(CStr(data(i, 3))))
        If Not IsExcludedWC(wcVal) Then
            If NumVal(data(i, 14)) + NumVal(data(i, 17)) > 0 Then          ' done = something confirmed
                opDate = OpDateOf(data(i, 20), data(i, 12))
                If opDate > maxDate Then maxDate = opDate
            End If
        End If
    Next i
    If maxDate = 0 Then
        MsgBox "No dated operations found on " & wsSrc.Name, vbExclamation
        GoTo CleanUp
    End If
    winEnd = maxDate
    If USE_CALENDAR Then
        winStart = winEnd - WINDOW_DAYS + 1
    Else
        winStart = StartOfWorkingWindow(winEnd, WINDOW_DAYS)
    End If

    ReDim outArr(1 To lastRow, 1 To 9)
    ReDim mFlags(1 To lastRow)
    ReDim sFlags(1 To lastRow)
    ReDim ordDate(1 To lastRow)

    dN = 0
    ReDim dPlant(1 To lastRow): ReDim dOrder(1 To lastRow): ReDim dItem(1 To lastRow)
    ReDim dOp(1 To lastRow): ReDim dWC(1 To lastRow)
    ReDim dMach(1 To lastRow): ReDim dSetup(1 To lastRow)
    ReDim dStartD(1 To lastRow): ReDim dFinishD(1 To lastRow)
    ReDim dMRed(1 To lastRow): ReDim dSRed(1 To lastRow): ReDim dDate(1 To lastRow)
    ReDim dDept(1 To lastRow)

    wN = 0
    ReDim wOrder(1 To lastRow): ReDim wItem(1 To lastRow)
    ReDim wOp(1 To lastRow): ReDim wWC(1 To lastRow): ReDim wDate(1 To lastRow)
    ReDim wType(1 To lastRow)
    ReDim wPM(1 To lastRow): ReDim wCM(1 To lastRow)
    ReDim wPS(1 To lastRow): ReDim wCS(1 To lastRow)

    hN = 0
    ReDim hSrc(1 To lastRow): ReDim hOrder(1 To lastRow)
    ReDim hItem(1 To lastRow): ReDim hOp(1 To lastRow): ReDim hPlant(1 To lastRow)
    ReDim wHd(1 To lastRow)

    gN = 0
    ReDim gPlant(1 To lastRow): ReDim gOrder(1 To lastRow): ReDim gItem(1 To lastRow)
    ReDim gOp(1 To lastRow): ReDim gWC(1 To lastRow)
    ReDim gMach(1 To lastRow): ReDim gSetup(1 To lastRow)
    ReDim gStartD(1 To lastRow): ReDim gFinishD(1 To lastRow): ReDim gDate(1 To lastRow)
    ReDim gMHi(1 To lastRow): ReDim gSHi(1 To lastRow)

    ' -------------------------------------------------------
    ' 2. Walk the data, build one output row per order
    '    A Order | B Activity | C Work Center | H Op Qty
    '    L Act Start | M Setup std | N Setup conf
    '    P Machine std/unit | Q Machine conf
    '    R Material | T Act Finish | V Plant
    ' -------------------------------------------------------
    curOrder = ""
    n = 0

    For i = 2 To lastRow + 1

        If i <= lastRow Then
            rawOrder = Trim(CStr(data(i, 1)))
            If rawOrder <> "" Then thisOrder = rawOrder Else thisOrder = curOrder
        Else
            thisOrder = "##END##"
        End If

        ' Order changed -> write the finished order into the output array
        If thisOrder <> curOrder And curOrder <> "" Then
            n = n + 1
            outArr(n, 1) = AsNum(curPlant)
            outArr(n, 2) = AsNum(curOrder)
            outArr(n, 3) = AsNum(curItem)
            If opLines <> "" Then
                outArr(n, 4) = opLines
                outArr(n, 5) = machineLines
                outArr(n, 6) = setupLines
                outArr(n, 8) = startLines
                outArr(n, 9) = finishLines
            End If
            mFlags(n) = mFlag
            sFlags(n) = sFlag
            ordDate(n) = ordMax
            ordMax = 0

            opLines = "": startLines = "": finishLines = ""
            machineLines = "": setupLines = ""
            mFlag = "": sFlag = ""
            curPlant = "": curItem = ""
        End If

        If i > lastRow Then Exit For

        curOrder = thisOrder
        If Trim(CStr(data(i, 22))) <> "" Then curPlant = Trim(CStr(data(i, 22)))
        If Trim(CStr(data(i, 18))) <> "" Then curItem = Trim(CStr(data(i, 18)))

        ' Skip GI* / GR* / *INSP* work centers
        wcVal = UCase(Trim(CStr(data(i, 3))))
        If IsExcludedWC(wcVal) Then
            GoTo NextRow
        End If

        activity = Trim(CStr(data(i, 2)))
        workctr = Trim(CStr(data(i, 3)))
        plannedSetup = Round(NumVal(data(i, 13)), 4)
        confirmedSetup = Round(NumVal(data(i, 14)), 4)
        plannedMachine = Round(NumVal(data(i, 8)) * NumVal(data(i, 16)), 4)   ' Op qty x std value 2
        confirmedMachine = Round(NumVal(data(i, 17)), 4)
        opLabel = "OP" & CStr(CLng(val(activity))) & "(" & workctr & ")"

        ' Operation is "done" only if some time was confirmed
        isDone = (confirmedSetup + confirmedMachine > 0)
        ' No red when both dates are blank (shown as "-") or the op isn't done
        canRed = isDone And Not (DateStr(data(i, 12)) = "-" And DateStr(data(i, 20)) = "-")

        ' Every DONE operation goes to HOURS DATA
        If isDone Then
            hN = hN + 1
            hSrc(hN) = i: hOrder(hN) = curOrder: hItem(hN) = curItem: hPlant(hN) = curPlant
            hOp(hN) = "OP" & CStr(CLng(val(activity)))
        End If

        ' Record DONE operations that fall in the review period
        opDate = OpDateOf(data(i, 20), data(i, 12))
        If isDone And opDate >= winStart And opDate <= winEnd And opDate > 0 Then
            wN = wN + 1
            wHd(wN) = hN + 1
            wOrder(wN) = curOrder: wItem(wN) = curItem
            wType(wN) = Trim(CStr(data(i, 4)))
            wOp(wN) = "OP" & CStr(CLng(val(activity))): wWC(wN) = workctr
            wDate(wN) = opDate
            wPM(wN) = plannedMachine: wCM(wN) = confirmedMachine
            wPS(wN) = plannedSetup: wCS(wN) = confirmedSetup
        End If

        ' Skip the operation if BOTH machine and setup match
        If Abs(plannedMachine - confirmedMachine) < 0.005 And _
           Abs(plannedSetup - confirmedSetup) < 0.005 Then
            GoTo NextRow
        End If

        mRedNow = False: sRedNow = False
        mGoodNow = False: sGoodNow = False
        machineText = "-": setupText = "-"
        goodMach = "-": goodSetup = "-"

        ' Machine: over the thresholds AND confirmed below planned = "good" -> Good Variances sheet only
        If Abs(plannedMachine - confirmedMachine) >= 0.005 Then
            thisM = "Planned " & FormatHrs(plannedMachine) & " ; Confirmed " & FormatHrs(confirmedMachine)
            If canRed And IsRed(plannedMachine, confirmedMachine, MACH_HRS_THRESHOLD) Then
                If confirmedMachine < plannedMachine Then
                    mGoodNow = True
                    goodMach = thisM
                Else
                    mRedNow = True
                    machineText = thisM
                End If
            Else
                machineText = thisM
            End If
        End If

        ' Setup: same rule
        If Abs(plannedSetup - confirmedSetup) >= 0.005 Then
            thisS = "Planned " & FormatHrs(plannedSetup) & " ; Confirmed " & FormatHrs(confirmedSetup)
            If canRed And IsRed(plannedSetup, confirmedSetup, SETUP_HRS_THRESHOLD) Then
                If confirmedSetup < plannedSetup Then
                    sGoodNow = True
                    goodSetup = thisS
                Else
                    sRedNow = True
                    setupText = thisS
                End If
            Else
                setupText = thisS
            End If
        End If

        ' Good variances go to their own sheet
        If mGoodNow Or sGoodNow Then
            gN = gN + 1
            gPlant(gN) = curPlant: gOrder(gN) = curOrder: gItem(gN) = curItem
            gOp(gN) = "OP" & CStr(CLng(val(activity))): gWC(gN) = workctr
            gMach(gN) = goodMach: gSetup(gN) = goodSetup
            gStartD(gN) = ToDate(data(i, 12)): gFinishD(gN) = ToDate(data(i, 20))
            gDate(gN) = opDate
            ' highlight: under plan by more than GOOD_HRS_HILITE hours
            gMHi(gN) = mGoodNow And ((plannedMachine - confirmedMachine) > GOOD_HRS_HILITE)
            gSHi(gN) = sGoodNow And ((plannedSetup - confirmedSetup) > GOOD_HRS_HILITE)
        End If

        ' Nothing left for the Comments sheet / department tabs
        If machineText = "-" And setupText = "-" Then
            GoTo NextRow
        End If

        mFlag = mFlag & IIf(mRedNow, "1", "0")
        sFlag = sFlag & IIf(sRedNow, "1", "0")

        ' Keep the operation for the department tabs
        dN = dN + 1
        dPlant(dN) = curPlant: dOrder(dN) = curOrder: dItem(dN) = curItem
        dOp(dN) = "OP" & CStr(CLng(val(activity))): dWC(dN) = workctr
        dMach(dN) = machineText: dSetup(dN) = setupText
        dStartD(dN) = ToDate(data(i, 12)): dFinishD(dN) = ToDate(data(i, 20))
        dMRed(dN) = mRedNow: dSRed(dN) = sRedNow
        dDate(dN) = opDate
        If opDate > ordMax Then ordMax = opDate

        If opLines = "" Then
            opLines = opLabel
            startLines = DateStr(data(i, 12))
            finishLines = DateStr(data(i, 20))
            machineLines = machineText
            setupLines = setupText
        Else
            opLines = opLines & vbLf & opLabel
            startLines = startLines & vbLf & DateStr(data(i, 12))
            finishLines = finishLines & vbLf & DateStr(data(i, 20))
            machineLines = machineLines & vbLf & machineText
            setupLines = setupLines & vbLf & setupText
        End If

NextRow:
    Next i

    ' -------------------------------------------------------
    ' 3. Keep any Remarks already typed, then rebuild the sheets
    ' -------------------------------------------------------
    HarvestRemarks

    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If IsOutputSheet(ws.Name) Then
            ws.Visible = xlSheetVisible
            ws.Delete
        End If
    Next i

    ' Order in which the orders are listed: newest first
    Dim ordIdx() As Long, ordKey() As Double, k As Long
    If n > 0 Then
        ReDim ordIdx(1 To n): ReDim ordKey(1 To n)
        For k = 1 To n
            ordIdx(k) = k
            ordKey(k) = IIf(SORT_NEWEST_FIRST, CDbl(ordDate(k)), -CDbl(ordDate(k)))
        Next k
        SortDesc ordKey, ordIdx, n
    End If

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = SUMMARY_NAME
    wsOut.Cells.Font.Name = "Aptos Narrow"
    wsOut.Cells.Font.Size = 11

    HeaderBlock wsOut

    ' -------------------------------------------------------
    ' 4. Write the Comments sheet (one row per order)
    ' -------------------------------------------------------
    If n > 0 Then
        Dim rngData As Range, finalArr() As Variant, src As Long
        Set rngData = wsOut.Range("A3").Resize(n, 9)
        wsOut.Range("D3").Resize(n, 6).NumberFormat = "@"

        ReDim finalArr(1 To n, 1 To 9)
        For r = 1 To n
            src = ordIdx(r)
            For c = 1 To 9
                finalArr(r, c) = outArr(src, c)
            Next c
            finalArr(r, 7) = GetRemark(RemarkKey(CStr(outArr(src, 2)), "", ""))
        Next r
        rngData.value = finalArr

        With rngData
            .VerticalAlignment = xlBottom
            .WrapText = True
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(191, 191, 191)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(191, 191, 191)
        End With

        wsOut.Range("A3").Resize(n, 3).HorizontalAlignment = xlRight
        wsOut.Range("D3").Resize(n, 4).HorizontalAlignment = xlLeft
        wsOut.Range("H3").Resize(n, 2).HorizontalAlignment = xlCenter

        With Union(wsOut.Range("D3").Resize(n, 3), wsOut.Range("H3").Resize(n, 2)).Font
            .Italic = True
            .Color = RGB(128, 128, 128)
        End With

        For r = 1 To n
            src = ordIdx(r)
            If InStr(mFlags(src), "1") > 0 Then ColorLines wsOut.Cells(r + 2, 5), CStr(outArr(src, 5)), mFlags(src)
            If InStr(sFlags(src), "1") > 0 Then ColorLines wsOut.Cells(r + 2, 6), CStr(outArr(src, 6)), sFlags(src)
        Next r
    End If

    SetWidths wsOut
    If n > 0 Then wsOut.Range("A3").Resize(n, 9).rows.AutoFit
    ApplyFilter wsOut, 9, n + 2
    FreezeTop wsOut

    ' -------------------------------------------------------
    ' 4b. Department tabs
    ' -------------------------------------------------------
    ' Department of every operation: the special materials of plant 1300 first, else by work center.
    ' Departments without any operation get no sheet.
    Dim dd As Long, sp As Variant
    For dd = 1 To dN
        dDept(dd) = DeptOf(dPlant(dd), dItem(dd), dWC(dd))
    Next dd
    BuildDeptSheet deptName(DEPT_WELD)
    BuildDeptSheet deptName(DEPT_PAINT)
    BuildDeptSheet deptName(DEPT_MACH)
    BuildDeptSheet deptName(DEPT_ASSY)
    BuildDeptSheet deptName(DEPT_OTHER)
    For Each sp In Split(SPECIAL_MATERIALS, ";")
        BuildDeptSheet Trim(Split(CStr(sp), "=")(1))
    Next sp
    BuildGoodSheet

    ' -------------------------------------------------------
    ' 5. Hidden data sheet + hidden overview sheet + pop-ups
    ' -------------------------------------------------------
    Dim wsOv As Worksheet, wsHd As Worksheet, wsPk As Worksheet
    Set wsPk = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsPk.Name = PICK_NAME
    BuildPlantPick wsPk

    Set wsHd = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsHd.Name = DATA_NAME
    BuildDataSheet wsHd, wsSrc.Name

    Set wsOv = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOv.Name = OVERVIEW_NAME
    BuildOverview wsOv, winStart, winEnd

    ' Helper sheets and the export stay out of sight; the overview becomes the first sheet
    wsHd.Visible = xlSheetHidden
    wsPk.Visible = xlSheetHidden
    wsSrc.Visible = xlSheetHidden
    wsOv.Visible = xlSheetVisible
    wsOv.Move Before:=ThisWorkbook.Worksheets(1)

    ' With the screen on, so the charts and slicer land on the real positions
    Application.ScreenUpdating = True
    wsOv.Activate
    DoEvents
    PlaceOverviewCharts wsOv
    ActiveWindow.Zoom = 90
    wsOv.Range("A1").Select

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ShowPopups winStart, winEnd

    MsgBox "Done! '" & SUMMARY_NAME & "' created with " & n & " orders, " & _
           "plus the department tabs." & vbLf & _
           "'" & GOOD_NAME & "' lists " & gN & " operation(s) that finished better than planned." & vbLf & vbLf & _
           "The first sheet, '" & OVERVIEW_NAME & "', is the overview.", vbInformation, "Processing Complete"
    Exit Sub

CleanUp:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical

End Sub

' =========================================================
' Unhide / hide the overview sheet (and its data sheet)
' =========================================================
Sub ShowOverview()
    On Error Resume Next
    ThisWorkbook.Worksheets(DATA_NAME).Visible = xlSheetVisible
    With ThisWorkbook.Worksheets(OVERVIEW_NAME)
        If Err.Number <> 0 Then
            MsgBox "Run BuildHoursSummaryWeekly or BuildHoursSummaryMonthly first.", vbExclamation
            Exit Sub
        End If
        .Visible = xlSheetVisible
        .Activate
    End With
    PlaceOverviewCharts ThisWorkbook.Worksheets(OVERVIEW_NAME)
    ActiveWindow.Zoom = 90
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub

Sub HideOverview()
    On Error Resume Next
    ThisWorkbook.Worksheets(OVERVIEW_NAME).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(DATA_NAME).Visible = xlSheetHidden
End Sub

' =========================================================
' POP-UPS
'   1. Top 5 machines - worst average SETUP difference
'   2. Top 5 machines - worst average MACHINE difference
'   3. Worst single operations (with order number + order type)
'   4. Setup and machine differences by order type
'
' MsgBox is plain text, so structure comes from headings,
' divider lines and indented bullet lines.
' Characters used are from the standard Windows set
' (Chr(149) = bullet, Chr(150) = en dash, Chr(151) = em dash).
' =========================================================
Private Sub ShowPopups(winStart As Date, winEnd As Date)
    Dim hdr As String
    Dim keys() As String, nOps() As Long, nOrd() As Long
    Dim pm() As Double, cm() As Double, ps() As Double, cs() As Double, m As Long

    If wN = 0 Then
        MsgBox "No operations with an Actual Op. Finish in this period." & vbLf & vbLf & _
               "Period:  " & Format(winStart, "dd\/mm\/yyyy") & " " & Chr(150) & " " & Format(winEnd, "dd\/mm\/yyyy"), _
               vbInformation, "Hours variance review"
        Exit Sub
    End If

    ' Shared header block
    hdr = "Period:" & vbTab & Format(winStart, "dd\/mm\/yyyy") & "  " & Chr(150) & "  " & Format(winEnd, "dd\/mm\/yyyy") & _
          "   (" & WindowLabel() & ")" & vbLf & _
          "Scope:" & vbTab & CountDistinctOrders() & " orders  " & Chr(149) & "  " & wN & " operations"

    GroupBy False, keys, nOps, nOrd, pm, cm, ps, cs, m
    MsgBox Clip(TopFiveText(hdr, "SETUP", keys, nOps, nOrd, ps, cs, m)), vbExclamation, _
           "1 of 4  " & Chr(151) & "  Top 5 machines: setup"
    MsgBox Clip(TopFiveText(hdr, "MACHINE", keys, nOps, nOrd, pm, cm, m)), vbExclamation, _
           "2 of 4  " & Chr(151) & "  Top 5 machines: machine time"

    MsgBox Clip(WorstOpsText(hdr)), vbExclamation, _
           "3 of 4  " & Chr(151) & "  Worst single operations"

    GroupBy True, keys, nOps, nOrd, pm, cm, ps, cs, m
    MsgBox Clip(OrderTypeText(hdr, keys, nOps, nOrd, pm, cm, ps, cs, m)), vbInformation, _
           "4 of 4  " & Chr(151) & "  By order type"
End Sub

' Title + divider + header + divider
Private Function Banner(title As String, hdr As String) As String
    Banner = title & vbLf & Rule() & vbLf & hdr & vbLf & Rule() & vbLf
End Function

Private Function Rule() As String
    Rule = String(38, Chr(151))
End Function

Private Function Bul() As String
    Bul = "      " & Chr(149) & "  "
End Function

' MsgBox shows ~1024 characters at most
Private Function Clip(t As String) As String
    If Len(t) > 1000 Then
        Clip = Left(t, 960) & vbLf & "..." & vbLf & "(see the HOURS OVERVIEW sheet for the rest)"
    Else
        Clip = t
    End If
End Function

Private Function OverUnder(d As Double) As String
    If d > 0 Then OverUnder = "OVER plan" Else OverUnder = "UNDER plan"
End Function

Private Function TopFiveText(hdr As String, label As String, keys() As String, nOps() As Long, nOrd() As Long, _
                             pl() As Double, co() As Double, m As Long) As String
    Dim ord() As Long, avg() As Double, k As Long, j As Long, t As String, shown As Long
    Dim a As Double
    ReDim ord(1 To m): ReDim avg(1 To m)
    For k = 1 To m
        ord(k) = k
        avg(k) = Abs(co(k) - pl(k)) / nOps(k)
    Next k
    SortDesc avg, ord, m

    t = Banner("TOP 5 MACHINES  " & Chr(151) & "  WORST AVERAGE " & label & " DIFFERENCE", hdr)

    For k = 1 To m
        If shown >= 5 Or avg(k) < 0.005 Then Exit For
        j = ord(k)
        shown = shown + 1
        a = (co(j) - pl(j)) / nOps(j)
        t = t & vbLf & "#" & shown & "   " & keys(j) & "      " & SignedHrs(a) & " per operation  (" & OverUnder(a) & ")" & vbLf & _
            Bul() & nOps(j) & " operation(s)  " & Chr(149) & "  Planned " & FormatHrs(pl(j)) & _
            "  " & Chr(149) & "  Confirmed " & FormatHrs(co(j)) & vbLf
    Next k
    If shown = 0 Then t = t & vbLf & "No " & LCase(label) & " differences in this period." & vbLf

    t = t & vbLf & Rule() & vbLf & "Average = total difference / number of operations." & vbLf & "Ranked by size (over or under plan)."
    TopFiveText = t
End Function

Private Function WorstOpsText(hdr As String) As String
    Dim k As Long, bS As Long, bM As Long, t As String
    bS = 1: bM = 1
    For k = 2 To wN
        If Abs(wCS(k) - wPS(k)) > Abs(wCS(bS) - wPS(bS)) Then bS = k
        If Abs(wCM(k) - wPM(k)) > Abs(wCM(bM) - wPM(bM)) Then bM = k
    Next k

    t = Banner("WORST SINGLE OPERATIONS", hdr)

    t = t & vbLf & "SETUP" & vbLf
    If Abs(wCS(bS) - wPS(bS)) < 0.005 Then
        t = t & Bul() & "No setup differences in this period." & vbLf
    Else
        t = t & OpText(bS, False)
    End If

    t = t & vbLf & "MACHINE" & vbLf
    If Abs(wCM(bM) - wPM(bM)) < 0.005 Then
        t = t & Bul() & "No machine differences in this period." & vbLf
    Else
        t = t & OpText(bM, True)
    End If
    WorstOpsText = t
End Function

Private Function OpText(k As Long, isMachine As Boolean) As String
    Dim p As Double, c As Double, d As Double, pct As String
    If isMachine Then p = wPM(k): c = wCM(k) Else p = wPS(k): c = wCS(k)
    d = c - p
    If p <> 0 Then pct = "  (" & Format(Abs(d) / p, "0%") & ")" Else pct = ""
    OpText = Bul() & "Order:" & vbTab & wOrder(k) & "   " & Chr(150) & "   type " & IIf(wType(k) = "", "(blank)", wType(k)) & vbLf & _
             Bul() & "Op:" & vbTab & wOp(k) & " on " & wWC(k) & vbLf & _
             Bul() & "Item:" & vbTab & wItem(k) & "   " & Chr(150) & "   finished " & Format(wDate(k), "dd\/mm\/yyyy") & vbLf & _
             Bul() & "Hours:" & vbTab & "Planned " & FormatHrs(p) & "  " & Chr(149) & "  Confirmed " & FormatHrs(c) & vbLf & _
             Bul() & "Diff:" & vbTab & SignedHrs(d) & pct & "  " & OverUnder(d) & vbLf
End Function

Private Function OrderTypeText(hdr As String, keys() As String, nOps() As Long, nOrd() As Long, _
                               pm() As Double, cm() As Double, ps() As Double, cs() As Double, m As Long) As String
    Dim k As Long, t As String
    t = Banner("SETUP AND MACHINE DIFFERENCES BY ORDER TYPE", hdr)
    For k = 1 To m
        t = t & vbLf & keys(k) & "      " & nOrd(k) & " orders  " & Chr(149) & "  " & nOps(k) & " ops" & vbLf & _
            Bul() & "Setup:" & vbTab & SignedHrs(cs(k) - ps(k)) & "   (avg " & SignedHrs((cs(k) - ps(k)) / nOps(k)) & " per op)" & vbLf & _
            "            " & "Planned " & FormatHrs(ps(k)) & "  " & Chr(149) & "  Confirmed " & FormatHrs(cs(k)) & vbLf & _
            Bul() & "Machine:" & vbTab & SignedHrs(cm(k) - pm(k)) & "   (avg " & SignedHrs((cm(k) - pm(k)) / nOps(k)) & " per op)" & vbLf & _
            "            " & "Planned " & FormatHrs(pm(k)) & "  " & Chr(149) & "  Confirmed " & FormatHrs(cm(k)) & vbLf
    Next k
    OrderTypeText = t
End Function

' Group the period's operations by machine (byType = False) or order type (True)
Private Sub GroupBy(byType As Boolean, ByRef keys() As String, ByRef nOps() As Long, ByRef nOrd() As Long, _
                    ByRef pm() As Double, ByRef cm() As Double, ByRef ps() As Double, ByRef cs() As Double, _
                    ByRef m As Long)
    Dim seen() As String, k As Long, j As Long, found As Long, key As String
    ReDim keys(1 To wN): ReDim nOps(1 To wN): ReDim nOrd(1 To wN): ReDim seen(1 To wN)
    ReDim pm(1 To wN): ReDim cm(1 To wN): ReDim ps(1 To wN): ReDim cs(1 To wN)
    m = 0
    For k = 1 To wN
        If byType Then key = wType(k) Else key = wWC(k)
        If key = "" Then key = "(blank)"
        found = 0
        For j = 1 To m
            If keys(j) = key Then found = j: Exit For
        Next j
        If found = 0 Then
            m = m + 1: found = m
            keys(m) = key: seen(m) = "|"
        End If
        nOps(found) = nOps(found) + 1
        If InStr(seen(found), "|" & wOrder(k) & "|") = 0 Then
            seen(found) = seen(found) & wOrder(k) & "|"
            nOrd(found) = nOrd(found) + 1
        End If
        pm(found) = pm(found) + wPM(k): cm(found) = cm(found) + wCM(k)
        ps(found) = ps(found) + wPS(k): cs(found) = cs(found) + wCS(k)
    Next k
End Sub

Private Function CountDistinctOrders() As Long
    Dim k As Long, seen As String, cnt As Long
    seen = "|"
    For k = 1 To wN
        If InStr(seen, "|" & wOrder(k) & "|") = 0 Then
            seen = seen & wOrder(k) & "|"
            cnt = cnt + 1
        End If
    Next k
    CountDistinctOrders = cnt
End Function

' =========================================================
' PLANT SLICER ON THE OVERVIEW
' PLANT FILTER (hidden sheet) holds a small table: Both plants / each plant.
' The slicer on the overview filters that table; SUBTOTAL(103) tells
' which row is still showing, and B1 (named hv_Plant) turns that into
' "All" or the plant number. HOURS DATA's In Period column reads hv_Plant,
' so every overview formula follows the slicer.
' The Top 5 / Top 10 / order type / machine lists are live formulas
' (UNIQUE / FILTER / SORTBY), so they re-rank by themselves - no event needed.
' =========================================================
Private Sub BuildPlantPick(ws As Worksheet)
    Dim lst As String, k As Long, p As Variant, n As Long, lo As ListObject

    lst = "|"
    For k = 1 To hN
        If hPlant(k) <> "" And InStr(lst, "|" & hPlant(k) & "|") = 0 Then lst = lst & hPlant(k) & "|"
    Next k

    ws.Cells.Font.Name = "Aptos Narrow"
    ws.Cells.Font.Size = 11
    ws.Range("A1").value = "Selected plant"
    ws.Range("A2").value = "Overview built for"
    ws.Range("A1:A3").Font.Bold = True
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").value = "All"

    ws.Range("D:D").NumberFormat = "@"
    ws.Range("D1").value = "Plant"
    ws.Range("E1").value = "Shown"
    ws.Range("D2").value = "Both plants"
    n = 1
    If Len(lst) > 1 Then
        For Each p In Split(Mid(lst, 2, Len(lst) - 2), "|")
            n = n + 1
            ws.Cells(n + 1, 4).value = CStr(p)
        Next p
    End If

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("D1").Resize(n + 1, 2), , xlYes)
    lo.Name = "tblPlantPick"
    lo.TableStyle = "TableStyleLight9"
    lo.ListColumns("Shown").DataBodyRange.Formula = "=SUBTOTAL(103,[@Plant])"

    AddName "hv_Plant", ws, "$B$1"
    ws.Range("B1").Formula = "=IF(SUM(tblPlantPick[Shown])=1," & _
        "IF(INDEX(tblPlantPick[Plant],MATCH(1,tblPlantPick[Shown],0))=""Both plants"",""All""," & _
        "INDEX(tblPlantPick[Plant],MATCH(1,tblPlantPick[Shown],0))&""""),""All"")"

    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("D:E").ColumnWidth = 14
End Sub

' Slicer on the overview, once
Private Sub AddPlantSlicer(ws As Worksheet)
    Dim shp As Shape, lo As ListObject, sc As SlicerCache, sl As Slicer, i As Long

    For Each shp In ws.Shapes
        If shp.Type = msoSlicer Then Exit Sub
    Next shp

    ' clear slicer caches left behind by earlier runs
    On Error Resume Next
    For i = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        If ThisWorkbook.SlicerCaches(i).Slicers.Count = 0 Then ThisWorkbook.SlicerCaches(i).Delete
    Next i
    On Error GoTo NoSlicer

    Set lo = ThisWorkbook.Worksheets(PICK_NAME).ListObjects("tblPlantPick")
    Set sc = ThisWorkbook.SlicerCaches.Add2(lo, "Plant")
    Set sl = sc.Slicers.Add(ws, , , "Select plant", _
                            ws.rows(6).Top + 3, ws.Columns(3).Left + 4, 342, ws.rows(6).Height - 6)
    sl.NumberOfColumns = 3
    sl.DisplayHeader = True
    sl.RowHeight = 24
    sl.ColumnWidth = 106

    On Error Resume Next
    sl.Style = SlicerStyleName()
    If Err.Number <> 0 Then Err.Clear: sl.Style = "SlicerStyleDark1"
    On Error GoTo 0
    Exit Sub

NoSlicer:
    MsgBox "The plant slicer could not be added:" & vbLf & Err.Number & ": " & Err.Description & vbLf & vbLf & _
           "Everything else was built; the overview shows both plants.", vbExclamation
End Sub

' Corporate slicer style: blue selected button, light grey others, blue hover
Private Function SlicerStyleName() As String
    Const nm As String = "Hours Overview Plant"
    Dim ts As TableStyle

    On Error Resume Next
    ThisWorkbook.TableStyles(nm).Delete
    Err.Clear
    Set ts = ThisWorkbook.TableStyles.Add(nm)
    If ts Is Nothing Then SlicerStyleName = "SlicerStyleDark1": Exit Function

    ts.ShowAsAvailableSlicerStyle = True
    ts.ShowAsAvailableTableStyle = False
    ts.ShowAsAvailablePivotTableStyle = False

    StyleElement ts.TableStyleElements(xlWholeTable), RGB(255, 255, 255), RGB(64, 64, 64), False, RGB(191, 191, 191)
    StyleElement ts.TableStyleElements(xlHeaderRow), RGB(255, 255, 255), RGB(0, 112, 192), True, -1
    With ts.TableStyleElements(xlHeaderRow).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(0, 112, 192): .Weight = xlMedium
    End With

    StyleElement ts.TableStyleElements(xlSlicerSelectedItemWithData), RGB(0, 112, 192), RGB(255, 255, 255), True, RGB(0, 112, 192)
    StyleElement ts.TableStyleElements(xlSlicerHoveredSelectedItemWithData), RGB(0, 84, 147), RGB(255, 255, 255), True, RGB(0, 84, 147)
    StyleElement ts.TableStyleElements(xlSlicerUnselectedItemWithData), RGB(242, 242, 242), RGB(64, 64, 64), False, RGB(217, 217, 217)
    StyleElement ts.TableStyleElements(xlSlicerHoveredUnselectedItemWithData), RGB(221, 235, 247), RGB(0, 112, 192), True, RGB(0, 112, 192)
    StyleElement ts.TableStyleElements(xlSlicerSelectedItemWithNoData), RGB(166, 166, 166), RGB(255, 255, 255), False, RGB(166, 166, 166)
    StyleElement ts.TableStyleElements(xlSlicerUnselectedItemWithNoData), RGB(248, 248, 248), RGB(166, 166, 166), False, RGB(230, 230, 230)

    SlicerStyleName = nm
End Function

' fill, font colour, bold, border colour (-1 = no border) for one slicer style element
Private Sub StyleElement(el As TableStyleElement, fillClr As Long, fontClr As Long, isBold As Boolean, borderClr As Long)
    Dim b As Variant
    On Error Resume Next
    el.Interior.Color = fillClr
    el.Font.Color = fontClr
    el.Font.Bold = isBold
    If borderClr >= 0 Then
        For Each b In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            el.Borders(b).LineStyle = xlContinuous
            el.Borders(b).Color = borderClr
            el.Borders(b).Weight = xlThin
        Next b
    End If
End Sub

' Data rows of the 10 biggest |confirmed - planned| in the period/plant
Private Function TopTenRowsFormula(pl As String, co As String) As String
    TopTenRowsFormula = "=IFERROR(LET(dd,ABS(" & co & "-" & pl & ")," & _
        "ok,hd_In*(dd>=0.005)," & _
        "k,FILTER(SEQUENCE(ROWS(hd_In)),ok),kd,FILTER(dd,ok)," & _
        "s,SORTBY(k,kd,-1)," & _
        "INDEX(s,SEQUENCE(MIN(10,ROWS(s))))+1),"""")"
End Function

' 1 when that operation is over the hour limit or the % limit, else 0
Private Function OverLimitFormula(rowsRef As String, pl As String, co As String, thr As Double) As String
    OverLimitFormula = "=IFERROR(LET(r," & rowsRef & "-1," & _
        "p,INDEX(" & pl & ",r),c,INDEX(" & co & ",r),dd,ABS(c-p)," & _
        "--(((dd>=" & Num(thr) & ")+(p=0)+(IFERROR(dd/p,0)>=" & Num(PCT_THRESHOLD) & "))>0)),"""")"
End Function

' =========================================================
' HOURS DATA: one row per done operation, every number a formula
' pointing at its row on the SAP export
' =========================================================
Private Sub BuildDataSheet(ws As Worksheet, srcName As String)
    Dim arr() As Variant, k As Long, r As Long, s As String, q As String, lastR As Long

    q = "'" & Replace(srcName, "'", "''") & "'!"
    lastR = hN + 1

    ws.Cells.Font.Name = "Aptos Narrow"
    ws.Cells.Font.Size = 11
    ws.Range("A1:S1").value = Array("Src Row", "Order", "Order Type", "Item", "Op", "Work Center", _
        "Op Date", "Planned Setup", "Confirmed Setup", "Op Qty", "Std / Unit", "Planned Machine", _
        "Confirmed Machine", "Setup Diff", "Machine Diff", "In Period", "First in Period", "First for Type", "Plant")
    StyleHeader ws.Range("A1:S1"), RGB(221, 235, 247)

    ' Review period, as formulas
    ws.Range("T1").value = IIf(USE_CALENDAR, "Calendar days", "Working days (Mon-Sat)")
    ws.Range("T2").value = "Period end"
    ws.Range("T3").value = "Period start"
    ws.Range("T1:T3").Font.Bold = True
    ws.Range("U1").value = WINDOW_DAYS
    ws.Range("U2").Formula = "=MAX(G2:G" & lastR & ")"
    If USE_CALENDAR Then
        ws.Range("U3").Formula = "=U2-U1+1"
    Else
        ws.Range("U3").Formula = "=WORKDAY.INTL(U2+1,-U1,""0000001"")"
    End If
    ws.Range("T4").value = "Plant filter (from the slicer)"
    ws.Range("T4").Font.Bold = True
    ws.Range("U4").Formula = "=hv_Plant"
    ws.Range("U2:U3").NumberFormat = "dd/mm/yyyy"

    ' Flag thresholds (the conditional formatting reads these cells)
    ws.Range("T5").value = "Setup red from"
    ws.Range("T6").value = "Setup orange from"
    ws.Range("T7").value = "Machine red from"
    ws.Range("T8").value = "Machine orange from"
    ws.Range("T9").value = "Machine yellow from"
    ws.Range("T5:T9").Font.Bold = True
    ws.Range("U5").value = SETUP_RED_FROM
    ws.Range("U6").value = SETUP_ORANGE_FROM
    ws.Range("U7").value = MACH_RED_FROM
    ws.Range("U8").value = MACH_ORANGE_FROM
    ws.Range("U9").value = MACH_YELLOW_FROM
    AddName "th_SetupRed", ws, "$U$5"
    AddName "th_SetupOrange", ws, "$U$6"
    AddName "th_MachRed", ws, "$U$7"
    AddName "th_MachOrange", ws, "$U$8"
    AddName "th_MachYellow", ws, "$U$9"

    AddName "hv_Days", ws, "$U$1"
    AddName "hv_End", ws, "$U$2"
    AddName "hv_Start", ws, "$U$3"
    AddName "hd_Plant", ws, "$S$2:$S$" & lastR
    AddName "hd_Order", ws, "$B$2:$B$" & lastR
    AddName "hd_Type", ws, "$C$2:$C$" & lastR
    AddName "hd_WC", ws, "$F$2:$F$" & lastR
    AddName "hd_Date", ws, "$G$2:$G$" & lastR
    AddName "hd_PlanSetup", ws, "$H$2:$H$" & lastR
    AddName "hd_ConfSetup", ws, "$I$2:$I$" & lastR
    AddName "hd_PlanMach", ws, "$L$2:$L$" & lastR
    AddName "hd_ConfMach", ws, "$M$2:$M$" & lastR
    AddName "hd_In", ws, "$P$2:$P$" & lastR
    AddName "hd_FirstOrd", ws, "$Q$2:$Q$" & lastR
    AddName "hd_FirstType", ws, "$R$2:$R$" & lastR
    AddName "hd_FirstMach", ws, "$V$2:$V$" & lastR

    If hN = 0 Then Exit Sub

    ws.Range("B2").Resize(hN, 1).NumberFormat = "@"
    ws.Range("D2").Resize(hN, 1).NumberFormat = "@"
    ws.Range("S2").Resize(hN, 1).NumberFormat = "@"

    ReDim arr(1 To hN, 1 To 19)
    For k = 1 To hN
        r = k + 1: s = CStr(hSrc(k))
        arr(k, 1) = "=HYPERLINK(""#" & q & "A" & s & """," & s & ")"
        arr(k, 2) = hOrder(k)                                   ' carried down from the order's first row
        arr(k, 3) = "=TRIM(" & q & "D" & s & "&"""")"
        arr(k, 4) = hItem(k)                                    ' carried down
        arr(k, 5) = hOp(k)
        arr(k, 6) = "=TRIM(" & q & "C" & s & "&"""")"
        arr(k, 7) = "=IF(N(" & q & "T" & s & ")>0,INT(" & q & "T" & s & ")," & _
                    "IF(N(" & q & "L" & s & ")>0,INT(" & q & "L" & s & "),0))"
        arr(k, 8) = "=N(" & q & "M" & s & ")"
        arr(k, 9) = "=N(" & q & "N" & s & ")"
        arr(k, 10) = "=N(" & q & "H" & s & ")"
        arr(k, 11) = "=N(" & q & "P" & s & ")"
        arr(k, 12) = "=J" & r & "*K" & r
        arr(k, 13) = "=N(" & q & "Q" & s & ")"
        arr(k, 14) = "=I" & r & "-H" & r
        arr(k, 15) = "=M" & r & "-L" & r
        arr(k, 16) = "=AND(G" & r & ">0,G" & r & ">=hv_Start,G" & r & "<=hv_End," & _
                     "OR(hv_Plant=""All"",S" & r & "&""""=hv_Plant&""""))"
        arr(k, 17) = "=IF(P" & r & ",--(COUNTIFS(B$2:B" & r & ",B" & r & ",P$2:P" & r & ",TRUE)=1),0)"
        arr(k, 18) = "=IF(P" & r & ",--(COUNTIFS(B$2:B" & r & ",B" & r & ",C$2:C" & r & ",C" & r & _
                     ",P$2:P" & r & ",TRUE)=1),0)"
        arr(k, 19) = hPlant(k)                                  ' carried down
    Next k
    ws.Range("A2").Resize(hN, 19).Formula = arr

    ' First row of each order per machine (counts distinct orders per machine on the overview)
    Dim arrV() As Variant
    ReDim arrV(1 To hN, 1 To 1)
    For k = 1 To hN
        r = k + 1
        arrV(k, 1) = "=IF(P" & r & ",--(COUNTIFS(B$2:B" & r & ",B" & r & ",F$2:F" & r & ",F" & r & _
                     ",P$2:P" & r & ",TRUE)=1),0)"
    Next k
    ws.Range("V2").Resize(hN, 1).Formula = arrV

    ws.Range("G2").Resize(hN, 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("H2").Resize(hN, 6).NumberFormat = "0.00"
    ws.Range("N2").Resize(hN, 2).NumberFormat = "+0.00;[Red]-0.00;0.00"
    ws.Columns("A:S").ColumnWidth = 14
    ws.Columns("T").ColumnWidth = 24
    ws.Columns("U").ColumnWidth = 12
    ws.Range("A1").Resize(lastR, 19).AutoFilter

    ' Top 10 rankings for the overview (follow In Period, so they follow the plant slicer)
    ws.Range("V1").value = "First for Machine"
    ws.Range("W1").value = "Top 10 setup (data row)"
    ws.Range("X1").value = "Top 10 machine (data row)"
    ws.Range("Y1").value = "Setup over limit"
    ws.Range("Z1").value = "Machine over limit"
    StyleHeader ws.Range("V1:Z1"), RGB(221, 235, 247)
    ws.Range("W2").Formula2 = TopTenRowsFormula("hd_PlanSetup", "hd_ConfSetup")
    ws.Range("X2").Formula2 = TopTenRowsFormula("hd_PlanMach", "hd_ConfMach")
    ws.Range("Y2").Formula2 = OverLimitFormula("W2#", "hd_PlanSetup", "hd_ConfSetup", SETUP_HRS_THRESHOLD)
    ws.Range("Z2").Formula2 = OverLimitFormula("X2#", "hd_PlanMach", "hd_ConfMach", MACH_HRS_THRESHOLD)
    ws.Columns("V:Z").ColumnWidth = 16
    ws.Range("W:Z").HorizontalAlignment = xlCenter
    ' order type totals for the chart (live)
    ws.Range("AB1").value = "Chart: order type"
    ws.Range("AC1").value = "Chart: setup diff"
    ws.Range("AD1").value = "Chart: machine diff"
    StyleHeader ws.Range("AB1:AD1"), RGB(221, 235, 247)
    ws.Range("AB2").Formula2 = "=IFERROR(LET(t,UNIQUE(FILTER(hd_Type,hd_In)),IF(t="""",""(blank)"",t)),"""")"
    ws.Range("AC2").Formula2 = "=IFERROR(LET(t,UNIQUE(FILTER(hd_Type,hd_In))," & _
        "SUMIFS(hd_ConfSetup,hd_Type,t,hd_In,TRUE)-SUMIFS(hd_PlanSetup,hd_Type,t,hd_In,TRUE)),"""")"
    ws.Range("AD2").Formula2 = "=IFERROR(LET(t,UNIQUE(FILTER(hd_Type,hd_In))," & _
        "SUMIFS(hd_ConfMach,hd_Type,t,hd_In,TRUE)-SUMIFS(hd_PlanMach,hd_Type,t,hd_In,TRUE)),"""")"
    ' order type totals for the chart: the chart ranges grow/shrink with the number of order types
    Dim otCount As String
    otCount = "MAX(1,COUNTIF('" & ws.Name & "'!$AB$2:$AB$40,""?*""))"
    AddName "chOT_Cat", ws, "$AB$2:INDEX('" & ws.Name & "'!$AB$2:$AB$40," & otCount & ")"
    AddName "chOT_Setup", ws, "$AC$2:INDEX('" & ws.Name & "'!$AC$2:$AC$40," & otCount & ")"
    AddName "chOT_Mach", ws, "$AD$2:INDEX('" & ws.Name & "'!$AD$2:$AD$40," & otCount & ")"
    ws.Columns("AB:AD").ColumnWidth = 18
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub AddName(nm As String, ws As Worksheet, addr As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & ws.Name & "'!" & addr
End Sub

' SUMIFS criteria: the key cell, or "" for (blank)
Private Function Crit(key As String, cellAddr As String) As String
    If key = "(blank)" Then Crit = """""" Else Crit = cellAddr
End Function

' Coloured dot driven by the cell's value (the ABS of the average):
' red from redNm, orange from orangeNm, yellow from yellowNm
' (yellowNm = "" means everything below orange is yellow)
Private Sub AddFlagDots(rng As Range, redNm As String, orangeNm As String, yellowNm As String)
    Dim clrRed As Long, clrOrange As Long, clrYellow As Long
    clrRed = RGB(255, 0, 0)
    clrOrange = RGB(255, 120, 0)
    clrYellow = RGB(255, 204, 0)

    rng.FormatConditions.Delete
    rng.NumberFormat = """" & ChrW(9679) & """"     ' show a dot instead of the number
    rng.Font.Size = 14
    rng.HorizontalAlignment = xlCenter

    If yellowNm = "" Then
        rng.Font.Color = clrYellow
    Else
        AddBand rng, yellowNm, clrYellow
    End If
    AddBand rng, orangeNm, clrOrange
    AddBand rng, redNm, clrRed                      ' added last = top priority
End Sub

Private Sub AddBand(rng As Range, thresholdName As String, clr As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, _
                                      Formula1:="=" & thresholdName)
    fc.Font.Color = clr
    fc.SetFirstPriority
End Sub

' =========================================================
' HIDDEN OVERVIEW SHEET (every figure is a formula on HOURS DATA)
' =========================================================
Private Sub BuildOverview(ws As Worksheet, winStart As Date, winEnd As Date)
    Dim k As Long, j As Long, r As Long, cr As String, a As String
    Dim keys() As String, nOps() As Long, nOrd() As Long
    Dim pm() As Double, cm() As Double, ps() As Double, cs() As Double, m As Long
    Dim tk() As String, tOps() As Long, tOrd() As Long
    Dim tpm() As Double, tcm() As Double, tps() As Double, tcs() As Double, tm As Long
    Dim ord() As Long, sk() As Double
    Dim hdrColor As Long, bannerColor As Long
    Dim s5Top As Long, s5Bot As Long, m5Top As Long, m5Bot As Long
    Dim otTop As Long, otBot As Long, amTop As Long, amBot As Long, z1 As Long, z2 As Long

    hdrColor = RGB(221, 235, 247)
    bannerColor = RGB(0, 112, 192)

    ws.Cells.Font.Name = "Aptos Narrow"
    ws.Cells.Font.Size = 11
    ActiveWindow.DisplayGridlines = False

    With ws.Range("A1:K1")
        .Merge
        .value = "Hours Variance Overview"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = bannerColor
        .HorizontalAlignment = xlLeft: .IndentLevel = 1
    End With

    ' Period + scope, as formulas on HOURS DATA
    ws.Range("A2").value = "Period start"
    ws.Range("A3").value = "Period end"
    ws.Range("A4").value = "Orders"
    ws.Range("A5").value = "Operations"
    ws.Range("A2:A5").Font.Bold = True
    ws.Range("B2").Formula = "=hv_Start"
    ws.Range("B3").Formula = "=hv_End"
    ws.Range("B4").Formula = "=SUM(hd_FirstOrd)"
    ws.Range("B5").Formula = "=COUNTIF(hd_In,TRUE)"
    ws.Range("A6").value = "Plant"
    ws.Range("A6").Font.Bold = True
    ws.Range("B6").Formula = "=IF(hv_Plant=""All"",""Both plants"",""Plant ""&hv_Plant)"
    ws.Range("B6").Font.Bold = True
    ws.Range("B6").Interior.Color = RGB(255, 242, 204)
    ws.rows(6).RowHeight = 62
    ws.Range("A6:B6").VerticalAlignment = xlCenter
    AddPlantSlicer ws
    ws.Range("B2:B3").NumberFormat = "dd/mm/yyyy"
    ws.Range("B2:B6").HorizontalAlignment = xlLeft
    ws.Range("D2").value = "Last " & WindowLabel() & ", by Actual Op. Finish; " & _
                           "done operations only; excludes GI* / GR* / *INSP*"
    ws.Range("D3").value = "Every figure is calculated from the '" & DATA_NAME & "' sheet, " & _
                           "which links to each row of " & gSrcName & "."
    ws.Range("D4").value = "Generated: " & Format(Now, "dd\/mm\/yyyy hh:nn")
    ws.Range("D2:D4").Font.Italic = True
    ws.Range("D2:D4").Font.Color = RGB(89, 89, 89)

    If wN = 0 Then
        ws.Range("A8").value = "No operations finished in this period."
        ws.Range("A8").Font.Bold = True
        Exit Sub
    End If

    GroupBy False, keys, nOps, nOrd, pm, cm, ps, cs, m
    GroupBy True, tk, tOps, tOrd, tpm, tcm, tps, tcs, tm

    ' ---- Top 5 tables (live: re-rank with the plant slicer) ----
    r = 8
    r = TopFiveTable(ws, r, "Top 5 machines - worst average SETUP difference", _
                     hdrColor, bannerColor, s5Top, s5Bot, "hd_PlanSetup", "hd_ConfSetup")
    r = r + 2
    r = TopFiveTable(ws, r, "Top 5 machines - worst average MACHINE difference", _
                     hdrColor, bannerColor, m5Top, m5Bot, "hd_PlanMach", "hd_ConfMach")

    ' ---- Order type, All machines, Top 10 setup, Top 10 machine ----
    ' One live block that grows/shrinks with the plant slicer, so there are never gaps
    r = r + 2
    z1 = r
    z2 = r + (2 + tm) + 1 + (2 + m) + 1 + 12 + 1 + 12 + 2
    ws.Cells(z1, 1).Formula2 = LowerBlockFormula()
    ws.Columns(1).ColumnWidth = 29
    ws.Columns("B:K").ColumnWidth = 19
    ws.Columns("L:M").Hidden = True
    FormatLowerBlock ws, z1, z2

    ' ---- Charts: three compact charts stacked beside the Top 5 tables (columns H:L) ----
    ' PlaceOverviewCharts sizes and positions them from the real row/column sizes.
    ' Row heights: airy Top 5 rows, so the three charts have room
    ws.rows(9).RowHeight = 30
    ws.rows(17).RowHeight = 30
    ws.rows("10:14").RowHeight = 28
    ws.rows("18:22").RowHeight = 28
    ws.Range("A10:G14").VerticalAlignment = xlCenter
    ws.Range("A18:G22").VerticalAlignment = xlCenter

    Dim ch As Chart
    Dim hd As Worksheet
    Dim fbRows As Long
    Set hd = ThisWorkbook.Worksheets(DATA_NAME)
    fbRows = tm + 1
    If fbRows < 3 Then fbRows = 3

    Set ch = NewChart(ws, 20, 20, 400, 120, xlBarClustered, "Top 5 machines - total setup difference (hrs)")
    ch.Parent.Name = "chtTopSetup"
    AddSeries ch, "Total setup diff", ws.Range(ws.Cells(s5Top, 6), ws.Cells(s5Bot, 6)), _
              ws.Range(ws.Cells(s5Top, 2), ws.Cells(s5Bot, 2)), RGB(237, 125, 49)
    ch.HasLegend = False
    StyleChart ch

    Set ch = NewChart(ws, 20, 20, 400, 120, xlBarClustered, "Top 5 machines - total machine difference (hrs)")
    ch.Parent.Name = "chtTopMachine"
    AddSeries ch, "Total machine diff", ws.Range(ws.Cells(m5Top, 6), ws.Cells(m5Bot, 6)), _
              ws.Range(ws.Cells(m5Top, 2), ws.Cells(m5Bot, 2)), bannerColor
    ch.HasLegend = False
    StyleChart ch

    ' order type chart: only the order types that exist for the selected plant (live named ranges)
    Set ch = NewChart(ws, 20, 20, 400, 120, xlBarClustered, "Total difference by order type (hrs)")
    ch.Parent.Name = "chtOrderType"
    AddNamedSeries ch, "Setup diff", "chOT_Cat", "chOT_Setup", RGB(237, 125, 49), hd.Range("AB2:AB" & fbRows), hd.Range("AC2:AC" & fbRows)
    AddNamedSeries ch, "Machine diff", "chOT_Cat", "chOT_Mach", bannerColor, hd.Range("AB2:AB" & fbRows), hd.Range("AD2:AD" & fbRows)
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight
    ch.Legend.Font.Size = 8
    StyleChart ch

    PlaceOverviewCharts ws
End Sub

' Wrap a formula so the row stays blank when its key cell is empty
Private Function IfKey(keyAddr As String, expr As String) As String
    IfKey = "=IF(" & keyAddr & "="""",""""," & expr & ")"
End Function

' SUMIFS/COUNTIFS criteria for a key cell: "(blank)" means an empty value
Private Function KeyCrit(keyAddr As String) As String
    KeyCrit = "IF(" & keyAddr & "=""(blank)"",""""," & keyAddr & ")"
End Function

' Number for a formula, always with a "." decimal (VBA's Str does that)
Private Function Num(x As Double) As String
    Num = Trim(Str(x))
End Function

' Top-5 machines by average difference (live); returns last row used
Private Function TopFiveTable(ws As Worksheet, ByVal r As Long, title As String, _
                              hdrColor As Long, bannerColor As Long, ByRef firstData As Long, ByRef lastData As Long, _
                              plName As String, coName As String) As Long
    Dim k As Long, b As String, cr As String

    SectionTitle ws, r, title, 7, bannerColor
    r = r + 1
    ws.Cells(r, 1).Resize(1, 7).value = Array("Rank", "Machine", "Operations", _
                                              "Planned", "Confirmed", "Total Diff", "Avg Diff / Op")
    StyleHeader ws.Cells(r, 1).Resize(1, 7), hdrColor
    firstData = r + 1

    ' machines in the selected plant/period, ranked by |average difference|, top 5
    ws.Cells(firstData, 2).Formula2 = "=IFERROR(LET(w,UNIQUE(FILTER(hd_WC,hd_In))," & _
        "n,COUNTIFS(hd_WC,w,hd_In,TRUE)," & _
        "p,SUMIFS(" & plName & ",hd_WC,w,hd_In,TRUE)," & _
        "c,SUMIFS(" & coName & ",hd_WC,w,hd_In,TRUE)," & _
        "a,ABS((c-p)/n)," & _
        "k,FILTER(w,a>=0.005),ka,FILTER(a,a>=0.005)," & _
        "s,SORTBY(k,ka,-1)," & _
        "t,INDEX(s,SEQUENCE(MIN(5,ROWS(s))))," & _
        "IF(t="""",""(blank)"",t)),"""")"

    For k = 1 To 5
        r = r + 1
        b = "$B" & r
        cr = KeyCrit(b)
        ws.Cells(r, 1).Formula2 = IfKey(b, CStr(k))
        ws.Cells(r, 3).Formula2 = IfKey(b, "COUNTIFS(hd_WC," & cr & ",hd_In,TRUE)")
        ws.Cells(r, 4).Formula2 = IfKey(b, "SUMIFS(" & plName & ",hd_WC," & cr & ",hd_In,TRUE)")
        ws.Cells(r, 5).Formula2 = IfKey(b, "SUMIFS(" & coName & ",hd_WC," & cr & ",hd_In,TRUE)")
        ws.Cells(r, 6).Formula2 = IfKey(b, "E" & r & "-D" & r)
        ws.Cells(r, 7).Formula2 = IfKey(b, "IFERROR(F" & r & "/C" & r & ",0)")
    Next k
    lastData = r

    FormatHoursBlock ws, firstData, lastData, Array(4, 5), Array(6, 7)
    ws.Range(ws.Cells(firstData, 7), ws.Cells(lastData, 7)).Font.Bold = True
    ws.Range(ws.Cells(firstData, 1), ws.Cells(lastData, 1)).HorizontalAlignment = xlCenter
    LightBorders ws.Range(ws.Cells(firstData - 1, 1), ws.Cells(r, 7))
    TopFiveTable = r
End Function

Private Function LowerBlockFormula() As String
    ' Order types + All machines + Top 10 setup + Top 10 machine, as ONE live block.
    ' It grows and shrinks with the plant slicer, so the tables always sit right under each other.
    ' Same column roles in every table (B:C counts, D:K hours / dates), so plain column formats work.
    ' Columns L:M (hidden): L = row type, M = 9 when a top-10 operation is over the limit.
    Dim f As String
    f = f & "=IFERROR(LET(bl,{"""","""","""","""","""","""","""","""","""","""","""",0,0},ttlOT,{""By order type"","""","""","""","""","""","""","""","""","""","""",1,0},hdOT,{""Order Type"",""Orders"",""Operations"",""Plann"
    f = f & "ed Setup"",""Confirmed Setup"",""Setup Diff"",""Avg Setup Diff / Op"",""Planned Machine"",""Confirmed Machine"",""Machine Diff"",""Avg Machine Diff / Op"",2,0},ttlAM,{""All mac"
    f = f & "hines"","""","""","""","""","""","""","""","""","""","""",1,0},hdAM,{""Machine"",""Orders"",""Operations"",""Planned Setup"",""Confirmed Setup"",""Setup Diff"",""Avg Setup Diff / Op"",""Planned Ma"
    f = f & "chine"",""Confirmed Machine"",""Machine Diff"",""Avg Machine Diff / Op"",2,0},ttlTS,{""Top 10 setup differences"","""","""","""","""","""","""","""","""","""","""",1,0},ttlTM,{""Top 10 machi"
    f = f & "ne differences"","""","""","""","""","""","""","""","""","""","""",1,0},hdTT,{""Order"",""Order Type"",""Item"",""Operation"",""Machine"",""Finish Date"",""Planned"",""Confirmed"",""Difference"","""","""
    f = f & """,2,0},otT,UNIQUE(FILTER(hd_Type,hd_In)),otN,COUNTIFS(hd_Type,otT,hd_In,TRUE),otPS,SUMIFS(hd_PlanSetup,hd_Type,otT,hd_In,TRUE),otCS,SUMIFS(hd_ConfSetup,hd_Type,"
    f = f & "otT,hd_In,TRUE),otPM,SUMIFS(hd_PlanMach,hd_Type,otT,hd_In,TRUE),otCM,SUMIFS(hd_ConfMach,hd_Type,otT,hd_In,TRUE),otB,HSTACK(IF(otT="""",""(blank)"",otT),SUMIFS(hd_Fi"
    f = f & "rstType,hd_Type,otT),otN,otPS,otCS,otCS-otPS,(otCS-otPS)/otN,otPM,otCM,otCM-otPM,(otCM-otPM)/otN,IF(otN,3),IF(otN,0)),amW,UNIQUE(FILTER(hd_WC,hd_In)),amN,COUNTI"
    f = f & "FS(hd_WC,amW,hd_In,TRUE),amO,SUMIFS(hd_FirstMach,hd_WC,amW),amPS,SUMIFS(hd_PlanSetup,hd_WC,amW,hd_In,TRUE),amCS,SUMIFS(hd_ConfSetup,hd_WC,amW,hd_In,TRUE),amPM,S"
    f = f & "UMIFS(hd_PlanMach,hd_WC,amW,hd_In,TRUE),amCM,SUMIFS(hd_ConfMach,hd_WC,amW,hd_In,TRUE),amSA,(amCS-amPS)/amN,amMA,(amCM-amPM)/amN,amB,SORTBY(HSTACK(IF(amW="""",""(bl"
    f = f & "ank)"",amW),amO,amN,amPS,amCS,amCS-amPS,amSA,amPM,amCM,amCM-amPM,amMA,IF(amN,4),IF(amN,0)),ABS(amMA),-1),sRR,'HOURS DATA'!$W$2#,sFF,'HOURS DATA'!$Y$2#,tsB,IF(ISN"
    f = f & "UMBER(INDEX(sRR,1)),HSTACK(INDEX('HOURS DATA'!$B:$B,sRR),INDEX('HOURS DATA'!$C:$C,sRR),INDEX('HOURS DATA'!$D:$D,sRR),INDEX('HOURS DATA'!$E:$E,sRR),INDEX('HOURS "
    f = f & "DATA'!$F:$F,sRR),INDEX('HOURS DATA'!$G:$G,sRR),INDEX('HOURS DATA'!$H:$H,sRR),INDEX('HOURS DATA'!$I:$I,sRR),INDEX('HOURS DATA'!$I:$I,sRR)-INDEX('HOURS DATA'!$H:$"
    f = f & "H,sRR),IF(sRR,""""),IF(sRR,""""),IF(sRR,5),IF(sFF=1,9,0)),bl),mRR,'HOURS DATA'!$X$2#,mFF,'HOURS DATA'!$Z$2#,tmB,IF(ISNUMBER(INDEX(mRR,1)),HSTACK(INDEX('HOURS DATA'!"
    f = f & "$B:$B,mRR),INDEX('HOURS DATA'!$C:$C,mRR),INDEX('HOURS DATA'!$D:$D,mRR),INDEX('HOURS DATA'!$E:$E,mRR),INDEX('HOURS DATA'!$F:$F,mRR),INDEX('HOURS DATA'!$G:$G,mRR)"
    f = f & ",INDEX('HOURS DATA'!$L:$L,mRR),INDEX('HOURS DATA'!$M:$M,mRR),INDEX('HOURS DATA'!$M:$M,mRR)-INDEX('HOURS DATA'!$L:$L,mRR),IF(mRR,""""),IF(mRR,""""),IF(mRR,5),IF(mFF="
    f = f & "1,9,0)),bl),VSTACK(ttlOT,hdOT,otB,bl,ttlAM,hdAM,amB,bl,ttlTS,hdTT,tsB,bl,ttlTM,hdTT,tmB)),"""")"
    LowerBlockFormula = f
End Function




' Formatting for the live block.
' Every table uses the same column roles, so ordinary column formats do the number work:
'   B:C  counts (whole numbers)
'   D:K  hours and hour differences with 2 decimals; a real date (serial 40000 or more) shows as a date
' Conditional formatting only adds colours: title and header bands, row lines, red lines and red negatives.
'   L = row type: 1 title, 2 header, 3 order type, 4 machine, 5 top-10 operation
'   M = 9 when a top-10 operation is over the limit
Private Sub FormatLowerBlock(ws As Worksheet, ByVal z1 As Long, ByVal z2 As Long)
    Dim zn As Range
    Dim rowType As String
    Dim stage As String

    On Error GoTo Oops

    stage = "number formats"
    ws.Range(ws.Cells(z1, 2), ws.Cells(z2, 3)).NumberFormat = "0"
    ws.Range(ws.Cells(z1, 4), ws.Cells(z2, 11)).NumberFormat = "[>=40000]dd/mm/yyyy;0.00"

    stage = "alignment"
    Set zn = ws.Range(ws.Cells(z1, 1), ws.Cells(z2, 11))
    zn.HorizontalAlignment = xlCenter
    zn.VerticalAlignment = xlCenter
    ws.Range(ws.Cells(z1, 1), ws.Cells(z2, 1)).HorizontalAlignment = xlLeft

    ws.Activate
    rowType = "=$L" & z1

    stage = "title rows"
    PutBand ws, z1, z2, rowType & "=1", RGB(0, 112, 192), RGB(255, 255, 255), True

    stage = "header rows"
    PutBand ws, z1, z2, rowType & "=2", RGB(221, 235, 247), RGB(0, 0, 0), True

    stage = "row lines"
    PutLine ws, z1, z2, rowType & ">=3"

    stage = "top 10 lines over the limit"
    PutZoneFont ws, z1, z2, "=($M" & z1 & "=9)*(COLUMN()<=9)", RGB(192, 0, 0)

    stage = "negative numbers in red"
    PutZoneFont ws, z1, z2, "=(COLUMN()>=4)*ISNUMBER(A" & z1 & ")*(A" & z1 & "<0)", RGB(255, 0, 0)
    Exit Sub

Oops:
    MsgBox "The overview was built, but its colours stopped at: " & stage & vbLf & vbLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

' Coloured band over the whole block width for one row type
Private Sub PutBand(ws As Worksheet, ByVal z1 As Long, ByVal z2 As Long, ByVal cond As String, _
                    ByVal fillClr As Long, ByVal fontClr As Long, ByVal isBold As Boolean)
    Dim rng As Range
    Dim fc As FormatCondition

    Set rng = ws.Range(ws.Cells(z1, 1), ws.Cells(z2, 11))
    rng.Cells(1, 1).Select
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=cond)
    fc.Interior.Color = fillClr
    fc.Font.Color = fontClr
    fc.Font.Bold = isBold
End Sub

' Thin grey line under every data row
Private Sub PutLine(ws As Worksheet, ByVal z1 As Long, ByVal z2 As Long, ByVal cond As String)
    Dim rng As Range
    Dim fc As FormatCondition

    Set rng = ws.Range(ws.Cells(z1, 1), ws.Cells(z2, 11))
    rng.Cells(1, 1).Select
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=cond)
    fc.Borders(xlBottom).LineStyle = xlContinuous
    fc.Borders(xlBottom).Color = RGB(217, 217, 217)
End Sub

' Font colour over the whole block width when the condition is true (same kind of rule as the bands).
' Conditions are written without commas, (a)*(b) instead of AND(a,b): a conditional-format formula
' set from VBA uses the PC's list separator, which is a semicolon on a comma-decimal PC.
' The cell is picked inside the condition with COLUMN(); references to A are relative to each cell.
Private Sub PutZoneFont(ws As Worksheet, ByVal z1 As Long, ByVal z2 As Long, ByVal cond As String, ByVal fontClr As Long)
    Dim rng As Range
    Dim fc As FormatCondition

    Set rng = ws.Range(ws.Cells(z1, 1), ws.Cells(z2, 11))
    rng.Cells(1, 1).Select
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=cond)
    fc.Font.Color = fontClr
End Sub

' rows z1:z2 of the given columns, e.g. "D:E,H:I"
Private Function ColsOf(ws As Worksheet, ByVal z1 As Long, ByVal z2 As Long, ByVal cols As String) As Range
    Dim part As Variant
    Dim rng As Range
    Dim colA As String
    Dim colB As String

    For Each part In Split(cols, ",")
        If InStr(part, ":") > 0 Then
            colA = Split(part, ":")(0)
            colB = Split(part, ":")(1)
        Else
            colA = part
            colB = part
        End If
        If rng Is Nothing Then
            Set rng = ws.Range(colA & z1 & ":" & colB & z2)
        Else
            Set rng = Union(rng, ws.Range(colA & z1 & ":" & colB & z2))
        End If
    Next part
    Set ColsOf = rng
End Function

Private Sub FormatHoursBlock(ws As Worksheet, r1 As Long, r2 As Long, plainCols As Variant, diffCols As Variant)
    Dim c As Variant
    If r2 < r1 Then Exit Sub
    For Each c In plainCols
        ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).NumberFormat = "0.00"
    Next c
    For Each c In diffCols
        ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).NumberFormat = "+0.00;[Red]-0.00;0.00"
    Next c
End Sub

Private Function NewChart(ws As Worksheet, l As Double, t As Double, w As Double, h As Double, _
                          ctype As XlChartType, title As String) As Chart
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(l, t, w, h)
    With co.Chart
        .ChartType = ctype
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        .HasTitle = True
        .ChartTitle.Text = title
        .ChartTitle.Font.Size = 11
    End With
    Set NewChart = co.Chart
End Function

' Chart labels/axes want the LOCAL decimal separator (a "." in the code is a literal
' dot on a comma-decimal PC, which turned 0,95 into +0.01)
Private Function ChartNumFmt() As String
    Dim d As String
    d = Application.International(xlDecimalSeparator)
    ChartNumFmt = "+0" & d & "00;-0" & d & "00;0" & d & "00"
End Function

' Same look for all three charts: compact, horizontal bars, rank 1 / first type on top
Private Sub StyleChart(ch As Chart)
    Dim fmt As String
    Dim s As Series

    fmt = ChartNumFmt()
    On Error Resume Next

    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    ch.ChartTitle.Font.Size = 10
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(64, 64, 64)

    ch.Axes(xlCategory).ReversePlotOrder = True          ' first item at the top
    ch.Axes(xlCategory).Crosses = xlMaximum              ' keeps the value axis at the bottom
    ch.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow   ' names stay clear of negative bars
    ch.Axes(xlCategory).TickLabels.Font.Size = 8

    ch.Axes(xlValue).TickLabels.NumberFormatLinked = False
    ch.Axes(xlValue).TickLabels.NumberFormat = fmt
    ch.Axes(xlValue).TickLabels.Font.Size = 8
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 230, 230)

    ch.ChartGroups(1).GapWidth = 50

    For Each s In ch.SeriesCollection
        s.InvertIfNegative = False
        s.HasDataLabels = True
        s.DataLabels.NumberFormat = fmt
        s.DataLabels.Font.Size = 8
        s.DataLabels.Position = xlLabelPositionOutsideEnd
    Next s
End Sub

' Sum of the row heights above row n (points). Used instead of .Top, which can be stale
' while the screen is not updating.
Private Function RowsTop(ws As Worksheet, ByVal n As Long) As Double
    Dim i As Long
    For i = 1 To n - 1
        RowsTop = RowsTop + ws.rows(i).RowHeight
    Next i
End Function

' Sum of the column widths left of column n (points); hidden columns count as 0
Private Function ColsLeft(ws As Worksheet, ByVal n As Long) As Double
    Dim i As Long
    For i = 1 To n - 1
        ColsLeft = ColsLeft + ws.Columns(i).Width
    Next i
End Function

' Stacks the three charts in columns H:K beside the Top 5 tables (rows 8 to 22), equal heights,
' and puts the slicer back in row 6. Runs after the build and every time the overview is opened.
Private Sub PlaceOverviewCharts(ws As Worksheet)
    Dim co As ChartObject
    Dim shp As Shape
    Dim leftEdge As Double
    Dim rightEdge As Double
    Dim topEdge As Double
    Dim bottomEdge As Double
    Dim gap As Double
    Dim h As Double
    Dim i As Long
    Dim nm As Variant

    On Error Resume Next
    gap = 4
    leftEdge = ColsLeft(ws, 8) + 10                 ' column H
    rightEdge = ColsLeft(ws, 12) - 10               ' right edge of column K
    topEdge = RowsTop(ws, 8)
    bottomEdge = RowsTop(ws, 23) - 4
    h = (bottomEdge - topEdge - 2 * gap) / 3

    i = 0
    For Each nm In Array("chtTopSetup", "chtTopMachine", "chtOrderType")
        Set co = Nothing
        Set co = ws.ChartObjects(CStr(nm))
        If Not co Is Nothing Then
            co.Placement = xlFreeFloating
            co.Left = leftEdge
            co.Width = rightEdge - leftEdge
            co.Top = topEdge + i * (h + gap)
            co.Height = h
        End If
        i = i + 1
    Next nm

    For Each shp In ws.Shapes
        If shp.Type = msoSlicer Then
            shp.Left = ColsLeft(ws, 3) + 4
            shp.Top = RowsTop(ws, 6) + 3
        End If
    Next shp
End Sub

Private Sub AddNamedSeries(ch As Chart, ByVal nm As String, ByVal catName As String, ByVal valName As String, _
                           ByVal clr As Long, fbCat As Range, fbVal As Range)
    Dim s As Series
    Dim bk As String

    bk = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    Set s = ch.SeriesCollection.NewSeries
    On Error Resume Next
    s.Formula = "=SERIES(""" & nm & """," & bk & catName & "," & bk & valName & "," & ch.SeriesCollection.Count & ")"
    If Err.Number <> 0 Then
        Err.Clear
        s.Name = nm
        s.XValues = fbCat
        s.Values = fbVal
    End If
    s.Format.Fill.ForeColor.RGB = clr
End Sub

Private Sub AddSeries(ch As Chart, nm As String, vals As Range, cats As Range, clr As Long)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.Values = vals
    s.XValues = cats
    s.Format.Fill.ForeColor.RGB = clr
End Sub

' =========================================================
' SHEET BUILDING HELPERS
' =========================================================

' Sheets this macro owns and rebuilds every run
Private Function IsOutputSheet(nm As String) As Boolean
    IsOutputSheet = (nm = SUMMARY_NAME) Or (nm = OVERVIEW_NAME) Or (nm = DATA_NAME) Or (nm = PICK_NAME) Or (nm = GOOD_NAME) Or (nm = "Good Variances") Or (nm = "HOURS SUMMARY") Or _
                    (nm = deptName(DEPT_WELD)) Or (nm = deptName(DEPT_PAINT)) Or _
                    (nm = deptName(DEPT_MACH)) Or (nm = deptName(DEPT_ASSY)) Or _
                    (nm = deptName(DEPT_OTHER)) Or IsSpecialDeptSheet(nm)
End Function

Private Function deptName(spec As String) As String
    deptName = Split(spec, "|")(0)
End Function

' Row 1 banner + row 2 headers, same on every sheet
Private Sub HeaderBlock(ws As Worksheet)
    With ws.Range("D1:F1")
        .Merge
        .value = "Comments"
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A2:I2")
        .value = Array("Plant", "Production Order", "Item", "Operation", "Comments on Machine", _
                       "Comments on Setup", "Remarks", "Actual Start Date", "Actual Op. Finish")
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' Department tabs: Op and Work Center are separate columns
Private Sub DeptHeaderBlock(ws As Worksheet)
    With ws.Range("F1:H1")
        .Merge
        .value = "Comments"
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A2:J2")
        .value = Array("Plant", "Production Order", "Item", "Op", "Work Center", _
                       "Comments on Machine", "Comments on Setup", "Remarks", _
                       "Actual Start Date", "Actual Op. Finish")
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub DeptWidths(ws As Worksheet)
    ws.Columns(1).ColumnWidth = 9
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 13
    ws.Columns(4).ColumnWidth = 9
    ws.Columns(5).ColumnWidth = 14
    ws.Columns(6).ColumnWidth = 42
    ws.Columns(7).ColumnWidth = 42
    ws.Columns(8).ColumnWidth = 30
    ws.Columns(9).ColumnWidth = 18
    ws.Columns(10).ColumnWidth = 18
End Sub

Private Sub SetWidths(ws As Worksheet)
    ws.Columns(1).ColumnWidth = 9
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 13
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 42
    ws.Columns(6).ColumnWidth = 42
    ws.Columns(7).ColumnWidth = 30
    ws.Columns(8).ColumnWidth = 18
    ws.Columns(9).ColumnWidth = 18
End Sub

' Filter buttons on the row 2 headers
Private Sub ApplyFilter(ws As Worksheet, lastCol As Long, lastRow As Long)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If lastRow < 2 Then lastRow = 2
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    On Error GoTo 0
End Sub

Private Sub FreezeTop(ws As Worksheet)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
End Sub

' The export sheet: "Sheet1", else the first sheet that is not one of the sheets this macro builds
Private Function GetSourceSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set GetSourceSheet = ThisWorkbook.Worksheets(SOURCE_NAME)
    On Error GoTo 0
    If GetSourceSheet Is Nothing Then
        For Each ws In ThisWorkbook.Worksheets
            If Not IsOutputSheet(ws.Name) Then
                Set GetSourceSheet = ws
                Exit For
            End If
        Next ws
    End If
End Function

' Does this work center belong to the tab? patterns are ";"-separated Like masks
Private Function WCMatches(wc As String, patterns As String) As Boolean
    Dim p As Variant, u As String
    u = UCase(Trim(wc))
    For Each p In Split(UCase(patterns), ";")
        If u Like CStr(p) Then WCMatches = True: Exit Function
    Next p
End Function

' -------------------------------------------------------
' One department tab: one row per operation, newest order first
' deptName = the sheet name; the operations come from dDept()
' -------------------------------------------------------
Private Sub BuildDeptSheet(ByVal deptName As String)
    Dim ws As Worksheet
    Dim pick() As Long, cntPick As Long, k As Long, j As Long, r As Long
    Dim ordNames() As String, ordKeyD() As Double, ordIdx() As Long, nOrders As Long
    Dim o As Long, found As Long

    ' Which operations belong to this department
    If dN = 0 Then Exit Sub
    ReDim pick(1 To dN)
    For k = 1 To dN
        If dDept(k) = deptName Then
            cntPick = cntPick + 1
            pick(cntPick) = k
        End If
    Next k

    ' A department without any operations (no work centers) gets no sheet at all
    If cntPick = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = deptName
    ws.Cells.Font.Name = "Aptos Narrow"
    ws.Cells.Font.Size = 11
    ws.Cells.Font.Italic = False          ' department tabs: plain text, not commentary style
    ws.Cells.Font.Color = RGB(0, 0, 0)
    DeptHeaderBlock ws
    DeptWidths ws

    ' Group by order, key each order on its latest operation date
    ReDim ordNames(1 To cntPick): ReDim ordKeyD(1 To cntPick): ReDim ordIdx(1 To cntPick)
    For k = 1 To cntPick
        j = pick(k)
        found = 0
        For o = 1 To nOrders
            If ordNames(o) = dOrder(j) Then found = o: Exit For
        Next o
        If found = 0 Then
            nOrders = nOrders + 1
            found = nOrders
            ordNames(nOrders) = dOrder(j)
            ordIdx(nOrders) = nOrders
            ordKeyD(nOrders) = -1E+15
        End If
        If CDbl(dDate(j)) > ordKeyD(found) Then ordKeyD(found) = CDbl(dDate(j))
    Next k

    If Not SORT_NEWEST_FIRST Then
        For o = 1 To nOrders
            ordKeyD(o) = -ordKeyD(o)
        Next o
    End If
    SortDesc ordKeyD, ordIdx, nOrders

    ' Write the rows, one order block at a time
    r = 3
    For o = 1 To nOrders
        For k = 1 To cntPick
            j = pick(k)
            If dOrder(j) = ordNames(ordIdx(o)) Then
                ws.Cells(r, 1).value = AsNum(dPlant(j))
                ws.Cells(r, 2).value = AsNum(dOrder(j))
                ws.Cells(r, 3).value = AsNum(dItem(j))
                ws.Cells(r, 4).value = dOp(j)
                ws.Cells(r, 5).value = dWC(j)
                ws.Cells(r, 6).value = dMach(j)
                ws.Cells(r, 7).value = dSetup(j)
                ws.Cells(r, 8).value = GetRemark(RemarkKey(dOrder(j), dOp(j), dWC(j)))
                WriteDate ws.Cells(r, 9), dStartD(j)
                WriteDate ws.Cells(r, 10), dFinishD(j)
                If dMRed(j) Then ws.Cells(r, 6).Font.Color = RGB(255, 0, 0)
                If dSRed(j) Then ws.Cells(r, 7).Font.Color = RGB(255, 0, 0)
                r = r + 1
            End If
        Next k
    Next o

    Dim lastR As Long
    lastR = r - 1

    With ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 10))
        .VerticalAlignment = xlTop
        .WrapText = True
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(191, 191, 191)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 3)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(3, 4), ws.Cells(lastR, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(3, 6), ws.Cells(lastR, 8)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(3, 9), ws.Cells(lastR, 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 10)).rows.AutoFit
    ApplyFilter ws, 10, lastR
    FreezeTop ws
End Sub

' Real date value when there is one, otherwise a dash
Private Sub WriteDate(cel As Range, d As Date)
    If d = 0 Then
        cel.NumberFormat = "@"
        cel.value = "-"
    Else
        cel.NumberFormat = "dd/mm/yyyy"
        cel.value = CDate(d)
    End If
End Sub

' =========================================================
' GOOD VARIANCES SHEET
' Operations that beat the plan: confirmed hours BELOW planned hours, and over the same
' thresholds as the red lines (machine 2 hrs, setup 1 hr, or 50%). They are left off the
' Comments sheet and the department tabs. The Department column is the filter.
' =========================================================
Private Sub BuildGoodSheet()
    Dim ws As Worksheet
    Dim ordNames() As String, ordKeyD() As Double, ordIdx() As Long, nOrders As Long
    Dim k As Long, o As Long, r As Long, found As Long, lastR As Long
    Dim goodBanner As Long, goodHeader As Long

    goodBanner = RGB(0, 112, 192)
    goodHeader = RGB(221, 235, 247)

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = GOOD_NAME
    ws.Cells.Font.Name = "Aptos Narrow"
    ws.Cells.Font.Size = 11

    ' Note at the top: ask for the cause, plus the highlight key
    Dim noteTop As String, noteKey As String
    noteTop = "If there was a specific cause for the favourable variance, please address it in the Remarks column."
    noteKey = "Highlighted: under plan by more than " & GOOD_HRS_HILITE & " hrs"
    With ws.Range("A1:F1")
        .Merge
        .value = noteTop & vbLf & noteKey
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 11
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
    End With
    With ws.Range("A1").Characters(1, Len(noteTop)).Font
        .Bold = True
        .Italic = False
        .Color = RGB(0, 0, 0)
    End With
    ws.rows(1).RowHeight = 48

    With ws.Range("G1:I1")
        .Merge
        .value = "Comments - better than planned"
        .Interior.Color = goodBanner
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A2:K2")
        .value = Array("Plant", "Production Order", "Item", "Op", "Work Center", "Department", _
                       "Comments on Machine", "Comments on Setup", "Remarks", _
                       "Actual Start Date", "Actual Op. Finish")
        .Interior.Color = goodHeader
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Columns(1).ColumnWidth = 9
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 13
    ws.Columns(4).ColumnWidth = 9
    ws.Columns(5).ColumnWidth = 14
    ws.Columns(6).ColumnWidth = 26
    ws.Columns(7).ColumnWidth = 42
    ws.Columns(8).ColumnWidth = 42
    ws.Columns(9).ColumnWidth = 30
    ws.Columns(10).ColumnWidth = 18
    ws.Columns(11).ColumnWidth = 18

    If gN = 0 Then
        ws.Range("A3").value = "No operations finished better than planned (over the thresholds)."
        ws.Range("A3").Font.Italic = True
        GoodFilterAndFreeze ws, 3
        Exit Sub
    End If

    ' Orders newest first, keyed on the latest operation date of the order
    ReDim ordNames(1 To gN): ReDim ordKeyD(1 To gN): ReDim ordIdx(1 To gN)
    For k = 1 To gN
        found = 0
        For o = 1 To nOrders
            If ordNames(o) = gOrder(k) Then found = o: Exit For
        Next o
        If found = 0 Then
            nOrders = nOrders + 1
            found = nOrders
            ordNames(nOrders) = gOrder(k)
            ordIdx(nOrders) = nOrders
            ordKeyD(nOrders) = -1E+15
        End If
        If CDbl(gDate(k)) > ordKeyD(found) Then ordKeyD(found) = CDbl(gDate(k))
    Next k
    If Not SORT_NEWEST_FIRST Then
        For o = 1 To nOrders
            ordKeyD(o) = -ordKeyD(o)
        Next o
    End If
    SortDesc ordKeyD, ordIdx, nOrders

    r = 3
    For o = 1 To nOrders
        For k = 1 To gN
            If gOrder(k) = ordNames(ordIdx(o)) Then
                ws.Cells(r, 1).value = AsNum(gPlant(k))
                ws.Cells(r, 2).value = AsNum(gOrder(k))
                ws.Cells(r, 3).value = AsNum(gItem(k))
                ws.Cells(r, 4).value = gOp(k)
                ws.Cells(r, 5).value = gWC(k)
                ws.Cells(r, 6).value = DeptOf(gPlant(k), gItem(k), gWC(k))
                ws.Cells(r, 7).value = gMach(k)
                ws.Cells(r, 8).value = gSetup(k)
                ws.Cells(r, 9).value = GetRemark(RemarkKey(gOrder(k), gOp(k), gWC(k)))
                WriteDate ws.Cells(r, 10), gStartD(k)
                WriteDate ws.Cells(r, 11), gFinishD(k)
                If gMHi(k) Then HighlightFavourable ws.Cells(r, 7)
                If gSHi(k) Then HighlightFavourable ws.Cells(r, 8)
                r = r + 1
            End If
        Next k
    Next o
    lastR = r - 1

    With ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 11))
        .VerticalAlignment = xlTop
        .WrapText = True
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(191, 191, 191)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 3)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(3, 4), ws.Cells(lastR, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(3, 7), ws.Cells(lastR, 9)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(3, 10), ws.Cells(lastR, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 11)).rows.AutoFit
    GoodFilterAndFreeze ws, lastR
End Sub

' Filter buttons on the header row (Department is the main one), header frozen
Private Sub GoodFilterAndFreeze(ws As Worksheet, ByVal lastR As Long)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If lastR < 3 Then lastR = 3
    ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 11)).AutoFilter
    On Error GoTo 0
    FreezeTop ws
End Sub

' Soft amber highlight (friendlier than red) for a part that is well under plan
Private Sub HighlightFavourable(cel As Range)
    cel.Interior.Color = RGB(255, 235, 156)
    cel.Font.Color = RGB(156, 87, 0)
End Sub

' Department of a work center: same patterns, same order as the department tabs
Private Function DeptOfWC(ByVal wc As String) As String
    Dim specs As Variant
    Dim i As Long
    Dim pats As String

    specs = Array(DEPT_WELD, DEPT_PAINT, DEPT_MACH, DEPT_ASSY, DEPT_OTHER)
    For i = LBound(specs) To UBound(specs)
        pats = CStr(Split(CStr(specs(i)), "|")(1))
        If WCMatches(wc, pats) Then
            DeptOfWC = deptName(CStr(specs(i)))
            Exit Function
        End If
    Next i
    DeptOfWC = deptName(DEPT_OTHER)
End Function

' Special departments: plant 1300 working on one of the listed materials goes to the sheet named
' after the material (SPECIAL_MATERIALS), not to the work-center departments
Private Function SpecialDept(ByVal plant As String, ByVal item As String) As String
    Dim p As Variant
    Dim parts() As String

    If Trim(plant) <> SPECIAL_PLANT Then Exit Function
    item = UCase(Trim(item))
    For Each p In Split(SPECIAL_MATERIALS, ";")
        parts = Split(CStr(p), "=")
        If UCase(Trim(parts(0))) = item Then
            SpecialDept = Trim(parts(1))
            Exit Function
        End If
    Next p
End Function

' Department of an operation: special material of plant 1300 first, otherwise by work center
Private Function DeptOf(ByVal plant As String, ByVal item As String, ByVal wc As String) As String
    Dim nm As String

    nm = SpecialDept(plant, item)
    If Len(nm) > 0 Then
        DeptOf = nm
    Else
        DeptOf = DeptOfWC(wc)
    End If
End Function

Private Function IsSpecialDeptSheet(ByVal nm As String) As Boolean
    Dim p As Variant

    For Each p In Split(SPECIAL_MATERIALS, ";")
        If nm = Trim(Split(CStr(p), "=")(1)) Then
            IsSpecialDeptSheet = True
            Exit Function
        End If
    Next p
End Function

' =========================================================
' REMARKS: read what is already typed before rebuilding,
' then put it back on the matching row
'   Comments sheet key = order & "|"
'   Department tab key = order & "|OP20(MCH14)"
' =========================================================
' key = order | OP40 | MCH28   (op and work center blank for the Comments sheet)
Private Function RemarkKey(ordTxt As String, opTxt As String, wcTxt As String) As String
    RemarkKey = UCase(Trim(ordTxt)) & "|" & UCase(Trim(opTxt)) & "|" & UCase(Trim(wcTxt))
End Function

Private Sub HarvestRemarks()
    Dim ws As Worksheet, hRow As Long, c As Long
    Dim colOrder As Long, colOp As Long, colWC As Long, colRem As Long
    Dim r As Long, lastR As Long, ordTxt As String, opTxt As String, wcTxt As String, rmk As String

    rkN = 0
    ReDim rkKey(1 To 5000): ReDim rkVal(1 To 5000)

    For Each ws In ThisWorkbook.Worksheets
        If IsOutputSheet(ws.Name) And ws.Name <> OVERVIEW_NAME And ws.Name <> DATA_NAME And ws.Name <> PICK_NAME Then
            ' find the header row and the columns we need
            colOrder = 0: colOp = 0: colWC = 0: colRem = 0
            For hRow = 1 To 5
                For c = 1 To 14
                    Select Case Trim(CStr(ws.Cells(hRow, c).value))
                        Case "Production Order": colOrder = c
                        Case "Op", "Operation": colOp = c
                        Case "Work Center", "WorkCenter", "Work Centre": colWC = c
                        Case "Remarks": colRem = c
                    End Select
                Next c
                If colOrder > 0 And colRem > 0 Then Exit For
            Next hRow

            If colOrder > 0 And colRem > 0 Then
                lastR = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
                ordTxt = ""
                For r = hRow + 1 To lastR
                    If Trim(CStr(ws.Cells(r, colOrder).value)) <> "" Then ordTxt = Trim(CStr(ws.Cells(r, colOrder).value))
                    rmk = Trim(CStr(ws.Cells(r, colRem).value))
                    If ordTxt <> "" And rmk <> "" Then
                        opTxt = "": wcTxt = ""
                        If colOp > 0 Then opTxt = Trim(CStr(ws.Cells(r, colOp).value))
                        If colWC > 0 Then wcTxt = Trim(CStr(ws.Cells(r, colWC).value))
                        If InStr(opTxt, vbLf) > 0 Then opTxt = ""      ' Comments sheet: many ops in one cell
                        ' Old layout: "OP40(MCH28)" in one cell
                        If wcTxt = "" And InStr(opTxt, "(") > 0 Then
                            wcTxt = Replace(Mid(opTxt, InStr(opTxt, "(") + 1), ")", "")
                            opTxt = Left(opTxt, InStr(opTxt, "(") - 1)
                        End If
                        AddRemark RemarkKey(ordTxt, opTxt, wcTxt), rmk
                    End If
                Next r
            End If
        End If
    Next ws
End Sub

Private Sub AddRemark(key As String, txt As String)
    If rkN >= UBound(rkKey) Then Exit Sub
    rkN = rkN + 1
    rkKey(rkN) = key
    rkVal(rkN) = txt
End Sub

Private Function GetRemark(key As String) As String
    Dim k As Long
    For k = 1 To rkN
        If rkKey(k) = key Then GetRemark = rkVal(k): Exit Function
    Next k
End Function

Private Sub SectionTitle(ws As Worksheet, r As Long, txt As String, cols As Long, clr As Long)
    With ws.Cells(r, 1).Resize(1, cols)
        .Interior.Color = clr
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Cells(r, 1).value = txt
End Sub

Private Sub StyleHeader(rng As Range, clr As Long)
    With rng
        .Interior.Color = clr
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub LightBorders(rng As Range)
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Color = RGB(217, 217, 217)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(191, 191, 191)
    End With
End Sub

' Insertion sort, descending by keys, carrying the index array
Private Sub SortDesc(keys() As Double, idx() As Long, cnt As Long)
    Dim a As Long, b As Long, kTmp As Double, iTmp As Long
    For a = 2 To cnt
        kTmp = keys(a): iTmp = idx(a)
        b = a - 1
        Do While b >= 1
            If keys(b) >= kTmp Then Exit Do
            keys(b + 1) = keys(b): idx(b + 1) = idx(b)
            b = b - 1
        Loop
        keys(b + 1) = kTmp: idx(b + 1) = iTmp
    Next a
End Sub

' =========================================================
' Helpers
' =========================================================

' RED when the hour difference OR the % difference hits its threshold
Private Function IsRed(planned As Double, confirmed As Double, hrsThreshold As Double) As Boolean
    Dim d As Double
    d = Abs(confirmed - planned)
    If d < 0.005 Then
        IsRed = False
    ElseIf d >= hrsThreshold Then
        IsRed = True
    ElseIf planned = 0 Then
        IsRed = True                     ' anything confirmed on 0 planned = infinite %
    Else
        IsRed = (d / planned >= PCT_THRESHOLD)
    End If
End Function

Private Function DiffTxt(planned As Double, confirmed As Double) As String
    Dim d As Double
    d = confirmed - planned
    DiffTxt = SignedHrs(d) & IIf(d > 0, " over plan", " under plan")
    If planned <> 0 Then DiffTxt = DiffTxt & "  (" & Format(Abs(d) / planned, "0%") & ")"
End Function

Private Function SignedHrs(d As Double) As String
    SignedHrs = IIf(d >= 0, "+", "-") & FormatHrs(Abs(d))
End Function

' Work centers that are never counted: starts with GI (goods issue), starts with
' GR (goods receipt), or contains INSP (inspection booking points)
Private Function IsExcludedWC(ByVal wc As String) As Boolean
    wc = UCase(Trim(wc))
    IsExcludedWC = (wc Like "GI*") Or (wc Like "GR*") Or (InStr(wc, "INSP") > 0)
End Function

' "14 working days, Mon-Sat" or "30 calendar days"
Private Function WindowLabel() As String
    If USE_CALENDAR Then
        WindowLabel = WINDOW_DAYS & " calendar days"
    Else
        WindowLabel = WINDOW_DAYS & " working days, Mon" & Chr(150) & "Sat"
    End If
End Function

' Count back n working days (Mon-Sat) from endDate, endDate included
Private Function StartOfWorkingWindow(endDate As Date, n As Long) As Date
    Dim d As Date, counted As Long
    d = endDate
    Do
        If Weekday(d, vbMonday) <> 7 Then counted = counted + 1   ' 7 = Sunday
        If counted >= n Then Exit Do
        d = d - 1
    Loop
    StartOfWorkingWindow = d
End Function

' Finish date, else start date, else 0
Private Function OpDateOf(finishV As Variant, startV As Variant) As Date
    OpDateOf = ToDate(finishV)
    If OpDateOf = 0 Then OpDateOf = ToDate(startV)
End Function

Private Function ToDate(v As Variant) As Date
    ToDate = 0
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        ToDate = Int(v)
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then ToDate = Int(CDbl(v))
    ElseIf IsDate(v) Then
        ToDate = Int(CDate(v))
    End If
End Function

Private Sub ColorLines(cel As Range, txt As String, flags As String)
    Dim lines() As String, k As Long, pos As Long
    lines = Split(txt, vbLf)
    pos = 1
    On Error Resume Next
    For k = 0 To UBound(lines)
        If Mid(flags, k + 1, 1) = "1" Then
            cel.Characters(pos, Len(lines(k))).Font.Color = RGB(255, 0, 0)
        End If
        pos = pos + Len(lines(k)) + 1
    Next k
    On Error GoTo 0
End Sub

Private Function NumVal(v As Variant) As Double
    If IsError(v) Then
        NumVal = 0
    ElseIf IsNumeric(v) Then
        NumVal = CDbl(v)
    Else
        NumVal = 0
    End If
End Function

Private Function AsNum(s As String) As Variant
    If Len(s) > 0 And Not (s Like "*[!0-9]*") Then
        AsNum = CDbl(s)
    Else
        AsNum = s
    End If
End Function

Private Function DateStr(v As Variant) As String
    Dim d As Date
    d = ToDate(v)
    If d = 0 Then DateStr = "-" Else DateStr = Format(d, "dd\/mm\/yyyy")
End Function

Private Function FormatHrs(v As Double) As String
    Dim x As Double
    x = Round(v, 2)
    If x = 1 Then
        FormatHrs = "1hr"
    ElseIf x = Int(x) Then
        FormatHrs = CStr(CLng(x)) & "hrs"
    Else
        FormatHrs = Replace(CStr(x), ".", ",") & "hrs"
    End If
End Function


