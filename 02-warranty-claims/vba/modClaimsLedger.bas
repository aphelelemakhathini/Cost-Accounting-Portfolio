Attribute VB_Name = "modClaimsLedger"
'===============================================================================
' modClaimsLedger: step 2 of the monthly warranty & goodwill claims process
'
' Run with the finished month sheet (from modClaimsExport, renamed e.g.
' "October 2026") active inside the year-to-date workbook. Warranty orders
' (WA01) go to CLAIMS 2026, goodwill orders (GW01) to GOODWILL 2026: existing
' orders get this month's amount, new orders fill a spare row or are inserted
' above the Total row, and the YTD / Total-to-date / Total-row formulas are
' rebuilt. The GRAPH sheet's current-month cells are moved to this month.
' Remember to update the month's turnover on GRAPH.
'===============================================================================

Option Explicit

'==============================================================================
' FeedMonthlyDataToLedgers
' ----------------------------------------------------------------------------
' Run this with the FINISHED MONTHLY SHEET active (e.g. "August" or
' "August 2026" - the tab must contain the month name).
'
' Source sheet layout (your Claims Summary, inserted into the report):
'   Order | Plant | CLM NO | Descr.  Deb. | Description | VALUE | Order Type
'
' Splits by Order Type:
'   WA01  -> CLAIMS 2026    (Order No | Customer | DESCRIPTION | Claim No |
'                             Accum to Dec 25 | Jan..Dec | YTD 2026 |
'                             Total Costs to Date)
'   GW01  -> GOODWILL 2026  (Order No | Customer | Description |
'                             Jan..Dec | YTD 2026)
'
' For each ledger:
'   - Same Order No already there  -> writes the amount into THIS month's
'     column for that row (copied, not linked). Claim No filled in if blank.
'   - Order No not found           -> inserts a new row directly ABOVE that
'     sheet's Total row (so the Total row always stays at the bottom).
'   - After processing: fills down the YTD formula (and, on Claims 2026, the
'     Total-Costs-to-Date formula) for every data row, and refreshes the
'     Total row's SUM ranges so they always cover the full data range.
'
' Re-run safe: re-running for the same month overwrites that month's figure
' rather than doubling it.
'==============================================================================

Sub FeedMonthlyDataToLedgers()
    Dim wsSrc As Worksheet
    Dim srcHdrRow As Long
    Dim srcOrderCol As Long, srcClmCol As Long, srcDescrDebCol As Long
    Dim srcDescCol As Long, srcValueCol As Long, srcTypeCol As Long

    Set wsSrc = ActiveSheet
    srcHdrRow = FindHeaderRow(wsSrc, "Order")
    If srcHdrRow = 0 Then
        MsgBox "Could not find an 'Order' header on '" & wsSrc.Name & "'.", vbCritical
        Exit Sub
    End If

    srcOrderCol = FindHeaderCol(wsSrc, srcHdrRow, "Order")
    srcClmCol = FindHeaderCol(wsSrc, srcHdrRow, "CLM")
    srcDescrDebCol = FindHeaderCol(wsSrc, srcHdrRow, "Descr")
    srcDescCol = FindHeaderCol(wsSrc, srcHdrRow, "Description")
    srcValueCol = FindHeaderCol(wsSrc, srcHdrRow, "VALUE")
    srcTypeCol = FindHeaderCol(wsSrc, srcHdrRow, "Order Type")

    If srcOrderCol = 0 Or srcValueCol = 0 Or srcTypeCol = 0 Then
        MsgBox "Could not find Order / VALUE / Order Type columns on '" & wsSrc.Name & "'.", vbCritical
        Exit Sub
    End If

    Dim claimsUpd As Long, claimsNew As Long, gwUpd As Long, gwNew As Long

    Application.ScreenUpdating = False

    FeedLedger wsSrc, srcHdrRow, srcOrderCol, srcClmCol, srcDescrDebCol, srcDescCol, srcValueCol, srcTypeCol, _
               "WA01", ThisWorkbook.Worksheets("CLAIMS 2026"), True, True, claimsUpd, claimsNew

    FeedLedger wsSrc, srcHdrRow, srcOrderCol, srcClmCol, srcDescrDebCol, srcDescCol, srcValueCol, srcTypeCol, _
               "GW01", ThisWorkbook.Worksheets("GOODWILL 2026"), False, False, gwUpd, gwNew

    UpdateGraphSnapshot wsSrc.Name

    Application.ScreenUpdating = True

    MsgBox "Done feeding '" & wsSrc.Name & "':" & vbNewLine & vbNewLine & _
           "CLAIMS 2026 (WA01): " & claimsUpd & " updated, " & claimsNew & " new" & vbNewLine & _
           "GOODWILL 2026 (GW01): " & gwUpd & " updated, " & gwNew & " new" & vbNewLine & _
           "GRAPH snapshot (O2/P2) moved to this month", vbInformation
End Sub

'------------------------------------------------------------------------------
' Feeds all source rows matching typeFilter (e.g. "WA01") into wsDst.
'   hasClaimNoCol      - True for Claims 2026 (Claim No lives in column D)
'   hasTotalToDateCol  - True for Claims 2026 (has Accum-to-Dec25 + a
'                        Total-Costs-to-Date column after YTD); False for
'                        Goodwill 2026 (months start right after Description,
'                        only a YTD column follows).
'------------------------------------------------------------------------------
Private Sub FeedLedger(wsSrc As Worksheet, srcHdrRow As Long, srcOrderCol As Long, srcClmCol As Long, _
                        srcDescrDebCol As Long, srcDescCol As Long, srcValueCol As Long, srcTypeCol As Long, _
                        typeFilter As String, wsDst As Worksheet, hasClaimNoCol As Boolean, hasTotalToDateCol As Boolean, _
                        ByRef updateCount As Long, ByRef newCount As Long)

    Dim dstHdrRow As Long, dstFirstDataRow As Long, dstLastDataRow As Long, totalRow As Long
    Dim firstMonthCol As Long, lastMonthCol As Long, monthCol As Long, ytdCol As Long, totalToDateCol As Long
    Const claimNoCol As Long = 4   ' Claims 2026 only: Claim No sits in column D
    Dim r As Long, d As Long, matchRow As Long, blankRow As Long
    Dim orderNo As String, clmNo As Variant, custName As Variant, descr As Variant, otype As String
    Dim amt As Double

    updateCount = 0
    newCount = 0

    dstHdrRow = FindHeaderRow(wsDst, "Order")
    If dstHdrRow = 0 Then
        MsgBox "Could not find an 'Order' header on '" & wsDst.Name & "'.", vbCritical
        Exit Sub
    End If

    firstMonthCol = FindColumnForMonthNumber(wsDst, dstHdrRow, 1)    ' January
    lastMonthCol = FindColumnForMonthNumber(wsDst, dstHdrRow, 12)    ' December
    If firstMonthCol = 0 Or lastMonthCol = 0 Then
        MsgBox "Could not find the Jan-Dec month columns on '" & wsDst.Name & "'.", vbCritical
        Exit Sub
    End If

    monthCol = FindMonthColumn(wsDst, dstHdrRow, wsSrc.Name)
    If monthCol = 0 Then
        MsgBox "Could not match sheet name '" & wsSrc.Name & "' to a month column on '" & wsDst.Name & "'." & vbNewLine & _
               "Make sure the sheet tab name contains the month, e.g. 'August' or 'August 2026'.", vbCritical
        Exit Sub
    End If

    ytdCol = lastMonthCol + 1
    totalToDateCol = ytdCol + 1   ' only used when hasTotalToDateCol = True

    dstFirstDataRow = dstHdrRow + 1
    Do While Trim(CStr(wsDst.Cells(dstFirstDataRow, "A").value)) = "" And dstFirstDataRow < dstHdrRow + 5
        dstFirstDataRow = dstFirstDataRow + 1
    Loop

    totalRow = FindTotalRow(wsDst, dstFirstDataRow)
    If totalRow = 0 Then
        MsgBox "Could not find the 'Total' row on '" & wsDst.Name & "'. Stopping so nothing gets misplaced.", vbCritical
        Exit Sub
    End If
    dstLastDataRow = totalRow - 1

    r = srcHdrRow + 1
    Do While r <= wsSrc.Cells(wsSrc.rows.Count, srcOrderCol).End(xlUp).Row
        orderNo = Trim(CStr(wsSrc.Cells(r, srcOrderCol).value))
        otype = Trim(CStr(wsSrc.Cells(r, srcTypeCol).value))

        If orderNo = "" Then GoTo NextRow                          ' TOTAL row on the source sheet
        If UCase(otype) <> UCase(typeFilter) Then GoTo NextRow     ' belongs to the other ledger

        amt = 0
        If IsNumeric(wsSrc.Cells(r, srcValueCol).value) Then amt = CDbl(wsSrc.Cells(r, srcValueCol).value)
        If srcClmCol > 0 Then clmNo = wsSrc.Cells(r, srcClmCol).value Else clmNo = ""
        If srcDescrDebCol > 0 Then custName = wsSrc.Cells(r, srcDescrDebCol).value Else custName = ""
        If srcDescCol > 0 Then descr = wsSrc.Cells(r, srcDescCol).value Else descr = ""

        matchRow = 0
        For d = dstFirstDataRow To dstLastDataRow
            If Trim(CStr(wsDst.Cells(d, "A").value)) = orderNo Then
                matchRow = d
                Exit For
            End If
        Next d

        If matchRow > 0 Then
            wsDst.Cells(matchRow, monthCol).value = amt            ' copy - not linked
            If hasClaimNoCol Then
                If Trim(CStr(wsDst.Cells(matchRow, claimNoCol).value)) = "" And Trim(CStr(clmNo)) <> "" Then
                    wsDst.Cells(matchRow, claimNoCol).value = clmNo
                End If
            End If
            updateCount = updateCount + 1
        Else
            ' Reuse an existing blank data row first (many report sheets are
            ' pre-built with spare rows and formulas above the Total row) -
            ' only insert a brand new row if none are free.
            blankRow = FindBlankDataRow(wsDst, dstFirstDataRow, dstLastDataRow)

            If blankRow > 0 Then
                wsDst.Cells(blankRow, "A").value = orderNo
                wsDst.Cells(blankRow, "B").value = custName
                wsDst.Cells(blankRow, "C").value = descr
                If hasClaimNoCol Then wsDst.Cells(blankRow, claimNoCol).value = clmNo
                wsDst.Cells(blankRow, monthCol).value = amt
            Else
                wsDst.rows(totalRow).Insert Shift:=xlDown          ' new row goes ABOVE the total
                wsDst.Cells(totalRow, "A").value = orderNo
                wsDst.Cells(totalRow, "B").value = custName
                wsDst.Cells(totalRow, "C").value = descr
                If hasClaimNoCol Then wsDst.Cells(totalRow, claimNoCol).value = clmNo
                wsDst.Cells(totalRow, monthCol).value = amt

                totalRow = totalRow + 1
                dstLastDataRow = dstLastDataRow + 1
            End If
            newCount = newCount + 1
        End If

NextRow:
        r = r + 1
    Loop

    '--- fill down YTD (and Total-Costs-to-Date on Claims 2026) for every data row ---
    Dim rr As Long, firstMonthLetter As String, lastMonthLetter As String
    Dim ytdLetter As String, accumLetter As String

    firstMonthLetter = Split(wsDst.Cells(1, firstMonthCol).Address, "$")(1)
    lastMonthLetter = Split(wsDst.Cells(1, lastMonthCol).Address, "$")(1)
    ytdLetter = Split(wsDst.Cells(1, ytdCol).Address, "$")(1)

    For rr = dstFirstDataRow To dstLastDataRow
        wsDst.Cells(rr, ytdCol).Formula = "=SUM(" & firstMonthLetter & rr & ":" & lastMonthLetter & rr & ")"
        If hasTotalToDateCol Then
            accumLetter = Split(wsDst.Cells(1, firstMonthCol - 1).Address, "$")(1)   ' Accum to Dec 25, just left of Jan
            wsDst.Cells(rr, totalToDateCol).Formula = "=" & ytdLetter & rr & "+" & accumLetter & rr
        End If
    Next rr

    '--- refresh the Total row's SUM formulas so they always cover the full data range ---
    If hasTotalToDateCol Then
        RefreshTotalRowFormulas wsDst, totalRow, dstFirstDataRow, dstLastDataRow, firstMonthCol - 1, totalToDateCol
    Else
        RefreshTotalRowFormulas wsDst, totalRow, dstFirstDataRow, dstLastDataRow, firstMonthCol, ytdCol
    End If
End Sub

'------------------------------------------------------------------------------
' Find the row (within the first 15 rows/cols) containing a header cell whose
' text includes the given search text.
'------------------------------------------------------------------------------
Private Function FindHeaderRow(ws As Worksheet, searchText As String) As Long
    Dim r As Long, c As Long
    For r = 1 To 15
        For c = 1 To 20
            If InStr(1, CStr(ws.Cells(r, c).value), searchText, vbTextCompare) > 0 Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    FindHeaderRow = 0
End Function

'------------------------------------------------------------------------------
' Find the column on the given header row whose text includes the search text.
'------------------------------------------------------------------------------
Private Function FindHeaderCol(ws As Worksheet, hdrRow As Long, searchText As String) As Long
    Dim c As Long
    For c = 1 To 30
        If InStr(1, CStr(ws.Cells(hdrRow, c).value), searchText, vbTextCompare) > 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

'------------------------------------------------------------------------------
' Find the column on hdrRow whose date value falls in the given month number
' (1 = January ... 12 = December), regardless of year.
'------------------------------------------------------------------------------
Private Function FindColumnForMonthNumber(ws As Worksheet, hdrRow As Long, monthNum As Long) As Long
    Dim c As Long
    For c = 1 To 40
        If IsDate(ws.Cells(hdrRow, c).value) Then
            If Month(CDate(ws.Cells(hdrRow, c).value)) = monthNum Then
                FindColumnForMonthNumber = c
                Exit Function
            End If
        End If
    Next c
    FindColumnForMonthNumber = 0
End Function

'------------------------------------------------------------------------------
' Find the month-date column on ws that matches the month named in the
' source sheet's tab name (e.g. "August" or "August 2026").
'------------------------------------------------------------------------------
Private Function FindMonthColumn(ws As Worksheet, hdrRow As Long, sheetName As String) As Long
    Dim monthNames As Variant
    Dim i As Long, monthNum As Long
    monthNames = Array("January", "February", "March", "April", "May", "June", _
                        "July", "August", "September", "October", "November", "December")

    monthNum = 0
    For i = 0 To 11
        If InStr(1, sheetName, monthNames(i), vbTextCompare) > 0 Then
            monthNum = i + 1
            Exit For
        End If
    Next i
    If monthNum = 0 Then Exit Function

    FindMonthColumn = FindColumnForMonthNumber(ws, hdrRow, monthNum)
End Function

'------------------------------------------------------------------------------
' Find the row whose column A contains "Total" (e.g. "Total Claims From
' January 2026 to December 2026" / "Total Goodwill From January 2026 to
' December 2026"), searching from the first data row down.
'------------------------------------------------------------------------------
Private Function FindTotalRow(ws As Worksheet, startRow As Long) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    For r = startRow To lastR
        If InStr(1, CStr(ws.Cells(r, "A").value), "Total", vbTextCompare) > 0 Then
            FindTotalRow = r
            Exit Function
        End If
    Next r
    FindTotalRow = 0
End Function

'------------------------------------------------------------------------------
' Find the first row between firstRow and lastRow whose column A is empty -
' i.e. an already-built spare row waiting for data, rather than inserting a
' brand new one right above the Total row.
'------------------------------------------------------------------------------
Private Function FindBlankDataRow(ws As Worksheet, firstRow As Long, lastRow As Long) As Long
    Dim r As Long
    For r = firstRow To lastRow
        If Trim(CStr(ws.Cells(r, "A").value)) = "" Then
            FindBlankDataRow = r
            Exit Function
        End If
    Next r
    FindBlankDataRow = 0
End Function

'------------------------------------------------------------------------------
' Updates the GRAPH sheet's "current month" snapshot cells (O2 = Pure Claim
' Cost, P2 = Goodwill Claim Cost) to point at this month's column on each
' ledger's Total row - e.g. 'CLAIMS 2026'!L111 for July becomes
' 'CLAIMS 2026'!M111 for August. Q2 (Total Claim Cost) is left alone since
' it already just sums O2:P2.
'------------------------------------------------------------------------------
Private Sub UpdateGraphSnapshot(monthSheetName As String)
    Dim wsGraph As Worksheet, wsClaims As Worksheet, wsGW As Worksheet
    Dim claimsHdrRow As Long, gwHdrRow As Long, claimsTotalRow As Long, gwTotalRow As Long
    Dim claimsMonthCol As Long, gwMonthCol As Long
    Dim claimsColLetter As String, gwColLetter As String

    On Error Resume Next
    Set wsGraph = ThisWorkbook.Worksheets("GRAPH")
    On Error GoTo 0
    If wsGraph Is Nothing Then Exit Sub   ' no GRAPH sheet in this workbook - skip quietly

    Set wsClaims = ThisWorkbook.Worksheets("CLAIMS 2026")
    Set wsGW = ThisWorkbook.Worksheets("GOODWILL 2026")

    claimsHdrRow = FindHeaderRow(wsClaims, "Order")
    gwHdrRow = FindHeaderRow(wsGW, "Order")
    If claimsHdrRow = 0 Or gwHdrRow = 0 Then Exit Sub

    claimsTotalRow = FindTotalRow(wsClaims, claimsHdrRow + 1)
    gwTotalRow = FindTotalRow(wsGW, gwHdrRow + 1)
    If claimsTotalRow = 0 Or gwTotalRow = 0 Then Exit Sub

    claimsMonthCol = FindMonthColumn(wsClaims, claimsHdrRow, monthSheetName)
    gwMonthCol = FindMonthColumn(wsGW, gwHdrRow, monthSheetName)
    If claimsMonthCol = 0 Or gwMonthCol = 0 Then Exit Sub

    claimsColLetter = Split(wsClaims.Cells(1, claimsMonthCol).Address, "$")(1)
    gwColLetter = Split(wsGW.Cells(1, gwMonthCol).Address, "$")(1)

    wsGraph.Range("O2").Formula = "='CLAIMS 2026'!" & claimsColLetter & claimsTotalRow
    wsGraph.Range("P2").Formula = "='GOODWILL 2026'!" & gwColLetter & gwTotalRow
End Sub

'------------------------------------------------------------------------------
' Rewrite every column's SUM formula on the total row (from startCol to
' endCol) so it covers exactly firstRow:lastRow.
'------------------------------------------------------------------------------
Private Sub RefreshTotalRowFormulas(ws As Worksheet, totalRow As Long, firstRow As Long, lastRow As Long, _
                                     startCol As Long, endCol As Long)
    Dim c As Long, colLetter As String
    For c = startCol To endCol
        colLetter = Split(ws.Cells(1, c).Address, "$")(1)
        ws.Cells(totalRow, c).Formula = "=SUM(" & colLetter & firstRow & ":" & colLetter & lastRow & ")"
    Next c
End Sub


