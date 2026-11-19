Attribute VB_Name = "modSalesMargin"
'===============================================================================
' modSalesMargin: invoiced sales CMII (contribution margin II) report
'
' Run with the invoiced-sales export (grouped view: detail lines plus one total
' line per order) active. Builds:
'   CMII&%   - one row per order: plant, channel, customer, market area,
'              segment, net sales, COGS (sum of the 9 cost columns), CMII, CMII%
'   Summary  - hidden sheet with the same numbers as the pop-ups, plus charts
' and shows pop-ups: overview, losses by market area / segment, margins by
' customer type, market area and segment, low/high margins, and progress
' against the monthly invoicing target with the rand needed per working day.
'
' Orders whose number contains a letter (internal adjustments) are excluded.
' Orders where COGS is 1000%+ of sales (usually partial invoicing) are kept in
' the table but left out of every margin average.
'===============================================================================

Private Const SISTER_CO_TAG As String = "Acme"   ' customer names containing this = group companies

Sub CreateCMIISheet()


    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim i As Long
    Dim tbl As ListObject

    ' -- Source sheet ----------------------------------------------
    ' Grouped View 2 layout: headers on row 1, data starts row 2.
    ' Each order has detail rows plus ONE subtotal/total row.
    ' Col A  = Order
    ' Col E  = Net Sales Revenue
    ' Col F  = Material for Sales
    ' Col G  = Variable production costs
    ' Col H  = Other special direct costs
    ' Col I  = Reject and rework
    ' Col J  = Fixed production costs
    ' Col K  = Production overhead costs (%)
    ' Col L  = Development Costs (WTE) from Sales
    ' Col M  = Development Costs (WTE) R&D
    ' Col N  = Material overhead costs (%)
    ' Col R  = Plant
    ' Col S  = Segment_Market Area
    ' Col T  = Market Area
    ' Col U  = Distr. Channel
    ' Col V  = Descr. Deb. (Customer)
    ' COGS is NOT pulled from the source's own "Cost of goods sold"
    ' column - it's computed as the sum of columns F:N (Material,
    ' Variable production, Other special direct, Reject and rework,
    ' Fixed production, Production overhead %, both Development
    ' Costs (WTE) columns, and Material overhead %).
    Set srcWs = ActiveSheet

    ' -- Find last non-blank row in source Col A (starting row 2) --
    lastRow = 2
    Do While srcWs.Cells(lastRow, 1).value <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1   ' last row with data (excludes the blank-order grand total row, if present)

    ' -- Flag total rows BEFORE filling blanks ----------------------
    ' Plant (col R) is blank ONLY on each order's total row in this
    ' export - Market Area / Segment_Market Area can legitimately be
    ' blank on ordinary detail lines here too, so they are NOT used
    ' as the flag.
    Dim isTotalRow() As Boolean
    ReDim isTotalRow(2 To lastRow)

    For i = 2 To lastRow
        isTotalRow(i) = (srcWs.Cells(i, 18).value = "")
    Next i

    ' -- Fill blank Plant / Segment_Market Area / Market Area / -----
    ' -- Distr. Channel / Descr. Deb. cells with the value above -----
    ' Only the TOTAL rows' blanks end up in the destination table
    ' (detail rows are discarded), and a total row's "row above" is
    ' always the last detail line of its OWN order.
    Dim col As Long
    For col = 18 To 22   ' R, S, T, U, V
        For i = 2 To lastRow
            If srcWs.Cells(i, col).value = "" Then
                srcWs.Cells(i, col).value = srcWs.Cells(i - 1, col).value
            End If
        Next i
    Next col

    ' -- Delete existing CMII&% sheet if it exists -----------------
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("CMII&%").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' -- Create destination sheet ----------------------------------
    Set destWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    destWs.Name = "CMII&%"

    ' -- Headers in Row 1 -------------------------------------------
    destWs.Range("A1").value = "Plant"
    destWs.Range("B1").value = "Distr. Channel"
    destWs.Range("C1").value = "Customer"
    destWs.Range("D1").value = "Market Area"
    destWs.Range("E1").value = "Segment_Market Area"
    destWs.Range("F1").value = "Order Number"
    destWs.Range("G1").value = "Net Sales"
    destWs.Range("H1").value = "COGS"
    destWs.Range("I1").value = "CMII"
    destWs.Range("J1").value = "CMII%"
    destWs.Range("K1").value = "Comments"

    ' -- Copy ONLY the flagged total rows, excluding any Order -----
    ' -- Number that contains a letter -------------------------------
    ' COGS is computed as a plain value (sum of columns F:N - the 9
    ' cost line items) rather than a live formula, to avoid Excel's
    ' "Formula omits adjacent cells" warning.
    Dim cogs As Double
    Dim netSales As Double
    Dim cmii As Double
    Dim cmPct As Double
    Dim lossCount As Long
    Dim noCostCount As Long
    Dim lossByArea As Object
    Dim mktArea As String
    Set lossByArea = CreateObject("Scripting.Dictionary")

    ' Running totals of each cost component's % of Net Sales, but only
    ' for LOSS orders (cmPct < 0) - used to report which cost item is,
    ' on average, the biggest contributor on the orders that lost money.
    Dim costLabels(1 To 9) As String
    Dim costSum(1 To 9) As Double
    Dim costCount(1 To 9) As Long
    Dim ci As Long
    costLabels(1) = "Material for Sales"
    costLabels(2) = "Variable production costs"
    costLabels(3) = "Other special direct costs"
    costLabels(4) = "Reject and rework"
    costLabels(5) = "Fixed production costs"
    costLabels(6) = "Production overhead costs (%)"
    costLabels(7) = "Development Costs (WTE) from Sales"
    costLabels(8) = "Development Costs (WTE) R&D"
    costLabels(9) = "Material overhead costs (%)"

    ' Margin (CMII%) accumulators - only orders with BOTH non-zero Net
    ' Sales AND non-zero COGS count here (excludes no-sales, no-cost-
    ' with-sales, and no-cost-no-sales orders alike).
    Dim grpSum As Double, grpCount As Long          ' sister (group) companies
    Dim sparesSum As Double, sparesCount As Long
    Dim otherSum As Double, otherCount As Long   ' Other Parties (not group companies, not Spares)
    Dim custName As String

    Dim areaMarginSum As Object, areaMarginCount As Object
    Set areaMarginSum = CreateObject("Scripting.Dictionary")
    Set areaMarginCount = CreateObject("Scripting.Dictionary")
    Dim marginAreaKey As String

    Dim seSum As Double, seCount As Long   ' Service (segment "SE")
    Dim vaSum As Double, vaCount As Long   ' Valves (segment "VA")
    Dim seLossCount As Long, vaLossCount As Long
    Dim segKey As String
    Dim isOutlier As Boolean
    Dim isGroupCo As Boolean, isSpares As Boolean
    Dim totalSales As Double
    Dim zeroSalesCount As Long
    Dim zeroCostZeroSalesCount As Long

    ' Low margin (< 10%) and high margin (> 60%) counts, broken down by
    ' Market Area and by Segment - same qualifying orders as Margin Summary
    Dim lowMarginCount As Long, highMarginCount As Long
    Dim lowMarginByArea As Object, highMarginByArea As Object
    Set lowMarginByArea = CreateObject("Scripting.Dictionary")
    Set highMarginByArea = CreateObject("Scripting.Dictionary")
    Dim lowSE As Long, lowVA As Long, highSE As Long, highVA As Long

    destRow = 2
    For i = 2 To lastRow
        If isTotalRow(i) Then
            If Not ContainsLetter(CStr(srcWs.Cells(i, 1).value)) Then
                cogs = NZ(srcWs.Cells(i, 6).value) + NZ(srcWs.Cells(i, 7).value) + _
                       NZ(srcWs.Cells(i, 8).value) + NZ(srcWs.Cells(i, 9).value) + _
                       NZ(srcWs.Cells(i, 10).value) + NZ(srcWs.Cells(i, 11).value) + _
                       NZ(srcWs.Cells(i, 12).value) + NZ(srcWs.Cells(i, 13).value) + _
                       NZ(srcWs.Cells(i, 14).value)

                destWs.Cells(destRow, 1).value = srcWs.Cells(i, 18).value   ' Plant
                destWs.Cells(destRow, 2).value = srcWs.Cells(i, 21).value   ' Distr. Channel
                destWs.Cells(destRow, 3).value = srcWs.Cells(i, 22).value   ' Customer
                destWs.Cells(destRow, 4).value = srcWs.Cells(i, 20).value   ' Market Area
                destWs.Cells(destRow, 5).value = srcWs.Cells(i, 19).value   ' Segment_Market Area
                destWs.Cells(destRow, 6).value = srcWs.Cells(i, 1).value    ' Order Number
                destWs.Cells(destRow, 7).value = srcWs.Cells(i, 5).value    ' Net Sales
                destWs.Cells(destRow, 8).value = cogs                       ' COGS (computed value)

                netSales = NZ(srcWs.Cells(i, 5).value)
                cmii = netSales - cogs
                totalSales = totalSales + netSales
                If netSales <> 0 Then
                    cmPct = cmii / netSales
                Else
                    cmPct = 0   ' matches the IFERROR(...,0) fallback used in the CMII% formula
                End If

                ' Partial invoicing can make Sales and COGS wildly mismatched
                ' for one order (e.g. Sales=40, COGS=50288). Whenever COGS is
                ' 1000%+ of Net Sales (either direction), that whole order is
                ' excluded from every margin calculation below.
                isOutlier = False
                If netSales <> 0 Then
                    If Abs(cogs / netSales) >= 10 Then isOutlier = True
                End If

                If cmPct < 0 Then
                    lossCount = lossCount + 1   ' negative CMII%

                    mktArea = CStr(srcWs.Cells(i, 20).value)
                    If mktArea = "" Then mktArea = "(blank)"
                    If lossByArea.Exists(mktArea) Then
                        lossByArea(mktArea) = lossByArea(mktArea) + 1
                    Else
                        lossByArea.Add mktArea, 1
                    End If

                    segKey = CStr(srcWs.Cells(i, 19).value)
                    If segKey = "SE" Then
                        seLossCount = seLossCount + 1
                    ElseIf segKey = "VA" Then
                        vaLossCount = vaLossCount + 1
                    End If

                    ' accumulate each cost column's % of Net Sales for this loss
                    ' order - skipped entirely for outlier orders (see above)
                    If Not isOutlier Then
                        For ci = 1 To 9
                            costSum(ci) = costSum(ci) + NZ(srcWs.Cells(i, 5 + ci).value) / netSales
                            costCount(ci) = costCount(ci) + 1
                        Next ci
                    End If
                End If

                If netSales <> 0 And cogs = 0 Then
                    noCostCount = noCostCount + 1   ' has sales, but zero cost
                End If

                If netSales = 0 And cogs = 0 Then
                    zeroCostZeroSalesCount = zeroCostZeroSalesCount + 1   ' both zero - counted separately
                ElseIf netSales = 0 Then
                    zeroSalesCount = zeroSalesCount + 1   ' zero sales, but has cost
                End If

                ' Margin accumulators - only orders with real Sales AND real
                ' COGS, and not an outlier order, count here
                If netSales <> 0 And cogs <> 0 And Not isOutlier Then
                    custName = CStr(srcWs.Cells(i, 22).value)
                    isGroupCo = (InStr(1, custName, SISTER_CO_TAG, vbTextCompare) > 0)
                    isSpares = (val(srcWs.Cells(i, 21).value) = 2)

                    If isGroupCo Then
                        grpSum = grpSum + cmPct
                        grpCount = grpCount + 1
                    End If
                    If isSpares Then
                        sparesSum = sparesSum + cmPct
                        sparesCount = sparesCount + 1
                    End If
                    If Not isGroupCo And Not isSpares Then
                        otherSum = otherSum + cmPct
                        otherCount = otherCount + 1
                    End If

                    marginAreaKey = CStr(srcWs.Cells(i, 20).value)
                    If marginAreaKey = "" Then marginAreaKey = "(blank)"
                    If areaMarginSum.Exists(marginAreaKey) Then
                        areaMarginSum(marginAreaKey) = areaMarginSum(marginAreaKey) + cmPct
                        areaMarginCount(marginAreaKey) = areaMarginCount(marginAreaKey) + 1
                    Else
                        areaMarginSum.Add marginAreaKey, cmPct
                        areaMarginCount.Add marginAreaKey, 1
                    End If

                    segKey = CStr(srcWs.Cells(i, 19).value)
                    If segKey = "SE" Then
                        seSum = seSum + cmPct
                        seCount = seCount + 1
                    ElseIf segKey = "VA" Then
                        vaSum = vaSum + cmPct
                        vaCount = vaCount + 1
                    End If

                    If cmPct >= 0 And cmPct < 0.1 Then
                        lowMarginCount = lowMarginCount + 1
                        If lowMarginByArea.Exists(marginAreaKey) Then
                            lowMarginByArea(marginAreaKey) = lowMarginByArea(marginAreaKey) + 1
                        Else
                            lowMarginByArea.Add marginAreaKey, 1
                        End If
                        If segKey = "SE" Then
                            lowSE = lowSE + 1
                        ElseIf segKey = "VA" Then
                            lowVA = lowVA + 1
                        End If
                    ElseIf cmPct > 0.6 Then
                        highMarginCount = highMarginCount + 1
                        If highMarginByArea.Exists(marginAreaKey) Then
                            highMarginByArea(marginAreaKey) = highMarginByArea(marginAreaKey) + 1
                        Else
                            highMarginByArea.Add marginAreaKey, 1
                        End If
                        If segKey = "SE" Then
                            highSE = highSE + 1
                        ElseIf segKey = "VA" Then
                            highVA = highVA + 1
                        End If
                    End If
                End If

                destRow = destRow + 1
            End If
        End If
    Next i

    Dim dataLastRow As Long
    dataLastRow = destRow - 1   ' last row of pasted data

    If dataLastRow < 2 Then
        MsgBox "No total rows were found (Plant was never blank in column R). " & _
               "Check that this is the grouped export sheet with Order in col A, " & _
               "Net Sales in col E, Plant in col R.", vbExclamation
        Exit Sub
    End If

    ' -- CMII & CMII% formulas ---------------------------------------
    Dim r As Long
    For r = 2 To dataLastRow
        destWs.Cells(r, 9).Formula = "=IFERROR(SUM(G" & r & "-H" & r & "),0)"
        destWs.Cells(r, 10).Formula = "=IFERROR(SUM(I" & r & "/G" & r & "),0)"
    Next r

    ' -- Number Formatting (data rows only - totals row handled below) --
    Dim acctFormat As String
    acctFormat = "_-R* #,##0_-;-R* #,##0_-;_-R* ""-""_-;_-@_-"

    destWs.Range("G2:G" & dataLastRow).NumberFormat = acctFormat
    destWs.Range("H2:H" & dataLastRow).NumberFormat = acctFormat
    destWs.Range("I2:I" & dataLastRow).NumberFormat = acctFormat
    destWs.Range("J2:J" & dataLastRow).NumberFormat = "0%"

    ' -- Format as Table, with native Total Row enabled -------------
    Dim tblRange As Range
    Set tblRange = destWs.Range("A1:K" & dataLastRow)
    Set tbl = destWs.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "CMII_Table"
    tbl.TableStyle = "TableStyleLight9"

    tbl.ShowTotals = True

    Dim totRow As Long
    totRow = dataLastRow + 1

    destWs.Cells(totRow, 1).value = "Total"
    destWs.Cells(totRow, 7).Formula = "=SUBTOTAL(109,G2:G" & dataLastRow & ")"   ' Net Sales
    destWs.Cells(totRow, 8).Formula = "=SUBTOTAL(109,H2:H" & dataLastRow & ")"   ' COGS
    destWs.Cells(totRow, 9).Formula = "=SUBTOTAL(109,I2:I" & dataLastRow & ")"   ' CMII
    ' CMII% is left blank in the totals row - a straight sum/average of
    ' percentages would be misleading; if you want an overall CMII%,
    ' that's (Total CMII / Total Net Sales), not an average of rows.

    destWs.Cells(totRow, 1).Font.Bold = True
    destWs.Cells(totRow, 7).Font.Bold = True
    destWs.Cells(totRow, 8).Font.Bold = True
    destWs.Cells(totRow, 9).Font.Bold = True

    destWs.Range("G" & totRow).NumberFormat = acctFormat
    destWs.Range("H" & totRow).NumberFormat = acctFormat
    destWs.Range("I" & totRow).NumberFormat = acctFormat

    ' -- Column widths --------------------------------------------
    destWs.Columns("A").ColumnWidth = 10
    destWs.Columns("B").ColumnWidth = 14
    destWs.Columns("C").ColumnWidth = 30
    destWs.Columns("D").ColumnWidth = 14
    destWs.Columns("E").ColumnWidth = 20
    destWs.Columns("F").ColumnWidth = 18
    destWs.Columns("G").ColumnWidth = 16
    destWs.Columns("H").ColumnWidth = 16
    destWs.Columns("I").ColumnWidth = 16
    destWs.Columns("J").ColumnWidth = 10
    destWs.Columns("K").ColumnWidth = 20

    ' -- Center alignment: Plant, Distr. Channel, Market Area, -------
    ' -- Segment_Market Area ------------------------------------------
    destWs.Range("A2:A" & totRow).HorizontalAlignment = xlCenter
    destWs.Range("B2:B" & totRow).HorizontalAlignment = xlCenter
    destWs.Range("D2:D" & totRow).HorizontalAlignment = xlCenter
    destWs.Range("E2:E" & totRow).HorizontalAlignment = xlCenter

    ' -- Font -----------------------------------------------------
    destWs.Cells.Font.Name = "Arial"

    Dim areaBreakdown As String
    Dim k As Variant
    For Each k In SortKeysDesc(lossByArea)
        areaBreakdown = areaBreakdown & vbCrLf & "  " & k & ": " & lossByArea(k)
    Next k

    ' -- Which cost item contributes most to losses, on average? -----
    Dim topContributorLine As String
    If lossCount > 0 Then
        Dim maxAvg As Double
        Dim maxIdx As Long
        Dim avgVal As Double
        maxAvg = -1
        For ci = 1 To 9
            If costCount(ci) > 0 Then
                avgVal = costSum(ci) / costCount(ci)
                If avgVal > maxAvg Then
                    maxAvg = avgVal
                    maxIdx = ci
                End If
            End If
        Next ci
        topContributorLine = "Top % of Sales-Losses: " & costLabels(maxIdx) & " (" & Format(maxAvg, "0.0%") & ")"
    Else
        topContributorLine = "Top % of Sales-Losses: n/a"
    End If

    Dim grpLine As String, otherLine As String, sparesLine As String
    If grpCount > 0 Then
        grpLine = "Sister Companies Margin (" & grpCount & " orders): " & Format(grpSum / grpCount, "0.0%")
    Else
        grpLine = "Sister Companies Margin (0 orders): n/a"
    End If
    If otherCount > 0 Then
        otherLine = "Other Parties Margin (" & otherCount & " orders): " & Format(otherSum / otherCount, "0.0%")
    Else
        otherLine = "Other Parties Margin (0 orders): n/a"
    End If
    If sparesCount > 0 Then
        sparesLine = "Spares Margin (" & sparesCount & " orders) : " & Format(sparesSum / sparesCount, "0.0%")
    Else
        sparesLine = "Spares Margin (0 orders) : n/a"
    End If

    ' Area keys already carry their order count, so both the MsgBox
    ' breakdown and the Summary sheet rows (via WriteDictRows) show it
    ' without duplicating the lookup.
    Dim areaMarginAvg As Object
    Set areaMarginAvg = CreateObject("Scripting.Dictionary")
    For Each k In areaMarginSum.keys
        areaMarginAvg.Add k & " (" & areaMarginCount(k) & " orders)", areaMarginSum(k) / areaMarginCount(k)
    Next k

    Dim areaMarginBreakdown As String
    For Each k In SortKeysDesc(areaMarginAvg)
        areaMarginBreakdown = areaMarginBreakdown & vbCrLf & "  " & k & ": " & Format(areaMarginAvg(k), "0.0%")
    Next k

    Dim seLine As String, vaLine As String
    If seCount > 0 Then
        seLine = "Service (SE) (" & seCount & " orders): " & Format(seSum / seCount, "0.0%")
    Else
        seLine = "Service (SE) (0 orders): n/a"
    End If
    If vaCount > 0 Then
        vaLine = "Valves (VA) (" & vaCount & " orders): " & Format(vaSum / vaCount, "0.0%")
    Else
        vaLine = "Valves (VA) (0 orders): n/a"
    End If

    Dim allMarginAreas As Object
    Set allMarginAreas = CreateObject("Scripting.Dictionary")
    For Each k In lowMarginByArea.keys
        If Not allMarginAreas.Exists(k) Then allMarginAreas.Add k, 1
    Next k
    For Each k In highMarginByArea.keys
        If Not allMarginAreas.Exists(k) Then allMarginAreas.Add k, 1
    Next k

    ' Combined (Low + High) total per area, used purely to sort the
    ' breakdown largest to smallest - highest-activity area first.
    Dim combinedTotal As Object
    Set combinedTotal = CreateObject("Scripting.Dictionary")
    Dim lv As Long, hv As Long
    For Each k In allMarginAreas.keys
        lv = 0: hv = 0
        If lowMarginByArea.Exists(k) Then lv = lowMarginByArea(k)
        If highMarginByArea.Exists(k) Then hv = highMarginByArea(k)
        combinedTotal.Add k, lv + hv
    Next k

    ' One line per area with both counts together ("Low / High") rather
    ' than two padded columns - a MsgBox font is proportional, so
    ' character-count padding never lines up reliably (PCC and WAT don't
    ' render the same pixel width), which is what was drifting.
    Dim lowHighBreakdown As String
    For Each k In SortKeysDesc(combinedTotal)
        lv = 0: hv = 0
        If lowMarginByArea.Exists(k) Then lv = lowMarginByArea(k)
        If highMarginByArea.Exists(k) Then hv = highMarginByArea(k)
        lowHighBreakdown = lowHighBreakdown & vbCrLf & "  " & k & ":  " & lv & " Low  /  " & hv & " High"
    Next k

    Const DIV As String = "------------------------------"

    ' Split into several small popups instead of one giant MsgBox - a
    ' single MsgBox has no scrollbar, so once the text got this long it
    ' was silently clipping content off the bottom (missing market areas,
    ' missing NB notes) rather than showing everything.

    Dim nbText As String
    nbText = vbCrLf & vbCrLf & DIV & vbCrLf & _
             "NB: Data includes Partially and Fully Invoiced Jobs," & vbCrLf & _
             "therefore the Margins summary is subject to further investigation" & vbCrLf & vbCrLf & _
             "Internal adjustment orders (e.g. order numbers containing letters, e.g. ADJ...) are excluded" & vbCrLf & _
             "from this analysis"

    Dim msg1 As String
    msg1 = "Sheet Summary" & vbCrLf & DIV & vbCrLf
    msg1 = msg1 & "Orders: " & (dataLastRow - 1) & vbCrLf
    msg1 = msg1 & "Sales Total: R " & Format(totalSales, "#,##0") & vbCrLf
    msg1 = msg1 & "0-Revenue Orders: " & zeroSalesCount & vbCrLf
    msg1 = msg1 & "0-Costs 0-Sales: " & zeroCostZeroSalesCount & vbCrLf
    msg1 = msg1 & "No Cost- Orders: " & noCostCount & vbCrLf
    msg1 = msg1 & "Losses: " & lossCount

    Dim msg2 As String
    msg2 = "Losses by Market Area" & vbCrLf & DIV & areaBreakdown & vbCrLf & vbCrLf
    msg2 = msg2 & "Losses by Segment" & vbCrLf & DIV & vbCrLf
    If seLossCount >= vaLossCount Then
        msg2 = msg2 & "Service (SE): " & seLossCount & vbCrLf
        msg2 = msg2 & "Valves (VA): " & vaLossCount
    Else
        msg2 = msg2 & "Valves (VA): " & vaLossCount & vbCrLf
        msg2 = msg2 & "Service (SE): " & seLossCount
    End If

    Dim msg3 As String
    msg3 = "Margin Summary" & vbCrLf & DIV & vbCrLf
    msg3 = msg3 & "(excludes 0-cost, 0-sales, and outlier orders" & vbCrLf
    msg3 = msg3 & "where COGS is 1000%+ of Net Sales)" & vbCrLf & vbCrLf
    msg3 = msg3 & topContributorLine & vbCrLf
    msg3 = msg3 & grpLine & vbCrLf
    msg3 = msg3 & otherLine & vbCrLf
    msg3 = msg3 & sparesLine & vbCrLf & vbCrLf
    msg3 = msg3 & "Market Area Margins" & vbCrLf & DIV & areaMarginBreakdown & vbCrLf & vbCrLf
    Dim seVal As Double, vaVal As Double
    seVal = IIf(seCount > 0, seSum / seCount, -999999)
    vaVal = IIf(vaCount > 0, vaSum / vaCount, -999999)
    If seVal >= vaVal Then
        msg3 = msg3 & seLine & vbCrLf & vaLine
    Else
        msg3 = msg3 & vaLine & vbCrLf & seLine
    End If

    Dim msg4 As String
    msg4 = "Low Margins (<10%): " & lowMarginCount & vbCrLf
    msg4 = msg4 & "High Margins (>60%): " & highMarginCount & vbCrLf & vbCrLf
    msg4 = msg4 & "By Market Area (Low / High):" & lowHighBreakdown & vbCrLf & vbCrLf
    If (lowSE + highSE) >= (lowVA + highVA) Then
        msg4 = msg4 & "Service (SE):  " & lowSE & " Low  /  " & highSE & " High" & vbCrLf
        msg4 = msg4 & "Valves (VA):  " & lowVA & " Low  /  " & highVA & " High"
    Else
        msg4 = msg4 & "Valves (VA):  " & lowVA & " Low  /  " & highVA & " High" & vbCrLf
        msg4 = msg4 & "Service (SE):  " & lowSE & " Low  /  " & highSE & " High"
    End If

    ' -- Build the (hidden) Summary sheet - same numbers as the popups --
    ' -- below, but as a real sheet so Low/High Margins can use actual --
    ' -- worksheet columns instead of text padding. ----------------------
    Dim sumWs As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Summary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set sumWs = ThisWorkbook.Sheets.Add(After:=destWs)
    sumWs.Name = "Summary"
    sumWs.Columns("A").ColumnWidth = 40
    sumWs.Columns("B").ColumnWidth = 14
    sumWs.Columns("C").ColumnWidth = 14
    sumWs.Cells.Font.Name = "Arial"

    Dim sr As Long
    sr = 1

    sumWs.Cells(sr, 1).value = "Sheet Summary"
    sumWs.Cells(sr, 1).Font.Bold = True
    sumWs.Cells(sr, 1).Font.Size = 14
    sr = sr + 2

    WriteSummaryRow sumWs, sr, "Orders", dataLastRow - 1, False
    WriteSummaryRow sumWs, sr, "Sales Total", totalSales, False, "R #,##0"
    WriteSummaryRow sumWs, sr, "0-Revenue Orders", zeroSalesCount, False
    WriteSummaryRow sumWs, sr, "0-Costs 0-Sales", zeroCostZeroSalesCount, False
    WriteSummaryRow sumWs, sr, "No Cost- Orders", noCostCount, False
    WriteSummaryRow sumWs, sr, "Losses", lossCount, False
    sr = sr + 1

    WriteSummarySection sumWs, sr, "Losses by Market Area"
    Dim lossAreaStart As Long, lossAreaEnd As Long
    lossAreaStart = sr
    WriteDictRows sumWs, sr, lossByArea, False
    lossAreaEnd = sr - 1
    sr = sr + 1

    WriteSummarySection sumWs, sr, "Losses by Segment"
    If seLossCount >= vaLossCount Then
        WriteSummaryRow sumWs, sr, "Service (SE)", seLossCount, False
        WriteSummaryRow sumWs, sr, "Valves (VA)", vaLossCount, False
    Else
        WriteSummaryRow sumWs, sr, "Valves (VA)", vaLossCount, False
        WriteSummaryRow sumWs, sr, "Service (SE)", seLossCount, False
    End If
    sr = sr + 1

    WriteSummarySection sumWs, sr, "Margin Summary"
    sumWs.Cells(sr, 1).value = "(excludes 0-cost, 0-sales, and outlier orders where COGS is 1000%+ of Net Sales)"
    sumWs.Cells(sr, 1).Font.Italic = True
    sr = sr + 2
    sumWs.Cells(sr, 1).value = topContributorLine
    sr = sr + 1
    WriteSummaryRow sumWs, sr, "Sister Companies Margin (" & grpCount & " orders)", IIf(grpCount > 0, grpSum / grpCount, 0), True
    WriteSummaryRow sumWs, sr, "Other Parties Margin (" & otherCount & " orders)", IIf(otherCount > 0, otherSum / otherCount, 0), True
    WriteSummaryRow sumWs, sr, "Spares Margin (" & sparesCount & " orders)", IIf(sparesCount > 0, sparesSum / sparesCount, 0), True
    sr = sr + 1

    WriteSummarySection sumWs, sr, "Market Area Margins"
    Dim areaMarginStart As Long, areaMarginEnd As Long
    areaMarginStart = sr
    WriteDictRows sumWs, sr, areaMarginAvg, True
    areaMarginEnd = sr - 1
    sr = sr + 1
    If seVal >= vaVal Then
        WriteSummaryRow sumWs, sr, "Service (SE) (" & seCount & " orders)", IIf(seCount > 0, seSum / seCount, 0), True
        WriteSummaryRow sumWs, sr, "Valves (VA) (" & vaCount & " orders)", IIf(vaCount > 0, vaSum / vaCount, 0), True
    Else
        WriteSummaryRow sumWs, sr, "Valves (VA) (" & vaCount & " orders)", IIf(vaCount > 0, vaSum / vaCount, 0), True
        WriteSummaryRow sumWs, sr, "Service (SE) (" & seCount & " orders)", IIf(seCount > 0, seSum / seCount, 0), True
    End If
    sr = sr + 1

    WriteSummarySection sumWs, sr, "Low / High Margins (<10% / >60%)"
    sumWs.Cells(sr, 2).value = "Low"
    sumWs.Cells(sr, 3).value = "High"
    sumWs.Cells(sr, 2).Font.Bold = True
    sumWs.Cells(sr, 3).Font.Bold = True
    sr = sr + 1
    WriteSummaryRow2 sumWs, sr, "Total", lowMarginCount, highMarginCount
    sr = sr + 1
    Dim lowHighStart As Long, lowHighEnd As Long
    lowHighStart = sr
    Dim k2 As Variant
    For Each k2 In SortKeysDesc(combinedTotal)
        Dim lv2 As Long, hv2 As Long
        lv2 = 0: hv2 = 0
        If lowMarginByArea.Exists(k2) Then lv2 = lowMarginByArea(k2)
        If highMarginByArea.Exists(k2) Then hv2 = highMarginByArea(k2)
        WriteSummaryRow2 sumWs, sr, "  " & k2, lv2, hv2
    Next k2
    lowHighEnd = sr - 1
    sr = sr + 1
    If (lowSE + highSE) >= (lowVA + highVA) Then
        WriteSummaryRow2 sumWs, sr, "Service (SE)", lowSE, highSE
        WriteSummaryRow2 sumWs, sr, "Valves (VA)", lowVA, highVA
    Else
        WriteSummaryRow2 sumWs, sr, "Valves (VA)", lowVA, highVA
        WriteSummaryRow2 sumWs, sr, "Service (SE)", lowSE, highSE
    End If
    sr = sr + 2

    sumWs.Cells(sr, 1).value = "NB: Data includes Partially and Fully Invoiced Jobs, therefore the Margins summary is subject to further investigation"
    sumWs.Cells(sr, 1).Font.Italic = True
    sumWs.Cells(sr, 1).WrapText = True
    sr = sr + 1
    sumWs.Cells(sr, 1).value = "Internal adjustment orders (e.g. order numbers containing letters, e.g. ADJ...) are excluded from this analysis"
    sumWs.Cells(sr, 1).Font.Italic = True
    sumWs.Cells(sr, 1).WrapText = True

    sumWs.Columns("A").AutoFit
    If sumWs.Columns("A").ColumnWidth < 40 Then sumWs.Columns("A").ColumnWidth = 40

    ' -- Charts, placed to the right of the text (columns F onward) -----
    Dim chLeft As Double, chWidth As Double, chHeight As Double, chTop As Double
    chLeft = sumWs.Columns("F").Left
    chWidth = 380
    chHeight = 220

    ' Chart 1: Losses by Market Area
    If lossAreaEnd >= lossAreaStart Then
        Dim ch1 As ChartObject
        Set ch1 = sumWs.ChartObjects.Add(chLeft, 10, chWidth, chHeight)
        With ch1.Chart
            .SetSourceData sumWs.Range(sumWs.Cells(lossAreaStart, 1), sumWs.Cells(lossAreaEnd, 2))
            .ChartType = xlColumnClustered
            .HasTitle = True
            .ChartTitle.Text = "Losses by Market Area"
            .HasLegend = False
        End With
    End If

    ' Chart 2: Market Area Margins
    If areaMarginEnd >= areaMarginStart Then
        Dim ch2 As ChartObject
        Set ch2 = sumWs.ChartObjects.Add(chLeft, 10 + chHeight + 10, chWidth, chHeight)
        With ch2.Chart
            .SetSourceData sumWs.Range(sumWs.Cells(areaMarginStart, 1), sumWs.Cells(areaMarginEnd, 2))
            .ChartType = xlColumnClustered
            .HasTitle = True
            .ChartTitle.Text = "Market Area Margins"
            .HasLegend = False
            .Axes(xlValue).TickLabels.NumberFormat = "0%"
        End With
    End If

    ' Chart 3: (removed - Low vs High Margins by Market Area chart was
    ' taken out; the section itself still lists the numbers as text)

    sumWs.Visible = xlSheetHidden   ' regular hidden - right-click any tab > Unhide to reveal
    destWs.Activate

    MsgBox msg1 & nbText, vbInformation, "Summary 1 of 4 - Overview"
    MsgBox msg2 & nbText, vbInformation, "Summary 2 of 4 - Losses"
    MsgBox msg3 & nbText, vbInformation, "Summary 3 of 4 - Margins"
    MsgBox msg4 & nbText, vbInformation, "Summary 4 of 4 - Low vs High Margins"

    ' -- Monthly target progress -------------------------------------
    ' Skipped (with an explanatory note instead) if totalSales is more
    ' than 15% over target - that's a sign the data spans more than one
    ' settlement period rather than a single month, which would make a
    ' days-left-in-month calculation meaningless.
    Const MONTHLY_TARGET As Double = 60000000      ' monthly invoicing target (R)

    If totalSales > MONTHLY_TARGET * 1.15 Then
        MsgBox "Target progress not shown: total (R " & Format(totalSales, "#,##0") & _
               ") is 15%+ over the monthly target (R " & Format(MONTHLY_TARGET, "#,##0") & ")." & _
               vbCrLf & vbCrLf & _
               "Data probably spans multiple settlement periods rather than one month.", vbExclamation
    Else
        Dim balance As Double
        balance = MONTHLY_TARGET - totalSales

        Dim monthEnd As Date
        monthEnd = DateSerial(Year(Date), Month(Date) + 1, 0)

        Dim workDaysLeft As Long
        Dim d As Date
        For d = Date To monthEnd
            If Weekday(d, vbMonday) <= 5 Then workDaysLeft = workDaysLeft + 1
        Next d

        Dim msg5 As String
        msg5 = "As of: " & Format(Date, "dd mmmm yyyy") & vbCrLf & vbCrLf
        msg5 = msg5 & "Target: R " & Format(MONTHLY_TARGET, "#,##0") & vbCrLf
        msg5 = msg5 & "Invoiced: R " & Format(totalSales, "#,##0") & vbCrLf
        msg5 = msg5 & "Completion: " & Format(totalSales / MONTHLY_TARGET, "0.0%") & vbCrLf & vbCrLf
        msg5 = msg5 & "Balance to Target: R " & Format(balance, "#,##0") & vbCrLf
        msg5 = msg5 & "Working Days Left (to " & Format(monthEnd, "dd mmm") & "): " & workDaysLeft & vbCrLf

        If workDaysLeft > 0 Then
            msg5 = msg5 & "Required per Working Day: R " & Format(balance / workDaysLeft, "#,##0")
        Else
            msg5 = msg5 & "Required per Working Day: n/a (no working days left)"
        End If

        MsgBox msg5, vbInformation, "Target Progress - " & Format(Date, "mmmm yyyy")
    End If

End Sub

Private Function SortKeysDesc(d As Object) As Variant
    Dim n As Long
    n = d.Count
    If n = 0 Then
        SortKeysDesc = Array()
        Exit Function
    End If

    Dim keysArr() As Variant
    Dim valsArr() As Double
    ReDim keysArr(0 To n - 1)
    ReDim valsArr(0 To n - 1)

    Dim idx As Long
    idx = 0
    Dim kk As Variant
    For Each kk In d.keys
        keysArr(idx) = kk
        valsArr(idx) = CDbl(d(kk))
        idx = idx + 1
    Next kk

    Dim a As Long, b As Long
    Dim tmpV As Double, tmpK As Variant
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If valsArr(b) < valsArr(b + 1) Then
                tmpV = valsArr(b): valsArr(b) = valsArr(b + 1): valsArr(b + 1) = tmpV
                tmpK = keysArr(b): keysArr(b) = keysArr(b + 1): keysArr(b + 1) = tmpK
            End If
        Next b
    Next a

    SortKeysDesc = keysArr
End Function

Private Sub WriteSummaryRow(ws As Worksheet, ByRef r As Long, label As String, value As Variant, _
                            isPercent As Boolean, Optional fmt As String = "")
    ws.Cells(r, 1).value = label
    ws.Cells(r, 2).value = value
    If isPercent Then
        ws.Cells(r, 2).NumberFormat = "0.0%"
    ElseIf fmt <> "" Then
        ws.Cells(r, 2).NumberFormat = fmt
    End If
    r = r + 1
End Sub

Private Sub WriteSummaryRow2(ws As Worksheet, ByRef r As Long, label As String, valLow As Variant, valHigh As Variant)
    ws.Cells(r, 1).value = label
    ws.Cells(r, 2).value = valLow
    ws.Cells(r, 3).value = valHigh
    r = r + 1
End Sub

Private Sub WriteSummarySection(ws As Worksheet, ByRef r As Long, title As String)
    ws.Cells(r, 1).value = title
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Interior.Color = RGB(230, 230, 230)
    ws.Cells(r, 2).Interior.Color = RGB(230, 230, 230)
    ws.Cells(r, 3).Interior.Color = RGB(230, 230, 230)
    r = r + 1
End Sub

Private Sub WriteDictRows(ws As Worksheet, ByRef r As Long, d As Object, isPercent As Boolean)
    Dim key As Variant
    If d.Count = 0 Then
        ws.Cells(r, 1).value = "  (none)"
        r = r + 1
        Exit Sub
    End If
    For Each key In SortKeysDesc(d)
        ws.Cells(r, 1).value = "  " & key
        ws.Cells(r, 2).value = d(key)
        If isPercent Then ws.Cells(r, 2).NumberFormat = "0.0%"
        r = r + 1
    Next key
End Sub

Private Function ContainsLetter(s As String) As Boolean
    Dim k As Long
    For k = 1 To Len(s)
        If UCase(Mid(s, k, 1)) Like "[A-Z]" Then
            ContainsLetter = True
            Exit Function
        End If
    Next k
    ContainsLetter = False
End Function

Private Function NZ(ByVal v As Variant) As Double
    If IsEmpty(v) Or v = "" Then
        NZ = 0
    Else
        NZ = CDbl(v)
    End If
End Function


