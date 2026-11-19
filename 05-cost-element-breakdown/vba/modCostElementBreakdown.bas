Attribute VB_Name = "modCostElementBreakdown"
Option Explicit

'==================================================================
' modCostElementBreakdown: actual costs (expanded view) per order
'
' Turns the order's "actual costs - expanded view" report (text export
' opened in Excel) into a grouped sheet: one block per cost element with
' its line items (materials, activities, settlements), a SUM subtotal per
' element, a grand total, and the report's own total underneath as a check.
' Any gap between an element's reported total and its line items is shown
' as a "(reconciling adjustment)" line instead of disappearing.
'
' Cost elements don't always mean the same thing from order to order:
' read the output as a starting point for analysis, not as the answer.
'
' MAIN ENTRY POINT - run BuildCostBreakdownComplete with the raw report sheet active
'==================================================================
Sub BuildCostBreakdownComplete()

    Dim srcWs As Worksheet
    Set srcWs = ActiveSheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    '---------------------------------------------------------
    ' 1. Work out which column the data actually starts in.
    '---------------------------------------------------------
    Dim baseCol As Long
    baseCol = 0
    Dim rr As Long, cc As Long
    For rr = 1 To 30
        For cc = 1 To 3
            If InStr(1, CStr(srcWs.Cells(rr, cc).value), "Report group", vbTextCompare) = 1 Then
                baseCol = cc
                Exit For
            End If
        Next cc
        If baseCol > 0 Then Exit For
    Next rr
    If baseCol = 0 Then
        MsgBox "Could not find 'Report group' in the first 30 rows/3 columns." & vbCrLf & _
               "Make sure the raw SAP text sheet is the active sheet.", vbExclamation
        GoTo CleanExit
    End If

    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.rows.Count, baseCol).End(xlUp).Row
    Dim lastRow2 As Long
    lastRow2 = srcWs.Cells(srcWs.rows.Count, baseCol + 1).End(xlUp).Row
    If lastRow2 > lastRow Then lastRow = lastRow2

    '---------------------------------------------------------
    ' 2. Parse into cost-element groups
    '---------------------------------------------------------
    Dim groups As Collection
    Set groups = New Collection
    Dim buffer As Collection
    Set buffer = New Collection

    Dim grandTotalVal As Variant, grandFixedVal As Variant, grandVarVal As Variant
    grandTotalVal = Empty: grandFixedVal = Empty: grandVarVal = Empty

    Dim r As Long
    For r = 1 To lastRow
        Dim field0 As String
        field0 = CStr(srcWs.Cells(r, baseCol).value)

        If Not IsBoilerplate(field0) Then

            Dim totalV As Variant, fixedV As Variant, varV As Variant
            totalV = GetCellNumber(srcWs.Cells(r, baseCol + 1))
            fixedV = GetCellNumber(srcWs.Cells(r, baseCol + 2))
            varV = GetCellNumber(srcWs.Cells(r, baseCol + 3))
            Dim qtyCellVal As Variant
            qtyCellVal = srcWs.Cells(r, baseCol + 4).value

            Dim qV As Variant, uV As String
            ParseQtyUnit CStr(qtyCellVal), qV, uV

            If Left(field0, 2) = "**" Then
                grandTotalVal = totalV
                grandFixedVal = fixedV
                grandVarVal = varV

            ElseIf Left(field0, 1) = "*" Then
                Dim code As String, Desc As String
                SplitIdDesc Trim(Mid(field0, 2)), code, Desc

                Dim grp As Object
                Set grp = CreateObject("Scripting.Dictionary")
                grp("code") = code
                grp("desc") = Desc
                grp("total") = totalV
                grp("fixed") = fixedV
                grp("variable") = varV
                Set grp("details") = buffer
                groups.Add grp
                Set buffer = New Collection

            Else
                Dim itemId As String, itemDesc As String
                SplitIdDesc field0, itemId, itemDesc

                If Not (itemId = "" And itemDesc = "" And IsEmpty(totalV) And _
                        IsEmpty(fixedV) And IsEmpty(varV) And IsEmpty(qV)) Then
                    Dim det As Object
                    Set det = CreateObject("Scripting.Dictionary")
                    If itemId = "" Then itemId = "(unassigned)"
                    det("item_id") = itemId
                    det("item_desc") = itemDesc
                    det("total") = totalV
                    det("fixed") = fixedV
                    det("variable") = varV
                    det("qty") = qV
                    det("unit") = uV
                    buffer.Add det
                End If
            End If
        End If
    Next r

    '---------------------------------------------------------
    ' 3. Fix any gap between a group's reported total and the
    '    sum of its own detail rows
    '---------------------------------------------------------
    Dim g As Object
    For Each g In groups
        ReconcileGap g
    Next g

    '---------------------------------------------------------
    ' 4. Build the output sheet
    '---------------------------------------------------------
    BuildGroupedDetailSheet groups, grandTotalVal, grandFixedVal, grandVarVal

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If baseCol > 0 Then MsgBox "Done - " & groups.Count & " cost elements processed.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'==================================================================
' PARSING HELPERS
'==================================================================

Function IsBoilerplate(ByVal ln As String) As Boolean
    Dim prefixes As Variant
    prefixes = Array("Report group", "Report ", "Selections", "Sales order", _
                      "Item from", "     to", "Controlling area", "Period from", _
                      "Values in", "Cost Elements", "Output Quantities")
    If Trim(ln) = "" Then
        IsBoilerplate = True
        Exit Function
    End If
    Dim p As Variant
    For Each p In prefixes
        If Left(ln, Len(p)) = p Then
            IsBoilerplate = True
            Exit Function
        End If
    Next p
    IsBoilerplate = False
End Function

Function GetCellNumber(ByVal rng As Range) As Variant
    Dim v As Variant
    v = rng.value
    If IsEmpty(v) Then
        GetCellNumber = Empty
    ElseIf VarType(v) = vbString Then
        If Trim(v) = "" Then
            GetCellNumber = Empty
        Else
            GetCellNumber = ParseNumber(CStr(v))
        End If
    Else
        GetCellNumber = CDbl(v)
    End If
End Function

Function ParseNumber(ByVal s As String) As Variant
    s = Trim(s)
    If s = "" Then
        ParseNumber = Empty
        Exit Function
    End If
    Dim neg As Boolean
    neg = False
    If Left(s, 1) = "-" Then
        neg = True
        s = Mid(s, 2)
    End If
    If Right(s, 1) = "-" Then
        neg = True
        s = Trim(Left(s, Len(s) - 1))
    End If
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")
    s = Replace(s, ",", ".")
    If s = "" Then
        ParseNumber = Empty
        Exit Function
    End If
    Dim v As Double
    v = val(s)
    If neg Then v = -v
    ParseNumber = v
End Function

Sub ParseQtyUnit(ByVal raw As String, ByRef qtyOut As Variant, ByRef unitOut As String)
    Dim s As String
    s = Trim(raw)
    qtyOut = Empty
    unitOut = ""
    If s = "" Then Exit Sub

    Dim toks() As String
    toks = Split(s, " ")

    unitOut = toks(UBound(toks))

    Dim numStr As String
    numStr = ""
    Dim i As Long
    For i = 0 To UBound(toks) - 1
        numStr = numStr & toks(i)
    Next i

    If UBound(toks) = 0 Then
        numStr = toks(0)
        unitOut = ""
    End If

    qtyOut = ParseNumber(numStr)
End Sub

Sub SplitIdDesc(ByVal raw As String, ByRef idOut As String, ByRef descOut As String)
    Dim s As String
    s = raw
    Do While Len(s) > 0 And Left(s, 1) = " "
        s = Mid(s, 2)
    Loop
    Dim pos As Long
    pos = InStr(s, "  ")
    If pos = 0 Then
        idOut = Trim(s)
        descOut = ""
    Else
        idOut = Trim(Left(s, pos - 1))
        Dim rest As String
        rest = Mid(s, pos)
        Do While Len(rest) > 0 And Left(rest, 1) = " "
            rest = Mid(rest, 2)
        Loop
        descOut = Trim(rest)
    End If
End Sub

Sub ReconcileGap(ByRef g As Object)
    If IsEmpty(g("total")) Then Exit Sub

    Dim detTotal As Double, detFixed As Double, detVar As Double
    Dim d As Object
    For Each d In g("details")
        If Not IsEmpty(d("total")) Then detTotal = detTotal + d("total")
        If Not IsEmpty(d("fixed")) Then detFixed = detFixed + d("fixed")
        If Not IsEmpty(d("variable")) Then detVar = detVar + d("variable")
    Next d

    Dim gTotal As Double, gFixed As Double, gVar As Double
    gTotal = g("total")
    gFixed = IIf(IsEmpty(g("fixed")), 0, g("fixed"))
    gVar = IIf(IsEmpty(g("variable")), 0, g("variable"))

    Dim gap As Double, gapFixed As Double, gapVar As Double
    gap = Round(gTotal - detTotal, 2)
    gapFixed = Round(gFixed - detFixed, 2)
    gapVar = Round(gVar - detVar, 2)

    If Abs(gap) >= 0.01 Then
        Dim adj As Object
        Set adj = CreateObject("Scripting.Dictionary")
        adj("item_id") = "(reconciling adjustment)"
        adj("item_desc") = "Balance vs. SAP report total for this cost element"
        adj("total") = gap
        adj("fixed") = IIf(gapFixed <> 0, gapFixed, Empty)
        adj("variable") = IIf(gapVar <> 0, gapVar, Empty)
        adj("qty") = Empty
        adj("unit") = ""
        If g("details").Count = 0 Then
            g("details").Add adj
        Else
            g("details").Add adj, , 1
        End If
    End If
End Sub

'==================================================================
' WORKSHEET BUILD
'==================================================================
Sub BuildGroupedDetailSheet(ByVal groups As Collection, ByVal grandTotalVal As Variant, _
                             ByVal grandFixedVal As Variant, ByVal grandVarVal As Variant)

    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Detail").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Detail"

    Const NUMFMT As String = "#,##0.00;(#,##0.00);""-"""
    Const QTYFMT As String = "#,##0;(#,##0);""-"""

    Dim headers As Variant
    headers = Array("Cost Element Code", "Cost Element Description", _
                     "Item / Material ID", "Item Description", "Total (ZAR)", _
                     "Fixed (ZAR)", "Variable (ZAR)", "Quantity", "Unit")
    Dim c As Long
    For c = 0 To UBound(headers)
        With ws.Cells(1, c + 1)
            .value = headers(c)
            .Font.Bold = True
            .Font.Name = "Arial"
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(48, 84, 150)
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next c
    ws.rows(1).RowHeight = 28

    Dim r As Long
    r = 2
    Dim subtotalRows As Collection
    Set subtotalRows = New Collection

    Dim g As Object, d As Object
    For Each g In groups
        If g("details").Count > 0 Then
            Dim blockStart As Long, blockEnd As Long
            blockStart = r

            For Each d In g("details")
                ws.Cells(r, 1).value = g("code")
                ws.Cells(r, 2).value = g("desc")
                ws.Cells(r, 3).value = d("item_id")
                ws.Cells(r, 4).value = d("item_desc")
                ws.Cells(r, 5).value = IIf(IsEmpty(d("total")), 0, d("total"))
                ws.Cells(r, 6).value = IIf(IsEmpty(d("fixed")), 0, d("fixed"))
                ws.Cells(r, 7).value = IIf(IsEmpty(d("variable")), 0, d("variable"))
                If Not IsEmpty(d("qty")) Then ws.Cells(r, 8).value = d("qty")
                ws.Cells(r, 9).value = d("unit")
                r = r + 1
            Next d
            blockEnd = r - 1

            ws.Cells(r, 1).Formula = "=A" & blockEnd
            ws.Cells(r, 2).Formula = "=B" & blockEnd
            ws.Cells(r, 5).Formula = "=SUM(E" & blockStart & ":E" & blockEnd & ")"
            ws.Cells(r, 6).Formula = "=SUM(F" & blockStart & ":F" & blockEnd & ")"
            ws.Cells(r, 7).Formula = "=SUM(G" & blockStart & ":G" & blockEnd & ")"
            ws.Cells(r, 8).Formula = "=SUM(H" & blockStart & ":H" & blockEnd & ")"

            Dim uUnit As String, mixed As Boolean, first As Boolean
            first = True: mixed = False: uUnit = ""
            For Each d In g("details")
                If d("unit") <> "" Then
                    If first Then
                        uUnit = d("unit")
                        first = False
                    ElseIf d("unit") <> uUnit Then
                        mixed = True
                    End If
                End If
            Next d
            If Not mixed And uUnit <> "" Then ws.Cells(r, 9).value = uUnit

            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 9))
                .Font.Bold = True
                .Interior.Color = RGB(217, 225, 242)
            End With

            ws.rows(blockStart & ":" & blockEnd).Group

            subtotalRows.Add r
            r = r + 1
        Else
            ws.Cells(r, 1).value = g("code")
            ws.Cells(r, 2).value = g("desc")
            ws.Cells(r, 5).value = IIf(IsEmpty(g("total")), 0, g("total"))
            ws.Cells(r, 6).value = IIf(IsEmpty(g("fixed")), 0, g("fixed"))
            ws.Cells(r, 7).value = IIf(IsEmpty(g("variable")), 0, g("variable"))

            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 9))
                .Font.Bold = True
                .Interior.Color = RGB(217, 225, 242)
            End With

            subtotalRows.Add r
            r = r + 1
        End If
    Next g

    Dim grandRow As Long
    grandRow = r
    ws.Cells(grandRow, 4).value = "TOTAL - All Cost Elements"
    Dim colLetters As Variant
    colLetters = Array("E", "F", "G")
    Dim ci As Long
    For ci = 0 To 2
        Dim refs As String
        refs = ""
        Dim sr As Variant
        For Each sr In subtotalRows
            If refs <> "" Then refs = refs & ","
            refs = refs & colLetters(ci) & sr
        Next sr
        ws.Cells(grandRow, 5 + ci).Formula = "=SUM(" & refs & ")"
    Next ci
    With ws.Range(ws.Cells(grandRow, 1), ws.Cells(grandRow, 9))
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With

    ws.Range("E2:G" & grandRow).NumberFormat = NUMFMT
    ws.Range("H2:H" & grandRow).NumberFormat = QTYFMT

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 8
    ws.Range(ws.Cells(1, 1), ws.Cells(grandRow, 9)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(1, 1), ws.Cells(grandRow, 9)).Borders.Color = RGB(183, 183, 183)

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1:I" & grandRow).AutoFilter

    If Not IsEmpty(grandTotalVal) Then
        Dim noteRow As Long
        noteRow = grandRow + 2
        ws.Cells(noteRow, 1).value = "Source report grand total: R " & _
            Format(grandTotalVal, "#,##0.00") & "  (Fixed R " & _
            Format(grandFixedVal, "#,##0.00") & " / Variable R " & _
            Format(grandVarVal, "#,##0.00") & ")"
        ws.Cells(noteRow, 1).Font.Italic = True
        ws.Cells(noteRow, 1).Font.Size = 9
        ws.Cells(noteRow, 1).Font.Color = RGB(128, 128, 128)
    End If
End Sub


