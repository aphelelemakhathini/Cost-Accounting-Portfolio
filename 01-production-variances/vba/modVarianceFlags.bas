Attribute VB_Name = "modVarianceFlags"
'===============================================================================
' modVarianceFlags: labour AND material variance flagging (Excel VBA)
'
' Stage 1 of the variance process. Takes the production order cost report
' (actual vs standard, labour and material, surcharges excluded) and keeps the
' orders outside the thresholds:
'   Labour   (over +45% / under -45%) -> "Labour Raw Data", "Labour Pivot Summary",
'                                        "Labour Comments" (pre-filled op/comment layout)
'   Material (over +20% / under -20%) -> "Material Raw Data"
' The order numbers on Labour Comments then go into the operations report for
' stage 2 (modHoursVarianceReview).
'
' PLANTS: any 4-digit plant code in column B is picked up automatically.
' To restrict to specific plants, list them in PLANT_FILTER, e.g. "1100,1200".
'
' SOURCE LAYOUT (3-row block per order, starting at the plant row):
'   Row 1: B = Plant, E = Production Order, G = Item/Material
'   Row 2: F = Actual Labour, K = Standard Labour, I = Actual Material, L = Standard Material
'   Row 3: B = Quantity, O = Labour % variance, P = Material variance (R), Q = Material % variance
' Numbers arrive as text in the ERP's format ("1.234,56"); CleanValue converts them
' whatever the PC's regional settings are.
'===============================================================================

Option Explicit

Private Const PLANT_FILTER As String = ""
Private Const LABOUR_LIMIT As Double = 45
Private Const MATERIAL_LIMIT As Double = 20

'Formats
Private Const QTY_FMT As String = "#,##0"
Private Const RAND_FMT As String = "_-""R""* #,##0.00_-;-""R""* #,##0.00_-;_-""R""* ""-""??_-;_-@_-"   'Accounting (R)
'Known order types, used only to LOCATE the order type column in the export.
'Once found, whatever value sits in that column is picked up (PP01, PP02 or any other type).
Private Const ORDER_TYPES As String = "PP01,PP02"

Private otCol As Long      'column of the order type (0 = not found)
Private otOffset As Long   'row within the 3-row block (0, 1 or 2)


'=====================================
' HELPERS
'=====================================

'True if the value looks like a plant code (4 digits, dots stripped) and passes PLANT_FILTER
Private Function IsPlantRow(ByVal v As Variant) As Boolean
    Dim s As String

    If IsError(v) Or IsEmpty(v) Then Exit Function

    s = Trim$(Replace(CStr(v), ".", ""))
    If Not s Like "####" Then Exit Function

    If Len(PLANT_FILTER) = 0 Then
        IsPlantRow = True
    Else
        IsPlantRow = InStr("," & Replace(PLANT_FILTER, " ", "") & ",", "," & s & ",") > 0
    End If
End Function

'Converts ERP text ("1.234,56": dot thousands, comma decimal) to a number,
'whatever the PC's regional settings; non-numbers come back as trimmed text
Private Function CleanValue(ByVal v As Variant) As Variant
    Dim s As String
    If IsError(v) Then
        CleanValue = ""
        Exit Function
    End If
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        CleanValue = CDbl(v)                          ' already a real number
        Exit Function
    End If
    s = Trim$(CStr(v))
    s = Replace(s, ".", "")                           ' thousands separator
    If Right$(s, 1) = "-" Then s = "-" & Left$(s, Len(s) - 1)   ' trailing minus
    s = Replace(s, ",", Application.International(xlDecimalSeparator))
    If Len(s) > 0 And IsNumeric(s) Then
        CleanValue = CDbl(s)
    Else
        CleanValue = Trim$(CStr(v))
    End If
End Function

'True if the value is numeric and above +limit or below -limit
Private Function ExceedsLimit(ByVal v As Variant, ByVal limit As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    v = CleanValue(v)
    If VarType(v) <> vbDouble Then Exit Function
    ExceedsLimit = (v > limit Or v < -limit)
End Function

'Finds which column/row of the order block holds the order type, by looking for
'the first known order type in the data. Every order's type is then read from that same spot.
Private Sub DetectOrderTypeLocation(ws As Worksheet, ByVal lastRow As Long)
    Dim i As Long, r As Long, c As Long, v As Variant, txt As String
    Dim known As String
    known = "," & Replace(UCase$(ORDER_TYPES), " ", "") & ","
    otCol = 0
    otOffset = 0

    i = 1
    Do While i <= lastRow
        If IsPlantRow(ws.Cells(i, "B").value) And i + 2 <= lastRow Then
            For r = 0 To 2
                For c = 1 To 25   'columns A:Y
                    v = ws.Cells(i + r, c).value
                    If Not IsError(v) Then
                        txt = UCase$(Trim$(CStr(v)))
                        If Len(txt) > 0 And InStr(known, "," & txt & ",") > 0 Then
                            otCol = c
                            otOffset = r
                            Exit Sub
                        End If
                    End If
                Next c
            Next r
            i = i + 3
        Else
            i = i + 1
        End If
    Loop
End Sub

'Returns whatever order type is in the detected column for this order (any type, not just the known ones)
Private Function FindOrderType(ws As Worksheet, ByVal groupStart As Long) As String
    Dim v As Variant
    If otCol = 0 Then Exit Function
    v = ws.Cells(groupStart + otOffset, otCol).value
    If Not IsError(v) Then FindOrderType = UCase$(Trim$(CStr(v)))
End Function

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub NoSubtotals(pf As PivotField)
    pf.Subtotals = Array(False, False, False, False, False, False, _
                         False, False, False, False, False, False)
End Sub


'=====================================
' MAIN
'=====================================
Sub GroupFilterAndExport()

    '=====================================
    ' PART 1: Group and flag on the source sheet
    '   Z  = Labour qualifies
    '   AA = Material qualifies
    '   AB = Either qualifies (used for the filter)
    '=====================================
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim labFlag As Boolean, matFlag As Boolean

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).Row

    ws.Cells.EntireColumn.AutoFit

    On Error Resume Next
    ws.rows.ClearOutline
    ws.AutoFilterMode = False
    On Error GoTo 0

    ws.Columns("Z:AB").ClearContents
    ws.Cells(1, "Z").value = "Labour Keep"
    ws.Cells(1, "AA").value = "Material Keep"
    ws.Cells(1, "AB").value = "Keep"

    i = 1
    Do While i <= lastRow
        If IsPlantRow(ws.Cells(i, "B").value) And i + 2 <= lastRow Then
            ws.rows(i & ":" & i + 2).Group

            labFlag = ExceedsLimit(ws.Cells(i + 2, "O").value, LABOUR_LIMIT)
            matFlag = ExceedsLimit(ws.Cells(i + 2, "Q").value, MATERIAL_LIMIT)

            ws.Range("Z" & i & ":Z" & i + 2).value = IIf(labFlag, "Yes", "No")
            ws.Range("AA" & i & ":AA" & i + 2).value = IIf(matFlag, "Yes", "No")
            ws.Range("AB" & i & ":AB" & i + 2).value = IIf(labFlag Or matFlag, "Yes", "No")

            i = i + 3
        Else
            i = i + 1
        End If
    Loop

    ws.Range("A1:AB" & lastRow).AutoFilter Field:=28, Criteria1:="Yes"

    '=====================================
    ' PART 2: Build Labour and Material outputs
    '=====================================
    Dim plantsFound As Object
    Set plantsFound = CreateObject("Scripting.Dictionary")

    DetectOrderTypeLocation ws, lastRow

    Dim anchor As Worksheet
    Set anchor = ws

    Dim labourCount As Long, materialCount As Long
    labourCount = BuildLabour(ws, lastRow, plantsFound, anchor)
    materialCount = BuildMaterial(ws, lastRow, plantsFound, anchor)

    Dim plantList As String
    If plantsFound.Count > 0 Then
        plantList = Join(plantsFound.keys, ", ")
    Else
        plantList = "(none)"
    End If

    MsgBox "Done!" & vbCrLf & _
           "Labour orders exported: " & labourCount & vbCrLf & _
           "Material orders exported: " & materialCount & vbCrLf & _
           "Plants processed: " & plantList & _
           IIf(otCol = 0, vbCrLf & vbCrLf & "WARNING: Order type column not found (no PP01/PP02 in the data) - Order Type left blank.", ""), _
           IIf(otCol = 0, vbExclamation, vbInformation)

End Sub


'=====================================
' LABOUR
'=====================================
Private Function BuildLabour(ws As Worksheet, ByVal lastRow As Long, _
                             plantsFound As Object, ByRef anchor As Worksheet) As Long

    Dim wsRawData As Worksheet
    Dim exportRow As Long
    Dim groupStart As Long
    Dim plantCode As String
    Dim i As Long

    DeleteSheetIfExists "Labour Raw Data"
    DeleteSheetIfExists "Labour Pivot Summary"
    DeleteSheetIfExists "Labour Comments"

    Set wsRawData = ThisWorkbook.Sheets.Add(After:=anchor)
    wsRawData.Name = "Labour Raw Data"
    Set anchor = wsRawData

    wsRawData.Range("A1:K1").value = Array("Plant", "Order Type", "Production Order", "Item", "Quantity", _
        "Actual Labour", "Total Actual Labour", "Standard Labour", "Total Standard Labour", _
        "Variance", "% of Variance")

    exportRow = 2

    i = 1
    Do While i <= lastRow
        If IsPlantRow(ws.Cells(i, "B").value) And i + 2 <= lastRow Then

            If ws.Cells(i, "Z").value = "Yes" Then
                groupStart = i

                plantCode = Replace(CStr(ws.Cells(groupStart, "B").value), ".", "")
                wsRawData.Cells(exportRow, 1).NumberFormat = "@"
                wsRawData.Cells(exportRow, 1).value = plantCode
                wsRawData.Cells(exportRow, 1).Errors.item(xlNumberAsText).Ignore = True
                If Not plantsFound.Exists(plantCode) Then plantsFound.Add plantCode, True

                wsRawData.Cells(exportRow, 2).value = FindOrderType(ws, groupStart)                    ' Order Type
                wsRawData.Cells(exportRow, 3).value = CleanValue(ws.Cells(groupStart, "E").value)      ' Production Order
                wsRawData.Cells(exportRow, 4).value = CleanValue(ws.Cells(groupStart, "G").value)      ' Item
                wsRawData.Cells(exportRow, 5).value = CleanValue(ws.Cells(groupStart + 2, "B").value)  ' Quantity
                wsRawData.Cells(exportRow, 6).value = CleanValue(ws.Cells(groupStart + 1, "F").value)  ' Actual Labour
                wsRawData.Cells(exportRow, 8).value = CleanValue(ws.Cells(groupStart + 1, "K").value)  ' Standard Labour

                wsRawData.Cells(exportRow, 7).Formula = "=IF(AND(E" & exportRow & "<>"""",F" & exportRow & "<>""""),E" & exportRow & "*F" & exportRow & ","""")"
                wsRawData.Cells(exportRow, 9).Formula = "=IF(AND(E" & exportRow & "<>"""",H" & exportRow & "<>""""),E" & exportRow & "*H" & exportRow & ","""")"
                wsRawData.Cells(exportRow, 10).Formula = "=IF(AND(I" & exportRow & "<>"""",G" & exportRow & "<>""""),I" & exportRow & "-G" & exportRow & ","""")"
                wsRawData.Cells(exportRow, 11).Formula = "=IF(AND(J" & exportRow & "<>"""",G" & exportRow & "<>"""",G" & exportRow & "<>0),J" & exportRow & "/G" & exportRow & ","""")"

                exportRow = exportRow + 1
            End If

            i = i + 3
        Else
            i = i + 1
        End If
    Loop

    BuildLabour = exportRow - 2
    If BuildLabour = 0 Then
        wsRawData.Range("A2").value = "No labour variances above " & LABOUR_LIMIT & " / below -" & LABOUR_LIMIT
        Exit Function
    End If

    ' --- Table ---
    Dim dataLastRow As Long
    dataLastRow = exportRow - 1

    Dim lo As ListObject
    Set lo = wsRawData.ListObjects.Add(xlSrcRange, wsRawData.Range("A1:K" & dataLastRow), , xlYes)
    lo.Name = "DataTable"
    lo.TableStyle = "TableStyleLight9"   'Dark Teal, Table Style Light 9

    wsRawData.Range("D1:D" & dataLastRow).HorizontalAlignment = xlRight   ' Item/Material
    wsRawData.Range("E2:E" & dataLastRow).NumberFormat = QTY_FMT     ' Quantity
    wsRawData.Range("F2:J" & dataLastRow).NumberFormat = RAND_FMT    ' Labour costs + Variance
    wsRawData.Range("K2:K" & dataLastRow).NumberFormat = "0.00%"

    ' Totals row: sum the rand totals and variance, % = total variance / total actual
    lo.ShowTotals = True
    lo.ListColumns(1).TotalsCalculation = xlTotalsCalculationNone
    lo.TotalsRowRange.Cells(1, 1).value = "Total"
    lo.ListColumns("% of Variance").TotalsCalculation = xlTotalsCalculationNone
    lo.ListColumns("Total Actual Labour").TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns("Total Standard Labour").TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns("Variance").TotalsCalculation = xlTotalsCalculationSum
    With lo.ListColumns("% of Variance").total
        .Formula = "=IFERROR(SUBTOTAL(109,DataTable[Variance])/SUBTOTAL(109,DataTable[Total Actual Labour]),"""")"
        .NumberFormat = "0.00%"
    End With
    lo.TotalsRowRange.Columns("F:J").NumberFormat = RAND_FMT
    wsRawData.Cells.EntireColumn.AutoFit

    ' --- Pivot ---
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable

    Set wsPivot = ThisWorkbook.Sheets.Add(After:=anchor)
    wsPivot.Name = "Labour Pivot Summary"
    Set anchor = wsPivot

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=wsRawData.Range("A1:K" & dataLastRow))
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("B2"), TableName:="SummaryPivot")

    pt.TableStyle2 = "PivotStyleMedium2"
    pt.RowAxisLayout xlTabularRow
    pt.RowGrand = False
    pt.ColumnGrand = False

    With pt.PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
        .EnableMultiplePageItems = True
        .ClearAllFilters
    End With

    With pt.PivotFields("Order Type")
        .Orientation = xlPageField
        .Position = 2
        .EnableMultiplePageItems = True
        .ClearAllFilters
    End With

    With pt.PivotFields("Production Order")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("Item")
        .Orientation = xlRowField
        .Position = 2
    End With
    NoSubtotals pt.PivotFields("Production Order")
    NoSubtotals pt.PivotFields("Item")

    Dim valueFields As Variant, v As Long
    valueFields = Array("Quantity", "Actual Labour", "Total Actual Labour", _
                        "Standard Labour", "Total Standard Labour", "Variance")
    For v = LBound(valueFields) To UBound(valueFields)
        With pt.PivotFields(valueFields(v))
            .Orientation = xlDataField
            .Function = xlMax
            If valueFields(v) = "Quantity" Then
                .NumberFormat = QTY_FMT
            Else
                .NumberFormat = RAND_FMT
            End If
            .Name = valueFields(v) & " "
        End With
    Next v

    With pt.PivotFields("% of Variance")
        .Orientation = xlDataField
        .Function = xlMax
        .NumberFormat = "0.00%"
        .Name = "% of Variance "
    End With

    pt.PivotFields("Production Order").RepeatLabels = True
    pt.PivotFields("Item").RepeatLabels = True
    wsPivot.Cells.EntireColumn.AutoFit

    ' --- Comments ---
    Dim wsComments As Worksheet
    Dim r As Long

    Set wsComments = ThisWorkbook.Sheets.Add(After:=anchor)
    wsComments.Name = "Labour Comments"
    Set anchor = wsComments

    With wsComments.Range("D1:F1")
        .Merge
        .value = "Comments"
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    wsComments.Range("A2:H2").value = Array("Plant", "Production Order", "Item", "% of Variance", _
        "Operation", "Comments on Machine", "Comments on Setup", "Remarks")

    With wsComments.Range("A2:H2")
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim operationText As String, commentText As String
    operationText = "OP20(MCH19)" & Chr(10) & "OP30(MCH19)" & Chr(10) & "OP40(MCH19)"
    commentText = "Planned hrs ; Confirmed hrs" & Chr(10) & _
                  "Planned hrs ; Confirmed hrs" & Chr(10) & _
                  "Planned hrs ; Confirmed hrs"

    For r = 2 To dataLastRow
        wsComments.Cells(r + 1, 1).NumberFormat = "@"
        wsComments.Cells(r + 1, 1).value = wsRawData.Cells(r, 1).value
        wsComments.Cells(r + 1, 2).value = wsRawData.Cells(r, 3).value   ' Production Order
        wsComments.Cells(r + 1, 3).value = wsRawData.Cells(r, 4).value   ' Item
        wsComments.Cells(r + 1, 4).value = wsRawData.Cells(r, 11).value  ' % of Variance
        wsComments.Cells(r + 1, 4).NumberFormat = "0.00%"

        If wsComments.Cells(r + 1, 1).value <> "" Then
            wsComments.Cells(r + 1, 5).value = operationText
            wsComments.Cells(r + 1, 6).value = commentText
            wsComments.Cells(r + 1, 7).value = commentText
        End If
    Next r

    With wsComments.Range("D3:G" & dataLastRow + 1)
        .Style = "Explanatory Text"
        .WrapText = True
        .VerticalAlignment = xlTop
    End With

    wsComments.Columns("A:H").AutoFit
    wsComments.rows.AutoFit

    ' Hide the pivot sheet (unhide via right-click on any sheet tab > Unhide)
    wsPivot.Visible = xlSheetHidden

End Function


'=====================================
' MATERIAL
'=====================================
Private Function BuildMaterial(ws As Worksheet, ByVal lastRow As Long, _
                               plantsFound As Object, ByRef anchor As Worksheet) As Long

    Dim wsMat As Worksheet
    Dim exportRow As Long
    Dim groupStart As Long
    Dim plantCode As String
    Dim i As Long

    DeleteSheetIfExists "Material Raw Data"

    Set wsMat = ThisWorkbook.Sheets.Add(After:=anchor)
    wsMat.Name = "Material Raw Data"
    Set anchor = wsMat

    wsMat.Range("A1:K1").value = Array("Plant", "Order Type", "Production Order", "Material", "Quantity", _
        "Standard Mat. Costs", "Total Standard Mat. Costs", "Actual Mat. Costs", "Total Actual Mat. Costs", _
        "Variance", "% of Variance")

    exportRow = 2

    i = 1
    Do While i <= lastRow
        If IsPlantRow(ws.Cells(i, "B").value) And i + 2 <= lastRow Then

            If ws.Cells(i, "AA").value = "Yes" Then
                groupStart = i

                plantCode = Replace(CStr(ws.Cells(groupStart, "B").value), ".", "")
                wsMat.Cells(exportRow, 1).NumberFormat = "@"
                wsMat.Cells(exportRow, 1).value = plantCode
                wsMat.Cells(exportRow, 1).Errors.item(xlNumberAsText).Ignore = True
                If Not plantsFound.Exists(plantCode) Then plantsFound.Add plantCode, True

                wsMat.Cells(exportRow, 2).value = FindOrderType(ws, groupStart)                    ' Order Type
                wsMat.Cells(exportRow, 3).value = CleanValue(ws.Cells(groupStart, "E").value)      ' Production Order
                wsMat.Cells(exportRow, 4).value = CleanValue(ws.Cells(groupStart, "G").value)      ' Material
                wsMat.Cells(exportRow, 5).value = CleanValue(ws.Cells(groupStart + 2, "B").value)  ' Quantity
                wsMat.Cells(exportRow, 6).value = CleanValue(ws.Cells(groupStart + 1, "L").value)  ' Standard Mat. Costs (per unit)
                wsMat.Cells(exportRow, 8).value = CleanValue(ws.Cells(groupStart + 1, "I").value)  ' Actual Mat. Costs (per unit)

                ' G = Total Standard = Qty x Standard
                wsMat.Cells(exportRow, 7).Formula = "=IF(AND(E" & exportRow & "<>"""",F" & exportRow & "<>""""),E" & exportRow & "*F" & exportRow & ","""")"
                ' I = Total Actual = Qty x Actual
                wsMat.Cells(exportRow, 9).Formula = "=IF(AND(E" & exportRow & "<>"""",H" & exportRow & "<>""""),E" & exportRow & "*H" & exportRow & ","""")"
                ' J = Variance = Total Standard - Total Actual (same as labour)
                wsMat.Cells(exportRow, 10).Formula = "=IF(AND(G" & exportRow & "<>"""",I" & exportRow & "<>""""),G" & exportRow & "-I" & exportRow & ","""")"
                ' K = % of Variance = Variance / Total Actual (same as labour)
                wsMat.Cells(exportRow, 11).Formula = "=IF(AND(J" & exportRow & "<>"""",I" & exportRow & "<>"""",I" & exportRow & "<>0),J" & exportRow & "/I" & exportRow & ","""")"

                exportRow = exportRow + 1
            End If

            i = i + 3
        Else
            i = i + 1
        End If
    Loop

    BuildMaterial = exportRow - 2
    If BuildMaterial = 0 Then
        wsMat.Range("A2").value = "No material variances above " & MATERIAL_LIMIT & " / below -" & MATERIAL_LIMIT
        Exit Function
    End If

    ' --- Table ---
    Dim dataLastRow As Long
    dataLastRow = exportRow - 1

    Dim lo As ListObject
    Set lo = wsMat.ListObjects.Add(xlSrcRange, wsMat.Range("A1:K" & dataLastRow), , xlYes)
    lo.Name = "MaterialTable"
    lo.TableStyle = "TableStyleLight9"   'Dark Teal, Table Style Light 9

    wsMat.Range("D1:D" & dataLastRow).HorizontalAlignment = xlRight       ' Material
    wsMat.Range("E2:E" & dataLastRow).NumberFormat = QTY_FMT         ' Quantity
    wsMat.Range("F2:J" & dataLastRow).NumberFormat = RAND_FMT        ' Material costs + Variance
    wsMat.Range("K2:K" & dataLastRow).NumberFormat = "0.00%"

    ' Totals row: sum the rand totals and variance, % = total variance / total actual
    lo.ShowTotals = True
    lo.ListColumns(1).TotalsCalculation = xlTotalsCalculationNone
    lo.TotalsRowRange.Cells(1, 1).value = "Total"
    lo.ListColumns("% of Variance").TotalsCalculation = xlTotalsCalculationNone
    lo.ListColumns("Total Standard Mat. Costs").TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns("Total Actual Mat. Costs").TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns("Variance").TotalsCalculation = xlTotalsCalculationSum
    With lo.ListColumns("% of Variance").total
        .Formula = "=IFERROR(SUBTOTAL(109,MaterialTable[Variance])/SUBTOTAL(109,MaterialTable[Total Actual Mat. Costs]),"""")"
        .NumberFormat = "0.00%"
    End With
    lo.TotalsRowRange.Columns("F:J").NumberFormat = RAND_FMT
    wsMat.Cells.EntireColumn.AutoFit

End Function


