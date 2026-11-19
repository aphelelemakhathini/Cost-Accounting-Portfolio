Attribute VB_Name = "modCostedBomBuilder"
Option Explicit

'==============================================================================
' modCostedBomBuilder
'
' Turns a CK11N costed BOM (cost estimate with quantity structure) text export
' into a clean, grouped, formula-driven costing sheet. Ported from my Python
' version of the same tool (sap_bom_to_excel.py) to native Excel VBA.
'
' USE: open the raw SAP "Create/Display Material Cost Estimate with Quantity
' Structure" text export directly in Excel (File > Open on the .txt - Excel
' splits it into columns by tab automatically), make sure that sheet is the
' active one, then run BuildFormattedBom. It reads the raw cells and builds a
' clean, grouped, formula-driven BOM on a NEW sheet:
'   - Parent/sub-total rows  -> "=SUM(immediate children)"
'   - Leaf rows (material or operation lines) -> Total value = "=Price*Qty"
'
' WHY THIS WORKS ON ANY BOM DEPTH:
' SAP always pads the description field with tabs, then always appends the
' SAME fixed set of fields at the END of every line, in this order:
'   Material, Price, Quantity, Total value, Manuf.material, Traded material,
'   Manuf.Costs total, Mat. overhead, Manu overhead, Variable value,
'   Fixed Value, [Work Ctr, ActTyp]
' Because these open into fixed columns from the end of each row, the macro
' doesn't care how many hierarchy levels a given BOM has.
'
' HOW TO USE
'   1. Alt+F11 to open the VBA editor, File > Import File..., pick this .bas
'   2. Open the SAP .txt export in Excel and make sure it's the active sheet
'   3. Alt+F8, run "BuildFormattedBom" - a new formatted sheet is added
'==============================================================================

Private Type BomRow
    depth As Integer
    Desc As String
    material As String
    Price As Variant
    Qty As Variant
    total As Variant
    ManufMat As Variant
    TradedMat As Variant
    ManufCostTot As Variant
    MatOH As Variant
    ManuOH As Variant
    Variable As Variant
    Fixed As Variant
    workctr As String
    acttyp As String
    Children() As Long
    ChildCount As Long
End Type

Private Const CURRENCY_FMT As String = "_-""R""* #,##0.00_-;-""R""* #,##0.00_-;_-""R""* ""-""??_-;_-@_-"
Private Const QTY_FMT As String = "#,##0.000;-#,##0.000;""-"""

'------------------------------------------------------------------------------
' MAIN ENTRY POINT - run this. Reads the ACTIVE SHEET (the SAP .txt you
' already opened in Excel) and builds a formatted BOM on a new sheet.
'------------------------------------------------------------------------------
Public Sub BuildFormattedBom()

    Dim srcWs As Worksheet
    Set srcWs = ActiveSheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim metaDict As Object
    Set metaDict = CreateObject("Scripting.Dictionary")

    Dim rows() As BomRow
    rows = ParseSapBomFromSheet(srcWs, metaDict)

    Dim rowTotal As Long
    On Error Resume Next
    rowTotal = UBound(rows) - LBound(rows) + 1
    On Error GoTo 0

    If rowTotal <= 0 Then
        MsgBox "No data rows were found on the active sheet - make sure the" & _
               " raw SAP export is the sheet you're looking at when you run this.", vbExclamation
        GoTo CleanExit
    End If

    AttachChildren rows

    Dim warnCount As Long
    warnCount = ValidateRollups(rows)

    Dim outWs As Worksheet
    Dim sheetName As String
    sheetName = "BOM"
    If metaDict.Exists("Material") Then
        Dim matField As String
        matField = Trim(metaDict("Material"))
        If Len(matField) > 0 Then
            sheetName = CleanSheetName(CStr(Split(matField, " ")(0)))   ' material number only
        End If
    End If
    Set outWs = AddUniqueSheet(sheetName, srcWs.Parent)

    BuildWorkbook rows, metaDict, outWs

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    outWs.Activate
    outWs.Range("A1").Select

    If warnCount > 0 Then
        MsgBox "Done - built '" & outWs.Name & "'. " & warnCount & " rollup mismatch(es) logged to the" & _
               " Immediate window (Ctrl+G) - usually SAP's own setup-time rounding, safe to ignore if small.", vbInformation
    Else
        MsgBox "Done - built '" & outWs.Name & "'. Every rollup matched SAP's reported totals exactly.", vbInformation
    End If
    Exit Sub

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------
' 1. PARSE the SAP export directly from worksheet cells (already opened/split
'    into columns by Excel) into an array of BomRow
'------------------------------------------------------------------------------
Private Function ParseSapBomFromSheet(srcWs As Worksheet, ByRef metaDict As Object) As BomRow()

    Dim knownMetaKeys As Variant
    knownMetaKeys = Array("Material", "Plant", "Costing Variant", "Costing Version", _
                           "Costing Date from-to", "Lot Size", "Cost Base", _
                           "Cost Comp. Structure", "Created By", "Entered On")

    Dim ur As Range
    Set ur = srcWs.UsedRange
    Dim firstRow As Long, lastRow As Long
    firstRow = ur.Row
    lastRow = ur.Row + ur.rows.Count - 1

    Dim headerSeen As Boolean
    headerSeen = False

    Dim rows() As BomRow
    Dim rowCount As Long
    rowCount = 0
    ReDim rows(0 To 500)   ' generous starting size, trimmed at the end

    Dim rr As Long
    For rr = firstRow To lastRow

        Dim parts() As Variant
        parts = RowParts(srcWs, rr)

        Dim totalParts As Long
        On Error Resume Next
        totalParts = UBound(parts) - LBound(parts) + 1
        On Error GoTo 0
        If totalParts <= 0 Then GoTo ContinueLoop   ' fully blank row

        If Not AnyNonEmpty(parts) Then GoTo ContinueLoop

        If Not headerSeen Then
            If RowContains(parts, "Costing Structure") Then
                headerSeen = True
            Else
                Dim metaParts() As String
                metaParts = FilterNonEmpty(parts)
                Dim mCount As Long
                On Error Resume Next
                mCount = UBound(metaParts) - LBound(metaParts) + 1
                On Error GoTo 0
                If mCount >= 2 Then
                    Dim k As Long
                    For k = LBound(knownMetaKeys) To UBound(knownMetaKeys)
                        If metaParts(0) = knownMetaKeys(k) Then
                            Dim v As String, j As Long
                            v = ""
                            For j = 1 To UBound(metaParts)
                                v = v & IIf(j > 1, " ", "") & metaParts(j)
                            Next j
                            metaDict(CStr(knownMetaKeys(k))) = Trim(v)
                            Exit For
                        End If
                    Next k
                End If
            End If
            GoTo ContinueLoop
        End If

        If totalParts < 11 Then GoTo ContinueLoop

        Dim hasCodes As Boolean
        hasCodes = False
        If totalParts >= 13 Then
            Dim lastF As String, secondLastF As String
            lastF = Trim(CStr(parts(UBound(parts))))
            secondLastF = Trim(CStr(parts(UBound(parts) - 1)))
            If ContainsLetter(lastF) Or ContainsLetter(secondLastF) Then
                hasCodes = True
            End If
        End If

        Dim nTrailing As Long
        nTrailing = IIf(hasCodes, 13, 11)

        Dim nDesc As Long
        nDesc = totalParts - nTrailing
        If nDesc < 1 Then GoTo ContinueLoop

        ' find depth = index (0-based) of first non-empty description field
        Dim depth As Long, descText As String
        depth = -1
        Dim d As Long
        For d = 0 To nDesc - 1
            If Not IsBlankVal(parts(LBound(parts) + d)) Then
                depth = d
                descText = Trim(CStr(parts(LBound(parts) + d)))
                Exit For
            End If
        Next d
        If depth = -1 Then GoTo ContinueLoop

        Dim baseIdx As Long
        baseIdx = LBound(parts) + nDesc   ' start of the trailing numeric block

        Dim material As String, workctr As String, acttyp As String

        material = Trim(CStr(parts(baseIdx + 0)))

        If hasCodes Then
            workctr = Trim(CStr(parts(baseIdx + 11)))
            acttyp = Trim(CStr(parts(baseIdx + 12)))
        Else
            workctr = ""
            acttyp = ""
        End If

        If rowCount > UBound(rows) Then ReDim Preserve rows(0 To UBound(rows) + 500)

        With rows(rowCount)
            .depth = CInt(depth)
            .Desc = descText
            .material = material
            .Price = ToNumber(parts(baseIdx + 1))
            .Qty = ToNumber(parts(baseIdx + 2))
            .total = ToNumber(parts(baseIdx + 3))
            .ManufMat = ToNumber(parts(baseIdx + 4))
            .TradedMat = ToNumber(parts(baseIdx + 5))
            .ManufCostTot = ToNumber(parts(baseIdx + 6))
            .MatOH = ToNumber(parts(baseIdx + 7))
            .ManuOH = ToNumber(parts(baseIdx + 8))
            .Variable = ToNumber(parts(baseIdx + 9))
            .Fixed = ToNumber(parts(baseIdx + 10))
            .workctr = workctr
            .acttyp = acttyp
            .ChildCount = 0
        End With

        rowCount = rowCount + 1
ContinueLoop:
    Next rr

    If rowCount = 0 Then
        Dim emptyRows() As BomRow
        ReDim emptyRows(0 To -1)   ' zero-length array so the caller's UBound check catches it
        ParseSapBomFromSheet = emptyRows
        Exit Function
    End If

    ReDim Preserve rows(0 To rowCount - 1)

    ' Normalise depth so the shallowest row becomes depth 0
    Dim minDepth As Integer
    minDepth = rows(0).depth
    Dim i As Long
    For i = 1 To rowCount - 1
        If rows(i).depth < minDepth Then minDepth = rows(i).depth
    Next i
    If minDepth <> 0 Then
        For i = 0 To rowCount - 1
            rows(i).depth = rows(i).depth - minDepth
        Next i
    End If

    ParseSapBomFromSheet = rows
End Function

'------------------------------------------------------------------------------
' 2. Build immediate-child index lists for every row
'------------------------------------------------------------------------------
Private Sub AttachChildren(ByRef rows() As BomRow)
    Dim n As Long
    n = UBound(rows) - LBound(rows) + 1

    Dim i As Long, j As Long
    For i = 0 To n - 1
        Dim childList() As Long
        Dim cCount As Long
        cCount = 0
        ReDim childList(0 To n)   ' generous, trimmed after

        For j = i + 1 To n - 1
            If rows(j).depth <= rows(i).depth Then Exit For
            If rows(j).depth = rows(i).depth + 1 Then
                childList(cCount) = j
                cCount = cCount + 1
            End If
        Next j

        If cCount > 0 Then
            ReDim rows(i).Children(0 To cCount - 1)
            Dim k As Long
            For k = 0 To cCount - 1
                rows(i).Children(k) = childList(k)
            Next k
        End If
        rows(i).ChildCount = cCount
    Next i
End Sub

'------------------------------------------------------------------------------
' 3. Write everything to the worksheet
'------------------------------------------------------------------------------
Private Sub BuildWorkbook(ByRef rows() As BomRow, metaDict As Object, ws As Worksheet)

    Dim n As Long
    n = UBound(rows) - LBound(rows) + 1

    Dim maxDepth As Integer
    maxDepth = 0
    Dim i As Long
    For i = 0 To n - 1
        If rows(i).depth > maxDepth Then maxDepth = rows(i).depth
    Next i
    Dim nDescCols As Long
    nDescCols = maxDepth + 1

    Dim colMaterial As Long, colPrice As Long, colQty As Long, colTotal As Long
    Dim colManufMat As Long, colTradedMat As Long, colManufCostTot As Long
    Dim colMatOH As Long, colManuOH As Long, colVariable As Long, colFixed As Long
    Dim colWorkctr As Long, colActtyp As Long

    colMaterial = nDescCols + 1
    colPrice = colMaterial + 1
    colQty = colPrice + 1
    colTotal = colQty + 1
    colManufMat = colTotal + 1
    colTradedMat = colManufMat + 1
    colManufCostTot = colTradedMat + 1
    colMatOH = colManufCostTot + 1
    colManuOH = colMatOH + 1
    colVariable = colManuOH + 1
    colFixed = colVariable + 1
    colWorkctr = colFixed + 1
    colActtyp = colWorkctr + 1

    Dim numericCols(0 To 7) As Long
    numericCols(0) = colTotal
    numericCols(1) = colManufMat
    numericCols(2) = colTradedMat
    numericCols(3) = colManufCostTot
    numericCols(4) = colMatOH
    numericCols(5) = colManuOH
    numericCols(6) = colVariable
    numericCols(7) = colFixed

    ' --- Title / meta block ---
    Dim startRow As Long
    startRow = 1
    If metaDict.Count > 0 Then
        If metaDict.Exists("Material") Then
            ws.Cells(1, 1).value = Trim(metaDict("Material"))
            ws.Cells(1, 1).Font.Bold = True
            ws.Cells(1, 1).Font.Size = 14
        End If
        Dim infoBits As String
        Dim metaKeysToShow As Variant
        metaKeysToShow = Array("Plant", "Costing Variant", "Costing Date from-to", "Lot Size", "Created By", "Entered On")
        Dim mk As Long
        infoBits = ""
        For mk = LBound(metaKeysToShow) To UBound(metaKeysToShow)
            If metaDict.Exists(CStr(metaKeysToShow(mk))) Then
                If infoBits <> "" Then infoBits = infoBits & "   |   "
                infoBits = infoBits & metaKeysToShow(mk) & ": " & metaDict(CStr(metaKeysToShow(mk)))
            End If
        Next mk
        ws.Cells(2, 1).value = infoBits
        ws.Cells(2, 1).Font.Size = 10
        ws.Cells(2, 1).Font.Color = RGB(89, 89, 89)
        startRow = 4
    End If

    ' --- Header row ---
    Dim headerRow As Long
    headerRow = startRow
    Dim descHeaders() As String
    ReDim descHeaders(0 To nDescCols - 1)
    descHeaders(0) = "Costing Structure"
    If nDescCols >= 2 Then descHeaders(1) = "Material Description"
    If nDescCols >= 3 Then descHeaders(2) = "Component"
    If nDescCols >= 4 Then descHeaders(3) = "Sub-component"
    Dim extra As Long
    For extra = 4 To nDescCols - 1
        descHeaders(extra) = "Level " & (extra + 1)
    Next extra

    Dim c As Long
    For c = 0 To nDescCols - 1
        ws.Cells(headerRow, c + 1).value = descHeaders(c)
    Next c
    ws.Cells(headerRow, colMaterial).value = "Material"
    ws.Cells(headerRow, colPrice).value = "Price"
    ws.Cells(headerRow, colQty).value = "Quantity"
    ws.Cells(headerRow, colTotal).value = "Total value"
    ws.Cells(headerRow, colManufMat).value = "Manuf.material"
    ws.Cells(headerRow, colTradedMat).value = "Traded material"
    ws.Cells(headerRow, colManufCostTot).value = "Manuf. Costs total"
    ws.Cells(headerRow, colMatOH).value = "Mat. overhead"
    ws.Cells(headerRow, colManuOH).value = "Manu overhead"
    ws.Cells(headerRow, colVariable).value = "Variable value"
    ws.Cells(headerRow, colFixed).value = "Fixed Value"
    ws.Cells(headerRow, colWorkctr).value = "Work Ctr"
    ws.Cells(headerRow, colActtyp).value = "ActTyp"

    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, colActtyp))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    Dim dataStart As Long
    dataStart = headerRow + 1

    ' --- Data rows ---
    For i = 0 To n - 1
        Dim r As Long
        r = dataStart + i
        Dim descCol As Long
        descCol = rows(i).depth + 1
        ws.Cells(r, descCol).value = rows(i).Desc

        Dim isParent As Boolean
        isParent = (rows(i).ChildCount > 0)

        If rows(i).material <> "" Then ws.Cells(r, colMaterial).value = rows(i).material
        If Not IsEmpty(rows(i).Price) Then
            ws.Cells(r, colPrice).value = rows(i).Price
            ws.Cells(r, colPrice).NumberFormat = CURRENCY_FMT
        End If
        If Not IsEmpty(rows(i).Qty) Then
            ws.Cells(r, colQty).value = rows(i).Qty
            ws.Cells(r, colQty).NumberFormat = QTY_FMT
        End If
        If rows(i).workctr <> "" Then ws.Cells(r, colWorkctr).value = rows(i).workctr
        If rows(i).acttyp <> "" Then ws.Cells(r, colActtyp).value = rows(i).acttyp

        If isParent Then
            Dim colIx As Long
            For colIx = 0 To 7
                Dim anyChildHasVal As Boolean
                anyChildHasVal = False
                Dim ci As Long
                For ci = 0 To rows(i).ChildCount - 1
                    If Not IsEmpty(FieldByIndex(rows, rows(i).Children(ci), colIx)) Then anyChildHasVal = True
                Next ci

                If anyChildHasVal Then
                    Dim colLtr As String
                    colLtr = colLetter(numericCols(colIx))
                    Dim refs As String
                    refs = ""
                    For ci = 0 To rows(i).ChildCount - 1
                        Dim childR As Long
                        childR = dataStart + rows(i).Children(ci)
                        If refs <> "" Then refs = refs & ","
                        refs = refs & colLtr & childR
                    Next ci
                    ws.Cells(r, numericCols(colIx)).Formula = "=SUM(" & refs & ")"
                    ws.Cells(r, numericCols(colIx)).NumberFormat = CURRENCY_FMT
                End If
            Next colIx
        Else
            Dim vIx As Long
            For vIx = 0 To 7
                Dim leafVal As Variant
                leafVal = FieldByIndex(rows, i, vIx)
                If Not IsEmpty(leafVal) Then
                    ws.Cells(r, numericCols(vIx)).value = leafVal
                    ws.Cells(r, numericCols(vIx)).NumberFormat = CURRENCY_FMT
                End If
            Next vIx
            ' Leaf rows: Total value = Price * Quantity (live formula), when both exist
            If Not IsEmpty(rows(i).Price) And Not IsEmpty(rows(i).Qty) Then
                Dim priceLtr As String, qtyLtr As String
                priceLtr = colLetter(colPrice)
                qtyLtr = colLetter(colQty)
                ws.Cells(r, colTotal).Formula = "=" & priceLtr & r & "*" & qtyLtr & r
                ws.Cells(r, colTotal).NumberFormat = CURRENCY_FMT
            End If
        End If

        ' --- Styling by level ---
        Dim rowRange As Range
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, colActtyp))
        If rows(i).depth = 0 Then
            rowRange.Interior.Color = RGB(217, 226, 243)
            ws.Cells(r, descCol).Font.Bold = True
            ws.Cells(r, descCol).Font.Size = 12
        ElseIf rows(i).depth = 1 Then
            rowRange.Interior.Color = RGB(237, 241, 249)
            ws.Cells(r, descCol).Font.Bold = True
        ElseIf isParent Then
            ws.Cells(r, descCol).Font.Bold = True
            ws.Cells(r, descCol).Font.Italic = True
        End If

        If isParent Then
            Dim nc As Long
            For nc = 0 To UBound(numericCols)
                ws.Cells(r, numericCols(nc)).Borders(xlEdgeTop).LineStyle = xlContinuous
                ws.Cells(r, numericCols(nc)).Borders(xlEdgeTop).Color = RGB(191, 191, 191)
            Next nc
        End If
    Next i

    Dim dataEnd As Long
    dataEnd = dataStart + n - 1

    ' --- Font pass: Arial throughout ---
    ws.Range(ws.Cells(1, 1), ws.Cells(dataEnd, colActtyp)).Font.Name = "Arial"

    ' --- Outline levels for +/- grouping ---
    ' The root/PUMP ASSEMBLY row stays at level 0 (ungrouped) so it's the
    ' visible anchor when the group collapses - clicking the outline pane's
    ' "1" button hides every level-1+ row (all children, at every depth) and
    ' leaves just that root line showing. Handy once you've got several
    ' materials stacked or on separate sheets and want to fold each one away
    ' to a single summary line.
    For i = 0 To n - 1
        r = dataStart + i
        If rows(i).depth > 0 Then
            ws.rows(r).OutlineLevel = WorksheetFunction.Min(rows(i).depth, 7)
        End If
    Next i
    ws.Outline.SummaryRow = xlSummaryAbove   ' subtotal rows sit ABOVE their detail here

    ' --- Column widths ---
    Dim baseDescWidths As Variant
    baseDescWidths = Array(30, 34, 34, 30, 28, 28, 26)
    For c = 0 To nDescCols - 1
        Dim w As Double
        If c <= UBound(baseDescWidths) Then
            w = baseDescWidths(c)
        Else
            w = 26
        End If
        ws.Columns(c + 1).ColumnWidth = w
    Next c
    Dim otherWidths As Variant
    otherWidths = Array(12, 13, 11, 15, 16, 16, 17, 15, 15, 15, 14, 11, 9)
    For c = 0 To UBound(otherWidths)
        ws.Columns(colMaterial + c).ColumnWidth = otherWidths(c)
    Next c

    ws.rows(headerRow).RowHeight = 30

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & dataStart).Select
    ActiveWindow.FreezePanes = True
End Sub

'------------------------------------------------------------------------------
' 4. VALIDATION - flag any rollup that doesn't reconcile with SAP's own numbers
'    Returns the warning count; details go to the Immediate window (Ctrl+G).
'------------------------------------------------------------------------------
Private Function ValidateRollups(ByRef rows() As BomRow) As Long
    Dim n As Long
    n = UBound(rows) - LBound(rows) + 1

    Dim warnCount As Long
    warnCount = 0

    Dim i As Long
    For i = 0 To n - 1
        If rows(i).ChildCount > 0 Then
            warnCount = warnCount + CheckRollupCol(rows, i, "total", "Total value")
            warnCount = warnCount + CheckRollupCol(rows, i, "manufmat", "Manuf.material")
            warnCount = warnCount + CheckRollupCol(rows, i, "tradedmat", "Traded material")
            warnCount = warnCount + CheckRollupCol(rows, i, "manufcosttot", "Manuf.Costs total")
            warnCount = warnCount + CheckRollupCol(rows, i, "matoh", "Mat.overhead")
            warnCount = warnCount + CheckRollupCol(rows, i, "manuoh", "Manu overhead")
            warnCount = warnCount + CheckRollupCol(rows, i, "variable", "Variable value")
            warnCount = warnCount + CheckRollupCol(rows, i, "fixed", "Fixed Value")
        Else
            If Not IsEmpty(rows(i).Price) And Not IsEmpty(rows(i).Qty) And Not IsEmpty(rows(i).total) Then
                Dim calc As Double
                calc = CDbl(rows(i).Price) * CDbl(rows(i).Qty)
                If Abs(calc - CDbl(rows(i).total)) > 0.02 Then
                    Debug.Print "Row " & i & " (" & rows(i).Desc & ", ActTyp=" & _
                        IIf(rows(i).acttyp = "", "-", rows(i).acttyp) & "): Price*Qty=" & _
                        Format(calc, "0.00") & " vs SAP Total=" & Format(rows(i).total, "0.00") & _
                        " - usually a setup-time rate SAP doesn't display in the Price column"
                    warnCount = warnCount + 1
                End If
            End If
        End If
    Next i

    ValidateRollups = warnCount
End Function

Private Function CheckRollupCol(ByRef rows() As BomRow, i As Long, fieldName As String, label As String) As Long
    Dim own As Variant
    Select Case fieldName
        Case "total": own = rows(i).total
        Case "manufmat": own = rows(i).ManufMat
        Case "tradedmat": own = rows(i).TradedMat
        Case "manufcosttot": own = rows(i).ManufCostTot
        Case "matoh": own = rows(i).MatOH
        Case "manuoh": own = rows(i).ManuOH
        Case "variable": own = rows(i).Variable
        Case "fixed": own = rows(i).Fixed
    End Select
    If IsEmpty(own) Then
        CheckRollupCol = 0
        Exit Function
    End If

    Dim childSum As Double
    childSum = 0
    Dim ci As Long
    For ci = 0 To rows(i).ChildCount - 1
        Dim cr As Long
        cr = rows(i).Children(ci)
        Dim cv As Variant
        Select Case fieldName
            Case "total": cv = rows(cr).total
            Case "manufmat": cv = rows(cr).ManufMat
            Case "tradedmat": cv = rows(cr).TradedMat
            Case "manufcosttot": cv = rows(cr).ManufCostTot
            Case "matoh": cv = rows(cr).MatOH
            Case "manuoh": cv = rows(cr).ManuOH
            Case "variable": cv = rows(cr).Variable
            Case "fixed": cv = rows(cr).Fixed
        End Select
        If Not IsEmpty(cv) Then childSum = childSum + CDbl(cv)
    Next ci

    If Abs(CDbl(own) - childSum) > 0.02 Then
        Debug.Print "Row " & i & " (" & rows(i).Desc & ") " & label & ": SAP=" & _
            Format(CDbl(own), "0.00") & " vs children_sum=" & Format(childSum, "0.00") & _
            " (diff=" & Format(CDbl(own) - childSum, "0.00") & ")"
        CheckRollupCol = 1
    Else
        CheckRollupCol = 0
    End If
End Function

'------------------------------------------------------------------------------
' HELPERS
'------------------------------------------------------------------------------

' colIx: 0=Total 1=ManufMat 2=TradedMat 3=ManufCostTot 4=MatOH 5=ManuOH 6=Variable 7=Fixed
Private Function FieldByIndex(ByRef rows() As BomRow, idx As Long, colIx As Long) As Variant
    Select Case colIx
        Case 0: FieldByIndex = rows(idx).total
        Case 1: FieldByIndex = rows(idx).ManufMat
        Case 2: FieldByIndex = rows(idx).TradedMat
        Case 3: FieldByIndex = rows(idx).ManufCostTot
        Case 4: FieldByIndex = rows(idx).MatOH
        Case 5: FieldByIndex = rows(idx).ManuOH
        Case 6: FieldByIndex = rows(idx).Variable
        Case 7: FieldByIndex = rows(idx).Fixed
    End Select
End Function

' Parses a numeric string regardless of which regional convention it's in:
'   "102 646,21"  (SAP export: space thousands, comma decimal)
'   "145,788.42"  (Excel #,##0.00 display: comma thousands, period decimal)
'   "1234.56" / "1234,56" (plain, either style)
' Returns Empty (not 0) when the field is blank, matching "no value" in SAP.
Private Function ParseNumber(raw As String) As Variant
    Dim s As String
    s = raw
    s = Replace(s, Chr(160), " ")   ' non-breaking space
    s = Trim(s)
    If s = "" Or s = "-" Then
        ParseNumber = Empty
        Exit Function
    End If

    Dim hasComma As Boolean, hasDot As Boolean
    hasComma = (InStr(s, ",") > 0)
    hasDot = (InStr(s, ".") > 0)

    If hasComma And hasDot Then
        ' whichever separator appears LAST in the string is the decimal point
        If InStrRev(s, ",") > InStrRev(s, ".") Then
            s = Replace(s, ".", "")    ' '.' was thousands - drop it
            s = Replace(s, ",", ".")   ' ',' was decimal - make it '.'
        Else
            s = Replace(s, ",", "")    ' ',' was thousands - drop it
        End If
    ElseIf hasComma Then
        s = Replace(s, ",", ".")       ' only a comma present -> it's the decimal
    End If

    s = Replace(s, " ", "")            ' strip any remaining thousands spaces
    s = Trim(s)

    If IsNumeric(s) Then
        ParseNumber = CDbl(s)
    Else
        ParseNumber = Empty
    End If
End Function

Private Function ContainsLetter(s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = UCase(Mid(s, i, 1))
        If ch >= "A" And ch <= "Z" Then
            ContainsLetter = True
            Exit Function
        End If
    Next i
    ContainsLetter = False
End Function

' Pure arithmetic column-number -> letter conversion (no ActiveSheet dependency)
Private Function colLetter(ByVal colNum As Long) As String
    Dim n As Long, r As Long, s As String
    n = colNum
    s = ""
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    colLetter = s
End Function

' Returns the cell .Text for columns 1..lastUsedCol of row r as a 0-based
' array, mirroring the tab-split "parts" array from the raw text file.
' Returns an unallocated array if the row is entirely blank.
' Returns the raw cell .Value2 for columns 1..lastUsedCol of row r as a 0-based
' Variant array, mirroring the tab-split "parts" array from the raw text file.
' Using .Value2 (not .Text) means genuine numeric cells come through as real
' numbers regardless of how they're formatted for display - only cells that
' are actually TEXT (e.g. a raw SAP export opened as-is) go through the
' comma/space SAP-style number parser later on.
Private Function RowParts(ws As Worksheet, r As Long) As Variant()
    Dim lastCol As Long
    lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 1 And IsBlankVal(ws.Cells(r, 1).Value2) Then
        Dim emptyArr() As Variant
        RowParts = emptyArr
        Exit Function
    End If
    Dim arr() As Variant
    ReDim arr(0 To lastCol - 1)
    Dim c As Long
    For c = 1 To lastCol
        arr(c - 1) = ws.Cells(r, c).Value2
    Next c
    RowParts = arr
End Function

' True for an empty cell or a text cell containing only whitespace.
' A real number (including zero) is never "blank".
Private Function IsBlankVal(v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankVal = True
    ElseIf VarType(v) = vbString Then
        IsBlankVal = (Trim(v) = "")
    Else
        IsBlankVal = False
    End If
End Function

' A cell's real numeric value when it's already a genuine number; otherwise
' runs the SAP-style (space thousands, comma decimal) text parser on it.
Private Function ToNumber(v As Variant) As Variant
    If IsBlankVal(v) Then
        ToNumber = Empty
        Exit Function
    End If
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            ToNumber = CDbl(v)
        Case Else
            ToNumber = ParseNumber(CStr(v))
    End Select
End Function

Private Function AnyNonEmpty(arr() As Variant) As Boolean
    Dim i As Long
    On Error GoTo NoneFound
    For i = LBound(arr) To UBound(arr)
        If Not IsBlankVal(arr(i)) Then
            AnyNonEmpty = True
            Exit Function
        End If
    Next i
NoneFound:
    AnyNonEmpty = False
End Function

Private Function RowContains(arr() As Variant, needle As String) As Boolean
    Dim i As Long
    On Error GoTo NoneFound
    For i = LBound(arr) To UBound(arr)
        If VarType(arr(i)) = vbString Then
            If InStr(arr(i), needle) > 0 Then
                RowContains = True
                Exit Function
            End If
        End If
    Next i
NoneFound:
    RowContains = False
End Function

' Non-empty, trimmed text fields from an already-split row (used for meta lines,
' which are always text labels/values, never numeric)
Private Function FilterNonEmpty(arr() As Variant) As String()
    Dim outArr() As String
    Dim cnt As Long, i As Long
    cnt = 0
    ReDim outArr(0 To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        If Not IsBlankVal(arr(i)) Then
            outArr(cnt) = Trim(CStr(arr(i)))
            cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then
        Dim empty1() As String
        FilterNonEmpty = empty1
        Exit Function
    End If
    ReDim Preserve outArr(0 To cnt - 1)
    FilterNonEmpty = outArr
End Function

Private Function CleanSheetName(s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", "?", "*", "[", "]", ":")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "")
    Next i
    s = Trim(s)
    If Len(s) = 0 Then s = "BOM"
    If Len(s) > 31 Then s = Left(s, 31)
    CleanSheetName = s
End Function

Private Function AddUniqueSheet(baseName As String, wb As Workbook) As Worksheet
    Dim nm As String
    nm = baseName
    Dim n As Long
    n = 1
    Do While SheetExists(wb, nm)
        n = n + 1
        nm = Left(baseName, 31 - Len("_" & n)) & "_" & n
    Loop
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = nm
    Set AddUniqueSheet = ws
End Function

Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


