Attribute VB_Name = "modClaimsExport"
'===============================================================================
' modClaimsExport: step 1 of the monthly warranty & goodwill claims process
'
' Cleans the month's claims cost export (grouped outline view) and builds a
' one-line-per-order Claims Summary. Claim numbers are then captured in column C
' (looked up per order in the sales order list), the sheet is renamed to the month
' (e.g. "October 2026") and moved into the year-to-date claims workbook, where
' step 2 (modClaimsLedger) posts it.
'===============================================================================

Option Explicit

'==============================================================================
' ProcessClaimsData  -  Warranties / Goodwill claims processing
' ----------------------------------------------------------------------------
' Raw SAP export layout (grouped outline view, subtotals below details):
'   A: Settl. Per. | B: Order | C: Plant | D: Descr. Deb.
'   E: Order Type  | F: Description | G: Value (28.00 Warranties/goodwill)
'
' Row types in the raw export:
'   DETAIL      -> Settl. Per. filled
'   SUBTOTAL    -> Settl. Per. blank, Order filled   (one per order - yellow view)
'   GRAND TOTAL -> Settl. Per. blank, Order blank    (last row)
'
' What RunAll does, in order:
'   1) Classifies every row BEFORE anything is filled (so nothing is lost)
'   2) Fills all blanks in A:F from the cell above (fill-down)
'   3) Collapses the outline to level 2 (the grouped "view 2" you use)
'   4) Builds a "Claims Summary" sheet - one row per ORDER, laid out like
'      the claims register:
'         Order | Plant | CLM NO | Descr. Deb. | Description | VALUE | Order Type
'      CLM NO (col C) is left BLANK for you to capture manually.
'      Bold TOTAL row at the bottom.
'
' HOW TO USE:
'   Open the monthly claims export, import this module (Alt+F11 > File > Import),
'   make sure the RAW DATA sheet is the active sheet, then run  RunAll  (Alt+F8)
'==============================================================================

Sub RunAll()
    Dim wsRaw As Worksheet
    Dim lastRow As Long, r As Long
    Dim rowType() As Integer   ' 1 = detail, 2 = subtotal, 3 = grand total
    
    Set wsRaw = ActiveSheet
    lastRow = wsRaw.Cells(wsRaw.rows.Count, "G").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data found on '" & wsRaw.Name & "'. Make sure the raw data sheet is active.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '--- Step 0: expand the outline fully so every row is visible -----
    On Error Resume Next
    wsRaw.Outline.ShowLevels RowLevels:=8
    On Error GoTo 0
    
    '--- Step 1: classify each row BEFORE filling blanks ----------------------
    ReDim rowType(2 To lastRow)
    For r = 2 To lastRow
        If Trim(CStr(wsRaw.Cells(r, "A").value)) <> "" Then
            rowType(r) = 1                                   ' detail
        ElseIf Trim(CStr(wsRaw.Cells(r, "B").value)) <> "" Then
            rowType(r) = 2                                   ' order subtotal
        Else
            rowType(r) = 3                                   ' grand total
        End If
    Next r
    
    '--- Step 2: fill blanks from above (fill-down) ------
    FillBlanksFromAbove wsRaw, lastRow
    
    '--- Step 3: collapse the raw sheet to grouping view 2 --------------------
    On Error Resume Next
    wsRaw.Outline.ShowLevels RowLevels:=2
    On Error GoTo 0
    
    '--- Step 4: build the Claims Summary sheet -------------------------------
    BuildClaimsSummary wsRaw, lastRow, rowType
    
    Application.ScreenUpdating = True
    
    MsgBox "Done!" & vbNewLine & vbNewLine & _
           "1. Blanks filled on '" & wsRaw.Name & "' (values, not formulas)." & vbNewLine & _
           "2. Raw sheet set to grouping view 2." & vbNewLine & _
           "3. 'Claims Summary' created - one line per order." & vbNewLine & vbNewLine & _
           "Now capture the CLM NOs in column C of the summary.", vbInformation
End Sub

'------------------------------------------------------------------------------
' Fill every blank in A2:F<lastRow> with the value of the cell above.
' NOTE: SAP exports often leave "blank" grouping cells as an empty TEXT
' STRING (""), not a truly empty cell - SpecialCells(xlCellTypeBlanks)
' does not catch those. A plain cell-by-cell loop catches both cases.
' Column G (values) is deliberately excluded so subtotal amounts stay intact.
'------------------------------------------------------------------------------
Sub FillBlanksFromAbove(ws As Worksheet, lastRow As Long)
    Dim r As Long, c As Long
    
    For r = 2 To lastRow
        For c = 1 To 6   ' columns A to F
            If Trim(CStr(ws.Cells(r, c).value)) = "" Then
                ws.Cells(r, c).value = ws.Cells(r - 1, c).value
            End If
        Next c
    Next r
End Sub

'------------------------------------------------------------------------------
' Build the summary sheet in the claims-register layout:
'   Order | Plant | CLM NO | Descr. Deb. | Description | VALUE | Order Type
' One row per order (taken from the subtotal rows), CLM NO left blank,
' bold TOTAL at the bottom with a SUM formula.
'------------------------------------------------------------------------------
Sub BuildClaimsSummary(wsRaw As Worksheet, lastRow As Long, rowType() As Integer)
    Dim wsOut As Worksheet
    Dim r As Long, outR As Long
    
    ' Recreate the summary sheet from scratch each run
    Application.DisplayAlerts = False
    On Error Resume Next
    wsRaw.Parent.Worksheets("Claims Summary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsOut = wsRaw.Parent.Worksheets.Add(After:=wsRaw)
    wsOut.Name = "Claims Summary"
    
    ' Headers
    wsOut.Range("A1:G1").value = Array("Order", "Plant", "CLM NO", "Descr.  Deb.", _
                                       "Description", "VALUE", "Order Type")
    
    outR = 2
    For r = 2 To lastRow
        Select Case rowType(r)
            Case 2  ' order subtotal -> one summary line
                wsOut.Cells(outR, "A").value = wsRaw.Cells(r, "B").value  ' Order
                wsOut.Cells(outR, "B").value = wsRaw.Cells(r, "C").value  ' Plant
                ' C = CLM NO -> left blank for manual capture
                wsOut.Cells(outR, "D").value = wsRaw.Cells(r, "D").value  ' Descr. Deb.
                wsOut.Cells(outR, "E").value = wsRaw.Cells(r, "F").value  ' Description
                wsOut.Cells(outR, "F").value = wsRaw.Cells(r, "G").value  ' VALUE
                wsOut.Cells(outR, "G").value = wsRaw.Cells(r, "E").value  ' Order Type
                outR = outR + 1
        End Select
    Next r
    
    ' TOTAL row - live SUM so it recalculates if values are ever edited
    If outR > 2 Then
        wsOut.Cells(outR, "E").value = "TOTAL"
        wsOut.Cells(outR, "F").Formula = "=SUM(F2:F" & outR - 1 & ")"
        wsOut.Range("A" & outR & ":G" & outR).Font.Bold = True
    End If
    
    FormatSummary wsOut, outR
End Sub

'------------------------------------------------------------------------------
' Formatting: bordered table, bold centred header, right-aligned values
'------------------------------------------------------------------------------
Private Sub FormatSummary(ws As Worksheet, lastRow As Long)
    With ws.Range("A1:G1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("A1:G" & lastRow)
        .Font.Name = "Arial"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    With ws.Range("F2:F" & lastRow)
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    
    ws.Range("A2:B" & lastRow).HorizontalAlignment = xlLeft
    
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 32
    ws.Columns("E").ColumnWidth = 40
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 10
    
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub



