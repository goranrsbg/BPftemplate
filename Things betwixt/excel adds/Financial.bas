Attribute VB_Name = "Financial"
Sub AddPerson()
       
    Dim PProw As Integer ' Personnel Plan table row
    Dim SProw As Integer ' Designated Salary per Position table row
    Dim EProw As Integer ' Number of Employees per Position table row
    
    If (Not isHere("Personnel")) Then
        MsgBox "Sheet Personnel is not found."
        Exit Sub
    End If
    
    On Error GoTo er
    
    GetPersonnelRows PProw, SProw, EProw
    
    Cells(PProw - 1, 1).EntireRow.Insert Shift:=xlDown
    Cells(SProw - 1, 1).EntireRow.Insert Shift:=xlDown
    Cells(EProw - 1, 1).EntireRow.Insert Shift:=xlDown
    PProw = PProw + 3
    SProw = SProw + 2
    EProw = EProw + 1
    Cells(EProw - 1, 1).EntireRow.Copy
    Cells(EProw - 2, 1).EntireRow.PasteSpecial xlPasteFormulas
    Cells(SProw - 1, 1).EntireRow.Copy
    Cells(SProw - 2, 1).EntireRow.PasteSpecial xlPasteFormulas
    Cells(PProw - 1, 1).EntireRow.Copy
    Cells(PProw - 2, 1).EntireRow.PasteSpecial xlPasteFormulas
    
    Application.CutCopyMode = False
    
    For Each cell In Columns(1).Cells
        If IsEmpty(cell) Then
            cell.Select
            Exit For
        End If
    Next cell
    
    Exit Sub
er:
    MsgBox Err.Description
    
End Sub

Sub DeletePerson()
        
    Dim PProw As Integer
    Dim SProw As Integer
    Dim EProw As Integer
    
    If (Not isHere("Personnel")) Then
        MsgBox "Sheet Personnel is not found."
        Exit Sub
    End If
    
    On Error GoTo er
    
    GetPersonnelRows PProw, SProw, EProw
    
    Cells(PProw, 1).EntireRow.Offset(-1, 0).Delete
    Cells(SProw, 1).EntireRow.Offset(-1, 0).Delete
    Cells(EProw, 1).EntireRow.Offset(-1, 0).Delete
    
    Cells(1, 1).Select
        
    Exit Sub
er:
    MsgBox Err.Description
        
End Sub

Sub Paste()
    
    Dim FinalRow As Integer
    Dim FinalCol As Integer
    Dim rng As Range
    Dim aNames As Variant
    
    ' names for deletion from "Profit and Loss" and "Balance Sheet" tables
    aNames = Array("Total Liabilities and Capital", _
                   "Include Negative Taxes", _
                   "Sales and Marketing Expenses", _
                   "Expenses", _
                   "Total Expense", _
                   "Other Expenses:", _
                   "Other Expense", _
                   "Other Income", _
                   "Current Liabilities")
    
    On Error GoTo er
    
    ActiveSheet.Paste Range("A1")
    
    ActiveWindow.Zoom = 150
    
    Cells.Select
    
    With Selection.Font
        .name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Selection
        .RowHeight = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Cells(2, 1).value = Trim(Cells(1, 1).value)
    Cells(1, 1).value = ""
    
    ' delete unused columns
    Union(Columns("B:N"), Columns("P:AC")).Delete
    ' delete rows with specific name of the first cell or empty or zero
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = FinalRow To 1 Step -1
        Set rng = Cells(i, 1)
        If IsEmpty(rng) Or rng.value = 0 Or isInList(aNames, rng.value) Then
            rng.EntireRow.Delete
        End If
    Next i
    ' delete rows with zero values but not in Sales Forecast table
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(1, 1).value <> "Sales Forecast" Then
        For i = FinalRow To 2 Step -1
            If isAllZero(Cells(i, 2).Resize(1, 5)) Then
                Rows(i).Delete
            End If
        Next i
    End If
    ' change numberformat to be without [Red]
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = FinalRow To 2 Step -1
        For j = 2 To 6
            Cells(i, j).NumberFormat = Replace(Cells(i, j).NumberFormat, "[Red]", "")
        Next j
    Next i
    ' set table border
    Cells(1, 1).CurrentRegion.Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Font.Color = vbBlack
    ' change font color for top row to white
    FinalCol = Cells(1, Columns.Count).End(xlToLeft).Column
    Cells(1, 1).Resize(1, FinalCol).Font.Color = vbWhite
    ' change FY into Year
    For i = 2 To FinalCol
        Cells(1, i).value = Replace(Cells(1, i).value, "FY", "Year")
    Next i
    ' remove inside border for section rows
    Selection.Cells.SpecialCells(xlCellTypeBlanks).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
    FinalCol = Cells(1, Columns.Count).End(xlToLeft).Column
    Columns("A:" & ConvertToLetter(FinalCol)).AutoFit
    
    Cells(1, 1).CurrentRegion.Select
    
    Exit Sub
    
er:
    MsgBox Err.Description

End Sub

Sub Load_Feasibility_Sales_And_Marketing_Data()
    
    Dim FinalRow As Integer
    Dim MarketingRow As Integer
    Dim sp As Worksheet
    
    If Not isHere("Profit and Loss") Or Not isHere("Feasibility") Then
        MsgBox "Sheets Profit and Loss & Feasibility are not found."
        Exit Sub
    End If
    
    Set sp = Worksheets("Profit and Loss")
    
    On Error GoTo er
    
    MarketingRow = 0
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = FinalRow To 2 Step -1
        If Cells(i, 1).value = "Marketing" Then
            MarketingRow = i
            Exit For
        End If
    Next i
    
    FinalRow = sp.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = FinalRow To 2 Step -1
        Select Case Trim(sp.Cells(i, 1).value)
            Case "Sales"
                Cells(2, 2).Resize(1, 5).value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Marketing"
                Cells(MarketingRow, 2).Resize(1, 5).value = sp.Cells(i, 2).Resize(1, 5).value
        End Select
    Next i
    
    Columns("A:F").AutoFit
    
    Cells(1, 1).CurrentRegion.Select
    
    Exit Sub
er:
    MsgBox Err.Description
    
End Sub

Sub Load_Cost_Benefit_Data()
    
    Dim FinalRow As Integer
    Dim StartCount As Boolean
    Dim SubtotalIndirectCostRow As Integer
    Dim ToeSize As Integer
    Dim FirstOpEx As Integer
    Dim CurrentNumberOfRows As Integer
    Dim sp As Worksheet
    
    If (Not isHere("Profit and Loss") Or Not isHere("Cost Benefit")) Then
        MsgBox "Sheets Profit and Loss & Cost Benefit are not found."
        Exit Sub
    End If
    
    Set sp = Worksheets("Profit and Loss")
    
    On Error GoTo er
    
    ' Profit and Loss data collecting
    FinalRow = sp.Cells(Rows.Count, 1).End(xlUp).Row
    StartCount = False
    ToeSize = -1
    FirstOpEx = 0
    For i = FinalRow To 2 Step -1
        Select Case Trim(sp.Cells(i, 1).value)
            Case "Sales"
                Cells(2, 2).Resize(1, 5).value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Net Profit"
                Cells(3, 2).Resize(1, 5).value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Direct Cost of Sales"
                Cells(5, 2).Resize(1, 5).value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Total Operating Expenses"
                StartCount = True
            Case "Operating Expenses"
                StartCount = False
                FirstOpEx = i + 1
        End Select
        If StartCount Then
            ToeSize = ToeSize + 1
        End If
    Next i
    
    ' Cost Benefit table adjustments
    SubtotalIndirectCostRow = 0
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = FinalRow To 2 Step -1
        If Cells(i, 1).value = "Subtotal Indirect Cost" Then
            SubtotalIndirectCostRow = i
            Exit For
        End If
    Next i
    
    CurrentNumberOfRows = SubtotalIndirectCostRow - 8
    
    If CurrentNumberOfRows < ToeSize Then
        For i = 1 To ToeSize - CurrentNumberOfRows
            Cells(SubtotalIndirectCostRow, 1).EntireRow.Offset(-1, 0).Insert
            SubtotalIndirectCostRow = SubtotalIndirectCostRow + 1
        Next i
    ElseIf CurrentNumberOfRows > ToeSize Then
        If ToeSize < 2 Then
            ToeSize = 2
        End If
        For i = 1 To CurrentNumberOfRows - ToeSize
            Cells(SubtotalIndirectCostRow, 1).EntireRow.Offset(-1, 0).Delete
            SubtotalIndirectCostRow = SubtotalIndirectCostRow - 1
        Next i
    End If
    
    If FirstOpEx > 0 Then
        Cells(8, 1).Resize(ToeSize, 6).value = sp.Cells(FirstOpEx, 1).Resize(ToeSize, 6).value
    ElseIf FirstOpEx = 0 Then
        Cells(2, 2).Resize(2, 5).value = 0
        Cells(5, 2).Resize(1, 5).value = 0
        Cells(8, 1).Resize(2, 6).ClearContents
    End If
    
    Columns("A:F").AutoFit
    
    Cells(1, 1).CurrentRegion.Select
    
    Exit Sub
er:
    MsgBox Err.Description
    
End Sub

Sub Load_Charts_Data()
    
    Dim FinalRow As Integer
    Dim sp As Worksheet
    
    If (Not isHere("Profit and Loss") Or Not isHere("Charts")) Then
        MsgBox "Sheets Profit and Loss & Charts are not found."
        Exit Sub
    End If
    
    Set sp = Worksheets("Profit and Loss")
    
    On Error GoTo er
    
    FinalRow = sp.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = FinalRow To 2 Step -1
        Select Case Trim(sp.Cells(i, 1).value)
            Case "Sales"
                Range("M25:Q25").value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Gross Margin"
                Range("M24:Q24").value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Net Profit"
                Range("M23:Q23").value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Taxes Incurred"
                Range("M20:Q20").value = sp.Cells(i, 2).Resize(1, 5).value
            Case "Payroll Taxes"
                Range("M19:Q19").value = sp.Cells(i, 2).Resize(1, 5).value
        End Select
    Next i
    
    Columns("L:Q").AutoFit
    
    Range("L18").Select
    
    Exit Sub
er:
    MsgBox Err.Description
    
End Sub

Sub Things_Betwixt_Collect_Data()
    
    Dim iGoods As Integer
    Dim lRow As Integer
    Dim nGoods As Integer
    Dim iRow As Integer
    Dim wSheet As Worksheet
    
    If (Not isHere("Sales Forecast") Or Not isHere("Profit and Loss") Or Not isHere("Things Betwixt")) Then
        MsgBox "Sheets Things Betwixt, Sales Forecast or Profit and Loss are not found."
        Exit Sub
    End If
    
    'On Error GoTo er
    
    ' clear goods, sales, cost, other costs
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    If lRow > 1 Then
        With Range(Cells(2, 1), Cells(lRow, 6))
            .ClearContents
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            .NumberFormat = "General"
        End With
    End If
    Range("O2:S3").ClearContents   ' sales cost
    Range("O5:S5").ClearContents   ' max
    Range("O8:S11").value = 0      ' other costs
    Range("O13:S13").ClearContents ' k
    
    Set wSheet = Worksheets("Profit and Loss")
    
    ' collect from Profit and Loss
    With wSheet
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        For i = lRow To 3 Step -1
            Select Case .Cells(i, 1).value
                Case "Direct Cost of Sales"
                    Cells(8, 15).value = Cells(8, 15).value - .Cells(i, 2).value
                    Cells(8, 16).value = Cells(8, 16).value - .Range("C" & i).value
                    Cells(8, 17).value = Cells(8, 17).value - .Range("D" & i).value
                    Cells(8, 18).value = Cells(8, 18).value - .Range("E" & i).value
                    Cells(8, 19).value = Cells(8, 19).value - .Range("F" & i).value
                Case "Total Cost of Sales"
                    Range("O8:S8").value = .Range("B" & i & ":F" & i).value
                Case "Total Operating Expenses"
                    Range("O9:S9").value = .Range("B" & i & ":F" & i).value
                Case "Interest Expense"
                    Range("O10:S10").value = .Range("B" & i & ":F" & i).value
                Case "Net Other Income"
                    Range("O11:S11").value = .Range("B" & i & ":F" & i).value
            End Select
        Next i
    End With
    
    ' collect from Sales
    Set wSheet = Worksheets("Sales Forecast")
    nGoods = 0 ' nuber of goods
    iGoods = 2 ' index of first item in goods
    iRow = 3   ' first row in Salse Forecast with unit name
    With wSheet
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        If isUnitBased(.Range("A1:" & "A" & lRow)) Then
            ' count names
            While .Cells(iRow, 1).value <> "Total Unit Sales" And Not IsEmpty(.Cells(iRow, 1))
                Range("A" & iGoods).value = .Range("A" & iRow).value
                iGoods = iGoods + 1
                nGoods = nGoods + 1
                iRow = iRow + 1
            Wend
            ' add border to split unit, sale, cost
            Range("A" & iGoods & ":F" & iGoods).Borders(xlEdgeTop).LineStyle = xlDouble
            Range("A" & (iGoods + nGoods) & ":F" & (iGoods + nGoods)).Borders(xlEdgeTop).LineStyle = xlDouble
            ' add unit price
            iRow = iRow + 2
            While .Cells(iRow, 1).value <> "Sales"
                Range("A" & iGoods & ":F" & iGoods).value = .Range("A" & iRow & ":F" & iRow).value
                iGoods = iGoods + 1
                iRow = iRow + 1
            Wend
            ' add unit cost
            iRow = iRow + nGoods + 3
            While .Cells(iRow, 1).value <> "Direct Cost of Sales"
                Range("A" & iGoods & ":F" & iGoods).value = .Range("A" & iRow & ":F" & iRow).value
                iGoods = iGoods + 1
                iRow = iRow + 1
            Wend
            ' FUNCTIONS
            If nGoods > 0 Then
                ' unit sale
                With Range("B2:" & "F" & (nGoods + 1))
                    .FormulaR1C1 = "=ROUND(RC[6]*R13C[13],0)"
                    .NumberFormat = "0_);[Red](0)"
                End With
                ' Sales
                Range("O2:S2").FormulaR1C1 = "=SUMPRODUCT(RC[-13]:R[" & (nGoods - 1) & "]C[-13],R[" & nGoods & "]C[-13]:R[" & (2 * nGoods - 1) & "]C[-13])"
                ' Direct Cost of Sales
                Range("O3:S3").FormulaR1C1 = "=SUMPRODUCT(R[-1]C[-13]:R[" & (nGoods - 2) & "]C[-13],R[" & (2 * nGoods - 1) & "]C[-13]:R[" & (3 * nGoods - 2) & "]C[-13])"
                ' k value
                Range("O13:S13").FormulaR1C1 = "=(100%-R6C)*R12C/((100%-R6C-R4C)*SUMPRODUCT(R[-11]C[-7]:R[" & (nGoods - 12) & "]C[-7],R[" & (nGoods - 11) & "]C[-13]:R[" & (2 * nGoods - 12) & "]C[-13])-(100%-R6C)*SUMPRODUCT(R[-11]C[-7]:R[" & (nGoods - 12) & "]C[-7],R[" & (2 * nGoods - 11) & "]C[-13]:R[" & (3 * nGoods - 12) & "]C[-13]))"
                ' max values
                Range("O5:S5").FormulaR1C1 = "=(100%-R6C)*(SUMPRODUCT(R[-3]C[-7]:R[" & (nGoods - 4) & "]C[-7],R[" & (nGoods - 3) & "]C[-13]:R[" & (2 * nGoods - 4) & "]C[-13])-SUMPRODUCT(R[-3]C[-7]:R[" & (nGoods - 4) & "]C[-7],R[" & (2 * nGoods - 3) & "]C[-13]:R[" & (3 * nGoods - 4) & "]C[-13]))/SUMPRODUCT(R[-3]C[-7]:R[" & (nGoods - 4) & "]C[-7],R[" & (nGoods - 3) & "]C[-13]:R[" & (2 * nGoods - 4) & "]C[-13])"
            End If
            ' draw ratio border
            If nGoods > 0 Then
                Range("G2:" & "G" & (1 + nGoods)).Select
                With Selection.Interior
                    .Pattern = xlPatternRectangularGradient
                    .Gradient.RectangleLeft = 0
                    .Gradient.RectangleRight = 0
                    .Gradient.RectangleTop = 0
                    .Gradient.RectangleBottom = 0
                    .Gradient.ColorStops.Clear
                    .Gradient.ColorStops.Add(0).ThemeColor = xlThemeColorDark1
                    .Gradient.ColorStops.Add(1).ThemeColor = xlThemeColorAccent2
                End With
                Range("M2:" & "M" & (1 + nGoods)).Select
                With Selection.Interior
                    .Pattern = xlPatternRectangularGradient
                    .Gradient.RectangleLeft = 1
                    .Gradient.RectangleRight = 1
                    .Gradient.RectangleTop = 0
                    .Gradient.RectangleBottom = 0
                    .Gradient.ColorStops.Clear
                    .Gradient.ColorStops.Add(0).ThemeColor = xlThemeColorDark1
                    .Gradient.ColorStops.Add(1).ThemeColor = xlThemeColorAccent2
                End With
                Range("H2:" & "L" & (1 + nGoods)).Borders(xlInsideVertical).LineStyle = xlDot
            End If
        ElseIf isValueBased(.Range("A1:" & "A" & lRow)) Then
            MsgBox "value"
        End If
    End With
    
    ' clear ratio or add formula
    lRow = Cells(Rows.Count, 8).End(xlUp).Row
    If lRow > (nGoods + 1) Then
        With Range("G" & (nGoods + 2) & ":M" & lRow)
            .ClearContents
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Interior.ColorIndex = xlNone
        End With
    ElseIf lRow < (nGoods + 1) Then
        Range("I" & (lRow + 1) & ":L" & (nGoods + 1)).FormulaR1C1 = ("=RC[-1]")
        Range("H" & (lRow + 1) & ":H" & (nGoods + 1)).value = 0
    End If
    
    ' select ranges for user data input
    Union(Range("H2:" & "H" & (1 + nGoods)), Range("O4:S4")).Select
    
    Exit Sub
er:
    MsgBox Err.Description

End Sub

Sub Clear_Charts_Data()
    
    If (Not isHere("Charts")) Then
        MsgBox "Sheet Charts is not found."
        Exit Sub
    End If

    Union(Range("M19").Resize(2, 5), Range("M23").Resize(3, 5)).value = 0
    
    Range("L18").Select
    
End Sub

Sub GetPersonnelRows(one As Integer, two As Integer, three As Integer)
    
    Dim FinalRow As Integer
    
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    one = FinalRow - 1
    For i = one To 2 Step -1
        If IsEmpty(Cells(i, 1)) Then
            two = i
            Exit For
        End If
    Next i
    For i = two - 1 To 2 Step -1
        If IsEmpty(Cells(i, 1)) Then
            three = i - 1
            Exit For
        End If
    Next i
    
End Sub

Function isHere(name As String) As Boolean
    isHere = False
    For i = 1 To Worksheets.Count
        If (Sheets(i).name = name) Then
            isHere = True
            If Not (name = "Profit and Loss" Or name = "Sales Forecast" Or name = "Balance Sheet") Then
                Worksheets(name).Activate
            End If
            Exit Function
        End If
    Next i
End Function

Function isInList(arr As Variant, value As String) As Boolean
    isInList = False
    Dim e As Variant
    For Each e In arr
        If e = value Then
            isInList = True
            Exit Function
        End If
    Next e
End Function

Function isAllZero(rng As Range) As Boolean
    isAllZero = False
    If WorksheetFunction.CountBlank(rng) < rng.Count Then
        If WorksheetFunction.CountIf(rng, 0) = WorksheetFunction.Count(rng) Then
            isAllZero = True
            Exit Function
        End If
    End If
End Function

Function isUnitBased(rng As Range) As Boolean
    isUnitBased = False
    For Each cell In rng.Cells
        If cell.value = "Direct Unit Costs" Or cell.value = "Unit Price" Then
            isUnitBased = True
            Exit Function
        End If
    Next cell
End Function
Function isValueBased(rng As Range) As Boolean
    isValueBased = False
    For Each cell In rng.Cells
        If cell.value = "Subtotal Direct Cost of Sales" Then
            isValueBased = True
            Exit Function
        End If
    Next cell
End Function

Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function
































