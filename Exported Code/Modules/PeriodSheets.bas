Attribute VB_Name = "PeriodSheets"
Function render()
    Application.ScreenUpdating = False
    ReturnSheet = ActiveSheet.Name
    On Error Resume Next
        ReturnSelection = Selection.Address
        If Err.Number <> 0 Then
            ReturnSelection = "A1"
        End If
    On Error GoTo 0

    Dim PeriodArray() As String
    PeriodArray = f.getPerArray
    
    Dim PreviousPeriodSheet As String
    PreviousPeriodSheet = ""
    
    For Each PeriodSheet In PeriodArray
        Call RenderPeriod(CStr(PeriodSheet), PreviousPeriodSheet)
        
        PreviousPeriodSheet = PeriodSheet
    Next
    
    'Return to original selection
    Sheets(ReturnSheet).Select
    Range(ReturnSelection).Select
End Function
Function RenderPeriod(PeriodName As String, PrevPeriodName As String)
    IsFirstSheet = False
    If PrevPeriodName = "" Then IsFirstSheet = True
    ' render individual page
    ' if IsFirstSheet then don't reference others
    ' for new page renders
    
    P1Color = t.getP1Color
    P1FontName = t.getP1FontName
    P1FontColor = t.getP1FontColor
    P2Color = t.getP2Color
    P2FontName = t.getP2FontName
    P2FontColor = t.getP2FontColor
    P3Color = t.getP3Color
    BGColor = t.getBGColor
    
    CatArray = f.getCatArray
    CatCount = f.getCatCount
    ActCount = f.getActCount
    
    Sheets(PeriodName).Select
        
        HeightNeeded = ActiveSheet.Shapes.Range("Bar_Chart").Height + ActiveSheet.Shapes.Range("Add_Row_Button").Height + 30
        CurrentHeight = ActiveSheet.Range("A" & ActCount + CatCount + 8 & ":A" & f.getRowCount).Height
        
        If HeightNeeded > CurrentHeight Then
            RowsToAdd = Int((HeightNeeded - CurrentHeight) / 15) + 1
            For i = 0 To RowsToAdd Step 1
                addRow.addRow
            Next
        End If
        
        
        
        Dim RowCount As Integer
        RowCount = f.getRowCount
        ' Get Previous Account Balances
        If IsFirstSheet Then
            AnalyzingIndex = 0
            IndexCount = 0
            Do
                IndexCount = IndexCount + 1
            Loop While Range("H" & IndexCount + 4).Value <> "Net"
            
            Dim PrevBals() As String
            Dim PrevActs() As String
            ReDim Preserve PrevBals(IndexCount)
            ReDim Preserve PrevActs(IndexCount)
            For i = 0 To IndexCount - 1 Step 1
                PrevBals(AnalyzingIndex) = Range("I" & AnalyzingIndex + 4).Value
                PrevActs(AnalyzingIndex) = Range("H" & AnalyzingIndex + 4).Value
                AnalyzingIndex = AnalyzingIndex + 1
            Next
        End If
        
        ' Clear Previous
        With Range("H3:O" & RowCount)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = BGColor
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Name = t.getBGFontName
                .Color = t.getBGFontColor
                .Underline = xlUnderlineStyleNone
                .Bold = False
            End With
            
            .UnMerge
            .ClearContents
        End With
        
            
        ' ------------- Accounts -------------
        
        ' Start
        With Range("H3:I3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .FormulaR1C1 = "Start"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = P2Color
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Name = P2FontName
                .Color = P2FontColor
            End With
        End With
        
        pArray = f.getPerArray
        
        
        ' Current
        With Range("K3:L3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            ' Change "Current" to "End" if this is not the latest sheet
            If pArray(UBound(pArray)) = PeriodName Then
                .FormulaR1C1 = "Current"
            Else
                .FormulaR1C1 = "End"
            End If
            
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = P2Color
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Name = P2FontName
                .Color = P2FontColor
            End With
        End With
        
        ' Change
        With Range("N3:O3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .FormulaR1C1 = "Change"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = P2Color
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Name = P2FontName
                .Color = P2FontColor
            End With
        End With
        
        ' Account References
        RowIndex = 4
        For Each AccountName In f.getActArray
            With Range("H" & RowIndex)
                .FormulaR1C1 = AccountName
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
            End With
            If IsFirstSheet Then
                ItemIndex = 0
                For Each Item In PrevActs
                    If Item = AccountName Then
                        Range("I" & RowIndex).FormulaR1C1 = PrevBals(ItemIndex)
                    End If
                    ItemIndex = ItemIndex + 1
                Next
            End If
            With Range("I" & RowIndex)
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
                If IsFirstSheet = False Then
                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    .Formula = "='" & PrevPeriodName & "'!L" & RowIndex
                End If
            End With
            With Range("K" & RowIndex)
                .FormulaR1C1 = AccountName
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
            End With
            With Range("L" & RowIndex)
                .Formula = "=I" & RowIndex & "+SUMIF(F:F," & Chr(34) & AccountName & Chr(34) & ",E:E)"
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End With
            With Range("N" & RowIndex)
                .FormulaR1C1 = AccountName
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
            End With
            With Range("O" & RowIndex)
                .Formula = "=L" & RowIndex & "-I" & RowIndex
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End With
            RowIndex = RowIndex + 1
        Next
        
        ' Net Section
        With Range("H" & RowIndex)
            .FormulaR1C1 = "Net"
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        With Range("I" & RowIndex)
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            .Font.Underline = xlUnderlineStyleSingleAccounting
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Formula = "=SUM(I4:I" & RowIndex - 1 & ")"
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        With Range("K" & RowIndex)
            .FormulaR1C1 = "Net"
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        With Range("L" & RowIndex)
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            .Font.Underline = xlUnderlineStyleSingleAccounting
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Formula = "=SUM(L4:L" & RowIndex - 1 & ")"
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        With Range("N" & RowIndex)
            .FormulaR1C1 = "Net"
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        With Range("O" & RowIndex)
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            .Font.Underline = xlUnderlineStyleSingleAccounting
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Formula = "=SUM(O4:O" & RowIndex - 1 & ")"
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        
        'Add Borders
        With Range("H3:I" & RowIndex)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
        With Range("K3:L" & RowIndex)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = P3Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P3Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = P3Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = P3Color
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        
        With Range("N3:O" & RowIndex)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
        ' ------------- Categories -------------
        RowIndex = RowIndex + 2
        TopRow = RowIndex
        ' Categories Label
        With Range("H" & RowIndex & ":I" & RowIndex)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .FormulaR1C1 = "Categories"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = P2Color
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Name = P2FontName
                .Color = P2FontColor
            End With
        End With
        
        RowIndex = RowIndex + 1
        
        ' Categories Formulas
        For Each CategoryName In CatArray
            With Range("H" & RowIndex)
                .FormulaR1C1 = CategoryName
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
            End With
            With Range("I" & RowIndex)
                .Formula = "=SUMIF(D:D," & Chr(34) & CategoryName & Chr(34) & ",E:E)"
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End With
            RowIndex = RowIndex + 1
        Next
        
        ' Borders
        
        With Range("H" & TopRow & ":I" & RowIndex - 1)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = P2Color
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
        ' Position Charts
        
        With ActiveSheet.Shapes.Range(Array("Pie_Chart"))
            .Top = Range("K" & ActCount + 6).Top
            .Left = Range("K" & ActCount + 6).Left
        End With
        
        If CatCount < 7 Then
            With ActiveSheet.Shapes.Range(Array("Bar_Chart"))
                .Top = Range("H" & ActCount + 15).Top
                .Left = Range("H" & RowIndex + 1 + 6).Left
            End With
        Else
            With ActiveSheet.Shapes.Range(Array("Bar_Chart"))
                .Top = ActiveSheet.Range("A" & ActCount + CatCount + 8).Top
                .Left = Range("H" & RowIndex + 1 + 6).Left
            End With
        End If
        
        If CatCount > 7 Then
            CategoryBoxRange = "H" & ActCount + 6 & ":H" & ActCount + 6 + CatCount
        Else
            CategoryBoxRange = "H" & ActCount + 6 & ":H" & ActCount + 13
        End If
        
        ActiveSheet.Shapes.Range(Array("Pie_Chart")).Height = ActiveSheet.Range(CategoryBoxRange).Height
        
        ' Position Add Row Button
        addRow.positionAddRowButton
        
        ' Update Chart Datasets
        Range("Q1:R200").ClearContents
        For i = 1 To CatCount
            RowNumIt = i + ActCount + 6
            Range("Q" & RowNumIt).FormulaR1C1 = "=RC[-9]"
            Range("R" & RowNumIt).FormulaR1C1 = "=IF(RC[-9]<0,RC[-9],0)"
            Range("R" & RowNumIt).Style = "Currency"
        Next
        
        NewDataRange = "Q" & ActCount + 7 & ":R" _
            & CatCount + ActCount + 6
        With ActiveSheet.ChartObjects("Pie_Chart").Chart
            .SetSourceData Source:=Range(NewDataRange)
            .PlotVisibleOnly = False
        End With
        With ActiveSheet.ChartObjects("Bar_Chart").Chart
            .SetSourceData Source:=Range(NewDataRange)
            .PlotVisibleOnly = False
        End With
        
        ' Add Chart Labels
        LabelIndexNum = 1
        For Each CategoryName In CatArray
        With ActiveSheet.Range("J" & LabelIndexNum + ActCount + 6)
            .HorizontalAlignment = xlCenter
            
            .FormulaR1C1 = "=IF(RC[-1]<0,""" & ChrW(11035) & ""","""")"
            .Font.Color = ActiveSheet.ChartObjects("Pie_Chart").Chart.FullSeriesCollection(1).Points(LabelIndexNum).Interior.Color
        End With
            LabelIndexNum = LabelIndexNum + 1
        Next
        
        
        ActiveSheet.ChartObjects("Pie_Chart").Chart.FullSeriesCollection(1).Format.Line.ForeColor.RGB _
            = BGColor
        
        UpdateValidation
        Range("A1:P1").Select
    
End Function
Function UpdateValidation()
    endRow = f.getRowCount - 1
    
    
    CategoryList = ""
    CategoryIndex = 1
    For Each Category In f.getCatArray
        If CategoryIndex <> f.getCatCount Then
            CategoryList = CategoryList + Category + ","
        Else
            CategoryList = CategoryList + Category
        End If
        CategoryIndex = CategoryIndex + 1
    Next
    ' ERROR CAUSED WHEN CHART IS SELECTED
    Range("A1:P1").Select
    With Range("D4:D" & endRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=CategoryList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Invalid Category"
        .InputMessage = ""
        .ErrorMessage = _
        "Choose a valid category." & Chr(10) & "New categories can be added in the 'Control' sheet."
        .ShowInput = False
        .ShowError = True
        
    End With
    
    
    AccountList = ""
    AccountIndex = 1
    For Each Account In f.getActArray
        If AccountIndex <> f.getActCount Then
            AccountList = AccountList + Account + ","
        Else
            AccountList = AccountList + Account
        End If
        AccountIndex = AccountIndex + 1
    Next
    With Range("F4:F" & endRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=AccountList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Invalid Category"
        .InputMessage = ""
        .ErrorMessage = _
        "Choose a valid account." & Chr(10) & "New accounts can be added in the 'Control' sheet."
        .ShowInput = False
        .ShowError = True
    End With
    
    ' Date validation dropdown
    SheetNameSplit = Split(ActiveSheet.Name, " ")
    
    Dim DatesValidationString As String
    Dim IsMonthly As Boolean
    If SheetNameSplit(0) = "Jan" Or _
        SheetNameSplit(0) = "Feb" Or _
        SheetNameSplit(0) = "Mar" Or _
        SheetNameSplit(0) = "Apr" Or _
        SheetNameSplit(0) = "May" Or _
        SheetNameSplit(0) = "Jun" Or _
        SheetNameSplit(0) = "Jul" Or _
        SheetNameSplit(0) = "Aug" Or _
        SheetNameSplit(0) = "Sep" Or _
        SheetNameSplit(0) = "Oct" Or _
        SheetNameSplit(0) = "Nov" Or _
        SheetNameSplit(0) = "Dec" Then IsMonthly = True
        
    If IsMonthly Then
        If SheetNameSplit(0) = "Jan" Then
            NextVal = "Feb"
        ElseIf SheetNameSplit(0) = "Feb" Then NextVal = "Mar"
        ElseIf SheetNameSplit(0) = "Mar" Then NextVal = "Apr"
        ElseIf SheetNameSplit(0) = "Apr" Then NextVal = "May"
        ElseIf SheetNameSplit(0) = "May" Then NextVal = "Jun"
        ElseIf SheetNameSplit(0) = "Jun" Then NextVal = "Jul"
        ElseIf SheetNameSplit(0) = "Jul" Then NextVal = "Aug"
        ElseIf SheetNameSplit(0) = "Aug" Then NextVal = "Sep"
        ElseIf SheetNameSplit(0) = "Sep" Then NextVal = "Oct"
        ElseIf SheetNameSplit(0) = "Oct" Then NextVal = "Nov"
        ElseIf SheetNameSplit(0) = "Nov" Then NextVal = "Dec"
        ElseIf SheetNameSplit(0) = "Dec" Then NextVal = "Jan"
        End If
        
        DaysInMonth = CDate(NextVal & ", " & SheetNameSplit(1)) - CDate(ActiveSheet.Name)
        
        ' WeekdayName(Format(CDate(ActiveSheet.Name), "w"), True) & " " &
        DatesValidationString = Format(CDate(ActiveSheet.Name), "mmm d")

        For i = 1 To DaysInMonth - 1 Step 1
            ' WeekdayName(Format(CDate(ActiveSheet.Name) + i, "w"), True) & " " &
            DatesValidationString = DatesValidationString & "," & _
                                    Format(CDate(ActiveSheet.Name) + i, "mmm d")
            
        Next
    Else
        ' get list of dates in period
        ' WeekdayName(Format(CDate(SheetNameSplit(0)), "w"), True) & " " &
        DatesValidationString = Format(CDate(SheetNameSplit(0)), "mmm d")
        For i = 1 To CDate(SheetNameSplit(2)) - CDate(SheetNameSplit(0)) Step 1
            ' WeekdayName(Format(CDate(SheetNameSplit(0)) + i, "w"), True) & " " &
            DatesValidationString = DatesValidationString & "," & _
                                    Format(CDate(SheetNameSplit(0)) + i, "mmm d")

        Next
        
    End If
    
    With Range("B4:B" & endRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=DatesValidationString
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = False
        .ShowError = False
        
    End With
    
    
    
End Function
Sub Goto_Overview_Button()
    Sheets("Overview").Select
End Sub
Function renameCat(OldName As String, NewName As String)
    For ThisRow = 4 To f.getRowCount - 1 Step 1
        If Range("D" & ThisRow).Value = OldName Then
            Range("D" & ThisRow).Value = NewName
        End If
    Next
End Function
Function renameAct(OldName As String, NewName As String)
    For ThisRow = 4 To f.getRowCount - 1 Step 1
        If Range("F" & ThisRow).Value = OldName Then
            Range("F" & ThisRow).Value = NewName
        End If
    Next
End Function


