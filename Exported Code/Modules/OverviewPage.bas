Attribute VB_Name = "OverviewPage"
Function render()
    ' Re-render the entire chart area of the Overview Page
    ' Using getCatArray, getActArray and getPerArray
    Application.ScreenUpdating = False
    
    ReturnSheet = ActiveSheet.Name
    On Error Resume Next
        ReturnSelection = Selection.Address
        If Err.Number <> 0 Then
            ReturnSelection = "A1"
        End If
    On Error GoTo 0
    Sheets("Overview").Select
    
    CurrentCatCount = f.getCatCount
    CurrentCatArray = f.getCatArray
    CurrentActCount = f.getActCount
    CurrentActArray = f.getActArray
    CurrentPerCount = f.getPerCount
    
    P1Color = t.getP1Color
    P1FontName = t.getP1FontName
    P1FontColor = t.getP1FontColor
    P2Color = t.getP2Color
    P2FontName = t.getP2FontName
    P2FontColor = t.getP2FontColor
    P3Color = t.getP3Color
    BGColor = t.getBGColor
    BGFontName = t.getBGFontName
    BGFontColor = t.getBGFontColor
    BColor = t.getBColor
    BFontName = t.getBFontName
    BFontColor = t.getBFontColor
    
    ' Resetting
    With Sheets("Overview").Cells.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = BGColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("Overview").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlEdgeLeft).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlEdgeTop).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlEdgeBottom).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlEdgeRight).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("Overview").Cells.Borders(xlInsideHorizontal).LineStyle = xlNone
    Sheets("Overview").Rows("3:" & f.getRowCount("Overview")).ClearContents
    
    
    
    
    ' Fix Hidden Columns
    Cells.Select
    Selection.EntireColumn.Hidden = False
    
    Columns(f.numToLet(CurrentPerCount + 5)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    
    
    ' Correct Widths and Overview range
    HorizontalRows = CurrentPerCount + 4
    
    Rows("1:1").MergeCells = False
    Range("B1:" & f.numToLet(HorizontalRows - 1) & "1").MergeCells = True
    
    Columns("A").ColumnWidth = 2
    Columns("B").ColumnWidth = 15
    Columns("C:" & f.numToLet(HorizontalRows - 1)).ColumnWidth = 12
    Columns(f.numToLet(HorizontalRows + 0)).ColumnWidth = 2
    
    Rows("2:2").RowHeight = 30
    
    ' Overview text
     With Sheets("Overview").Range("B1:D1")
        .Interior.Color = BGColor
        .Font.Name = BGFontName
        .Font.Color = BGFontColor
     End With
    
    
    '--------------------------------- LABELS LABELS ------------------------------------
    Sheets("Overview").Range("B:B").Interior.Color = BGColor
    
    With Sheets("Overview").Range("B2")
        .Interior.Color = P2Color
        .Font.Name = P2FontName
        .Font.Color = P2FontColor
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
    End With
    
    ' Render category labels
    LoopIndex = 0
    For Each Category In CurrentCatArray
        With Sheets("Overview").Range("B" & LoopIndex + 3)
            .FormulaR1C1 = Category
            .Interior.Color = P2Color
            .Font.Name = P2FontName
            .Font.Color = P2FontColor
            .Font.Bold = False
            .Font.Underline = xlUnderlineStyleNone
        End With
        LoopIndex = LoopIndex + 1
    Next
    
    ' Render account labels
    For Each Account In CurrentActArray
        With Sheets("Overview").Range("B" & LoopIndex + 3)
            .FormulaR1C1 = Account
            .Interior.Color = P2Color
            .Font.Name = P2FontName
            .Font.Color = P2FontColor
            .Font.Bold = False
            .Font.Underline = xlUnderlineStyleNone
        End With
        LoopIndex = LoopIndex + 1
    Next
    
    With Sheets("Overview").Range("B" & LoopIndex + 3)
        .FormulaR1C1 = "Total Gained"
        .Interior.Color = P2Color
        .Font.Name = P2FontName
        .Font.Color = P2FontColor
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
    End With
    With Sheets("Overview").Range("B" & LoopIndex + 4)
        .FormulaR1C1 = "Total Spent"
        .Interior.Color = P2Color
        .Font.Name = P2FontName
        .Font.Color = P2FontColor
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
    End With
    With Sheets("Overview").Range("B" & LoopIndex + 5)
        .FormulaR1C1 = "Net Gain/Loss"
        .Interior.Color = P2Color
        .Font.Name = P2FontName
        .Font.Color = P2FontColor
        .Font.Bold = True
    End With
    
    '--------------------------------- LABELS COLUMN ------------------------------------
    
    
    
    '---------------------------------- PERIOD COLUMNS ----------------------------------
    PeriodIndex = 0
    For Each Period In f.getPerArray()
        ThisColumn = f.numToLet(3 + PeriodIndex)
        
        ' Period Label
        With Sheets("Overview").Range(ThisColumn & "2")
            .FormulaR1C1 = "'" & Period
            .Interior.Color = P2Color
            .Font.Name = P2FontName
            .Font.Color = P2FontColor
            .Font.Bold = False
            .Font.Underline = xlUnderlineStyleNone
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        ' Categories
        RowIndex = 3
        For Each Category In getCatArray()
            ThisRow = RowIndex
            ThisRowInPeriodSheet = CurrentActCount + 4 + RowIndex
            ' Add error  boundary if sheet doesnt exist
            With Sheets("Overview").Range(ThisColumn & ThisRow)
                .Formula = "='" & Period & "'!I" & ThisRowInPeriodSheet
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
                .Font.Bold = False
                .Font.Underline = xlUnderlineStyleNone
            End With
            RowIndex = RowIndex + 1
        Next
        
        ' Accounts
        AccountIndex = 0
        For Each Account In getActArray()
            ThisRow = RowIndex
            ThisRowInPeriodSheet = 4 + AccountIndex
            ' Add error boundary if sheet doesnt exist
            With Sheets("Overview").Range(ThisColumn & ThisRow)
                .Formula = "='" & Period & "'!L" & ThisRowInPeriodSheet
                .Interior.Color = P1Color
                .Font.Name = P1FontName
                .Font.Color = P1FontColor
                .Font.Bold = False
                .Font.Underline = xlUnderlineStyleNone
            End With
            RowIndex = RowIndex + 1
            AccountIndex = AccountIndex + 1
        Next
        
        ' Total Gained
        ThisRow = ThisRow + 1
        With Sheets("Overview").Range(ThisColumn & ThisRow)
            .Formula = "=SUMIF(" & ThisColumn & "3" & ":" & ThisColumn _
                    & 2 + CurrentCatCount _
                    & ", " & Chr(34) & ">0" & Chr(34) & ")"
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = False
            .Font.Underline = xlUnderlineStyleNone
        End With
        
        ' Total Spent
        ThisRow = ThisRow + 1
        With Sheets("Overview").Range(ThisColumn & ThisRow)
            .Formula = "=SUMIF(" & ThisColumn & "3" & ":" & ThisColumn _
                    & 2 + CurrentCatCount _
                    & ", " & Chr(34) & "<0" & Chr(34) & ")"
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = False
            .Font.Underline = xlUnderlineStyleNone
        End With
        
        ' Net Gain/Loss
        ThisRow = ThisRow + 1
        With Sheets("Overview").Range(ThisColumn & ThisRow)
            .Formula = "=" & ThisColumn & ThisRow - 2 & "+" & ThisColumn & ThisRow - 1
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = True
            .Font.Underline = xlUnderlineStyleSingleAccounting
        End With
        
        
        PeriodIndex = PeriodIndex + 1
        
        ' Create "GoTo" buttons
        On Error Resume Next
            Sheets("Overview").Shapes.Range("GOTO |" & Period & "|").Delete
        On Error GoTo 0
        Set NewGotoButton = Sheets("Overview").Shapes.Range("GOTO TEMPLATE").Duplicate
        
        NewGotoButton.Visible = msoTrue
        NewGotoButton.Name = "GOTO |" & Period & "|"
        NewGotoButton.Left = Sheets("Overview").Range(ThisColumn & "2").Left
        NewGotoButton.Top = Sheets("Overview").Range(ThisColumn & "2").Top
        NewGotoButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB _
                = P2FontColor
        
    Next
    '---------------------------------- PERIOD COLUMNS ----------------------------------
    
    
    '----------------------------------- TOTAL COLUMN -----------------------------------
    ThisColumn = f.numToLet(3 + CurrentPerCount)
    RowsCount = CurrentActCount + CurrentCatCount + 3 + 2
    RangeEnd = f.numToLet(2 + CurrentPerCount)
    CatRangeEnd = CurrentCatCount + 2
    ActRangeEnd = CurrentActCount + CatRangeEnd
    
    ' Label
    With Sheets("Overview").Range(ThisColumn & "2")
        .FormulaR1C1 = "Totals"
        .Interior.Color = P2Color
        .Font.Name = P2FontName
        .Font.Color = P2FontColor
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Rows
    For RowIndex = 3 To RowsCount Step 1
        
        With Sheets("Overview").Range(ThisColumn & RowIndex)
            .Interior.Color = P1Color
            .Font.Name = P1FontName
            .Font.Color = P1FontColor
            .Font.Bold = False
            .Font.Underline = xlUnderlineStyleNone
            If RowIndex > CatRangeEnd And RowIndex <= ActRangeEnd Then
                .FormulaR1C1 = "=RC[-1]"
            Else
                .Formula = "=SUM(C" & RowIndex & ":" & RangeEnd & RowIndex & ")"
            End If
            
            
            If RowIndex = RowsCount Then
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingleAccounting
            End If
        End With
    Next
    '----------------------------------- TOTAL COLUMN -----------------------------------
    
    
    
    '------------------------------------- BORDERS --------------------------------------
    ChartRange = "B2:" + f.numToLet(CurrentPerCount + 3) & _
        CurrentCatCount + CurrentActCount + 5

    
    ' Thin
    Sheets("Overview").Range(ChartRange).Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("Overview").Range(ChartRange).Borders(xlDiagonalUp).LineStyle = xlNone
    With Sheets("Overview").Range(ChartRange).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets("Overview").Range(ChartRange).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets("Overview").Range(ChartRange).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets("Overview").Range(ChartRange).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets("Overview").Range(ChartRange).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets("Overview").Range(ChartRange).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Medium
    With Sheets("Overview").Range("B1:" & ThisColumn & "1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    ' Thick
    CatLowerRange = "B" & CurrentCatCount + 2 & ":" _
            & ThisColumn & CurrentCatCount + 2
    ActLowerRange = "B" & CurrentCatCount + CurrentActCount + 2 _
            & ":" & ThisColumn & CurrentCatCount + CurrentActCount + 2
    With Sheets("Overview").Range(CatLowerRange & ", " & ActLowerRange).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
    TotalsRange = ThisColumn & "2:" & ThisColumn & CurrentCatCount + _
            CurrentActCount + 5
    With Sheets("Overview").Range(TotalsRange).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = P3Color
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
    checkRowCount
    
    ' Position Add Period Button
    Sheets("Overview").Shapes.Range(Array("Add_Period_Button")).Top = _
            Sheets("Overview").Range("A" & _
                CurrentActCount + CurrentCatCount + 6 _
            ).Top + 3
    Sheets("Overview").Shapes.Range(Array("Add_Period_Button")).Left = _
            Sheets("Overview").Range( _
                f.numToLet(CurrentPerCount + 2) & _
                CurrentActCount + CurrentCatCount + 6 _
            ).Left + 2
    Sheets("Overview").Shapes.Range(Array("Add_Period_Button")).Width = _
            Sheets("Overview").Range( _
                f.numToLet(CurrentPerCount + 2) & _
                CurrentActCount + CurrentCatCount + 6 _
                & ":" & _
                f.numToLet(CurrentPerCount + 3) & _
                CurrentActCount + CurrentCatCount + 6 _
            ).Width - 3.5
            
    With Sheets("Overview").Shapes.Range(Array("Add_Period_Button"))
        .Fill.ForeColor.RGB = BColor
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BFontColor
        .TextFrame2.TextRange.Font.Name = BFontName
    End With

    With Sheets("Overview").Shapes.Range(Array("Spending_Chart"))
        .Top = Sheets("Overview").Shapes.Range("Add_Period_Button").Top + _
                Sheets("Overview").Shapes.Range("Add_Period_Button").Height + 4
            
        .Height = 300
        .Width = Sheets("Overview").Range("B2:" & f.numToLet(CurrentPerCount + 3) & "2").Width
        .Left = Sheets("Overview").Range("B8").Left
        
    End With
    
    With Sheets("Overview").Shapes.Range(Array("Earning_Chart"))
        .Top = Sheets("Overview").Shapes.Range("Spending_Chart").Top + _
                Sheets("Overview").Shapes.Range("Spending_Chart").Height
        .Height = 250
        .Width = Sheets("Overview").Range("B2:" & f.numToLet(CurrentPerCount + 3) & "2").Width
        .Left = Sheets("Overview").Range("B8").Left
    End With
    
    With Sheets("Overview").Shapes.Range(Array("Accounts_Chart"))
        .Top = Sheets("Overview").Shapes.Range("Earning_Chart").Top + _
                Sheets("Overview").Shapes.Range("Earning_Chart").Height
        .Height = 300
        .Width = Sheets("Overview").Range("B2:" & f.numToLet(CurrentPerCount + 3) & "2").Width
        .Left = Sheets("Overview").Range("B8").Left
    End With
    
                
                
    LastPerColumn = f.numToLet(CurrentPerCount + 2)
    
    With Sheets("Overview").ChartObjects("Spending_Chart").Chart
        .FullSeriesCollection(1).Smooth = True
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Name = BGFontName
            .Fill.ForeColor.RGB = BGFontColor
        End With
        With .Axes(xlValue)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .MajorGridlines.Format.Line.ForeColor.RGB = P2Color
        End With
        With .Axes(xlCategory)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .Format.Line.ForeColor.RGB = P2Color
        End With
        With .Legend.Format.TextFrame2.TextRange.Font
            .Name = BGFontName
            .Fill.ForeColor.RGB = BGFontColor
        End With
        With .PlotArea.Format
            .Fill.ForeColor.RGB = P1Color
            .Line.ForeColor.RGB = P2Color
        End With
        
        ' * Set data range *
        .SetSourceData Source:=Sheets("Overview").Range("B2:" & _
                LastPerColumn & _
                CurrentCatCount + 2 _
        )
    End With
    
    
    With Sheets("Overview").ChartObjects("Earning_Chart").Chart
        '.FullSeriesCollection(1).Smooth = True
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Name = BGFontName
            .Fill.ForeColor.RGB = BGFontColor
        End With
        With .Axes(xlValue)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .MajorGridlines.Format.Line.ForeColor.RGB = P2Color
        End With
        With .Axes(xlCategory)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .Format.Line.ForeColor.RGB = P2Color
        End With
        With .PlotArea.Format
            .Fill.ForeColor.RGB = P1Color
            .Line.ForeColor.RGB = P2Color
        End With
        
        ' * Set data range *
        .SetSourceData Source:=Sheets("Overview").Range("B2:" & _
                LastPerColumn & _
                CurrentCatCount + 2 _
        )
        
    End With
    
    
    With Sheets("Overview").ChartObjects("Accounts_Chart").Chart
        '.FullSeriesCollection(1).Smooth = True
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Name = BGFontName
            .Fill.ForeColor.RGB = BGFontColor
        End With
        With .Axes(xlValue)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .MajorGridlines.Format.Line.ForeColor.RGB = P2Color
        End With
        With .Axes(xlCategory)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .Format.Line.ForeColor.RGB = P2Color
        End With
        With .PlotArea.Format
            .Fill.ForeColor.RGB = P1Color
            .Line.ForeColor.RGB = P2Color
        End With
        With .Legend.Format.TextFrame2.TextRange.Font
            .Name = BGFontName
            .Fill.ForeColor.RGB = BGFontColor
        End With
        
        ' * Set data range *
        .SetSourceData Source:=Sheets("Overview").Range( _
            "B2:" & LastPerColumn & "2," & _
            "B" & CurrentCatCount + 3 & ":" & LastPerColumn & CurrentCatCount + 2 + CurrentActCount _
        )
        
    End With
    
    
    
    If CurrentPerCount = 1 Then
        Sheets("Overview").ChartObjects("Spending_Chart").Chart.PlotBy = xlColumns
        Sheets("Overview").ChartObjects("Earning_Chart").Chart.PlotBy = xlColumns
        Sheets("Overview").ChartObjects("Accounts_Chart").Chart.PlotBy = xlColumns
    Else
        Sheets("Overview").ChartObjects("Spending_Chart").Chart.PlotBy = xlRows
        Sheets("Overview").ChartObjects("Earning_Chart").Chart.PlotBy = xlRows
        Sheets("Overview").ChartObjects("Accounts_Chart").Chart.PlotBy = xlRows
    End If
    
    ' * Smooth Spending & Earning Graph
    If CurrentPerCount > 1 Then
        For i = 1 To CurrentCatCount Step 1
            If Sheets("Overview").Range(f.numToLet(CurrentPerCount + 3) & i + 2).Value > 0 Then
                Sheets("Overview").ChartObjects("Earning_Chart").Chart.FullSeriesCollection(i).Smooth = True
                Sheets("Overview").ChartObjects("Spending_Chart").Chart.FullSeriesCollection(i).Smooth = False
            Else
                Sheets("Overview").ChartObjects("Earning_Chart").Chart.FullSeriesCollection(i).Smooth = False
                Sheets("Overview").ChartObjects("Spending_Chart").Chart.FullSeriesCollection(i).Smooth = True
            End If
        Next
    End If
    
    ' * Smooth Accounts Graph
    If CurrentPerCount > 1 Then
        For i = 1 To CurrentActCount Step 1
            Sheets("Overview").ChartObjects("Accounts_Chart").Chart.FullSeriesCollection(i).Smooth = True
        Next
    End If
    
    Range("B1:D1").Select
    Sheets(ReturnSheet).Select
    Range(ReturnSelection).Select

End Function
Function checkRowCount()

    bottomRow = f.getRowCount("Overview")

    CurrentHeight = Sheets("Overview").Range("A" & _
            f.getCatCount + f.getActCount + 6 & _
            ":A" & bottomRow _
        ).Height
    
    HeightNeeded = Sheets("Overview").Shapes.Range("Add_Period_Button").Height + _
            Sheets("Overview").Shapes.Range("Spending_Chart").Height + _
            Sheets("Overview").Shapes.Range("Earning_Chart").Height + _
            Sheets("Overview").Shapes.Range("Accounts_Chart").Height + 60
    
        
    Dim cellHeightNeeded As Double
    cellHeightNeeded = CurrentHeight - HeightNeeded
    
    
    If cellHeightNeeded > 15 Then
        cellsToHide = Int(cellHeightNeeded / 15)
        For i = 1 To cellsToHide Step 1
            Sheets("Overview").Rows(bottomRow - i + 1).EntireRow.Hidden = True
        Next
    ElseIf cellHeightNeeded < -15 Then
        cellsToAdd = Int(-(cellHeightNeeded / 15))
        For i = 1 To cellsToAdd Step 1
            Sheets("Overview").Rows(bottomRow + i).EntireRow.Hidden = False
        Next
    
    End If
    
        

End Function
Sub Goto_Button()
    On Error Resume Next
        GotoSheet = Split(ActiveSheet.Shapes(Application.Caller).Name, "|")(1)
        Sheets(GotoSheet).Select
    On Error GoTo 0
End Sub



