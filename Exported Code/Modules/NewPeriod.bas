Attribute VB_Name = "NewPeriod"
Sub Add_Period_Button()
'
' DO NOT TOUCH unless you know what you're doing.
'
    ReturnSheet = ActiveSheet.Name
    On Error Resume Next
        ReturnSelection = Selection.Address
        If Err.Number <> 0 Then
            ReturnSelection = "A1"
        End If
    On Error GoTo 0
    

    
    Application.ScreenUpdating = False
    With ActiveSheet.Shapes.Range(Array("Add_Period_Button"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    addPeriod
    
    Sheets(ReturnSheet).Select
    Range(ReturnSelection).Select
    Application.ScreenUpdating = True
    On Error Resume Next
        With ActiveSheet.Shapes.Range(Array("Add_Period_Button"))
            With .Shadow
                .OffsetX = 1.2246467991E-16
                .OffsetY = 2
            End With
            .ThreeD.BevelTopInset = 1
            .ThreeD.BevelTopDepth = 0.5
            .IncrementTop -1.2
        End With
        If Err.Number <> 0 Then
            periods = f.getPerArray
            Sheets(periods(UBound(periods))).Select
        End If
    On Error GoTo 0

End Sub
Function addPeriod()
Application.ScreenUpdating = False
    PrevVal = Sheets("Overview").Range(f.numToLet(f.getPerCount + 2) & "2").Value
    PrevMonth = Left(PrevVal, 3)
    
    P2Color = t.getP2Color
    BGColor = t.getBGColor
    BGFontName = t.getBGFontName
    BGFontColor = t.getBGFontColor
    BColor = t.getBColor
    BFontName = t.getBFontName
    BFontColor = t.getBFontColor
    
    NextVal = "NA"
    
    'Figure out if we're using Monthly Periods
    If PrevMonth = "Jan" Then
        NextVal = "Feb"
    ElseIf PrevMonth = "Feb" Then NextVal = "Mar"
    ElseIf PrevMonth = "Mar" Then NextVal = "Apr"
    ElseIf PrevMonth = "Apr" Then NextVal = "May"
    ElseIf PrevMonth = "May" Then NextVal = "Jun"
    ElseIf PrevMonth = "Jun" Then NextVal = "Jul"
    ElseIf PrevMonth = "Jul" Then NextVal = "Aug"
    ElseIf PrevMonth = "Aug" Then NextVal = "Sep"
    ElseIf PrevMonth = "Sep" Then NextVal = "Oct"
    ElseIf PrevMonth = "Oct" Then NextVal = "Nov"
    ElseIf PrevMonth = "Nov" Then NextVal = "Dec"
    ElseIf PrevMonth = "Dec" Then NextVal = "Jan"
    End If
    
    ' If using Months Then...
    If NextVal <> "NA" Then
        PrevYear = CInt(Right(PrevVal, 4))
        If NextVal = "Jan" Then
            NextVal = NextVal & " " & PrevYear + 1
        Else
            NextVal = NextVal & " " & PrevYear
        End If
    
    'If not using Months Then...
    Else
        PrevDates = Split(PrevVal, " to ")
        NextDateStart = CDate(PrevDates(1)) + 1
        NextDateEnd = NextDateStart + (CDate(PrevDates(1)) - CDate(PrevDates(0)))
        NextVal = Format(NextDateStart, "m-d") & " to " & Format(NextDateEnd, "m-d")
    End If
    
    
    ' Create period sheet
    NewSheetIndex = Sheets(PrevVal).Index
    
    Sheets("Interval").Visible = True
    Sheets("Interval").Select
    Sheets("Interval").Copy Before:=Sheets(NewSheetIndex)
    Sheets("Interval (2)").Select
    Sheets("Interval (2)").Name = NextVal
    Range("A1:P1").Select
    ActiveCell.FormulaR1C1 = "'" & NextVal
    Range("A2").Select
    Sheets("Interval").Visible = False
    
    ' Add to Overview
    Sheets("Overview").Select
    Range(f.numToLet(f.getPerCount + 3) & "2").Value = "'" & NextVal
    Range(f.numToLet(f.getPerCount + 4) & "2").Value = "Totals"
    

    
    ' Remove Add_New_Period button from old sheet and change rowheight
    periods = f.getPerArray

    Sheets(periods(UBound(periods) - 1)).Shapes.Range(Array("Add_Period_Button")).Delete
    Sheets(periods(UBound(periods) - 1)).Rows("2").RowHeight = 15
    ' Change "Current" to "End" on old sheet
    Sheets(periods(UBound(periods) - 1)).Range("K3:L3").FormulaR1C1 = "End"
    
        
    ' Resize Row 2 to fit Add Period Button
    'Sheets(periods(UBound(periods))).Shapes.Range ("Add_Period_Button")
    Sheets(periods(UBound(periods))).Rows("2").RowHeight = _
        Sheets(periods(UBound(periods))).Shapes.Range("Add_Period_Button").Height + 8
    Sheets(periods(UBound(periods))).Shapes.Range("Add_Period_Button").Top = _
        Sheets(periods(UBound(periods))).Range("A2").Top
    
    ' Rerender Overview
    OverviewPage.render

                    
    
    ' Rerender PeriodSheets
    Call PeriodSheets.RenderPeriod(CStr(NextVal), CStr(PrevVal))
    

        
    ' Apply Theming to only this page
    Dim NextValString As String
    NextValString = NextVal
    endRow = f.getRowCount(NextValString)
    
    BGRange = "A1:P2,A3:A" & endRow & ",B" & endRow & ":G" & endRow _
                & ",G3:G" & endRow - 1 & ",P3:P" & endRow
    P1Range = "B4:F" & endRow - 1
    
    With Sheets(NextVal).Range(BGRange)
        .Interior.Color = BGColor
        .Font.Name = BGFontName
        .Font.Color = BGFontColor
    End With
    
    With Sheets(NextVal).Range("B3:F3")
        .Interior.Color = P2Color
        .Font.Name = t.getP2FontName
        .Font.Color = t.getP2FontColor
    End With
    
    With Sheets(NextVal).Range(P1Range)
        .Interior.Color = t.getP1Color
        .Font.Name = t.getP1FontName
        .Font.Color = t.getP1FontColor
    End With
    
    Sheets(NextVal).Range("B2:F2").Borders(xlEdgeBottom).Color = P2Color
    Sheets(NextVal).Range("A3:A" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(NextVal).Range("B3:B" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(NextVal).Range("C3:C" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(NextVal).Range("D3:D" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(NextVal).Range("E3:E" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(NextVal).Range("F3:F" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(NextVal).Range("B" & endRow - 1 & ":F" & endRow - 1).Borders(xlEdgeBottom).Color = P2Color
    
    On Error Resume Next
        Sheets(NextVal).Rows("2").RowHeight = _
            Sheets(NextVal).Shapes.Range("Add_Period_Button").Height + 14
        With Sheets(NextVal).Shapes.Range("Add_Period_Button")
            .Fill.ForeColor.RGB = BColor
            With .TextFrame2.TextRange.Font
                .Fill.ForeColor.RGB = BFontColor
                .Name = BFontName
            End With
            .Top = Sheets(NextVal).Range("A2").Top + 7
        End With
        If Err.Number <> 0 Then
            Sheets(PeriodSheet).Rows("2").RowHeight = 15
        End If
    On Error GoTo 0
    
    With Sheets(NextVal).Shapes.Range("Add_Row_Button")
        .Fill.ForeColor.RGB = BColor
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = BFontColor
            .Name = BFontName
        End With
    End With
    
    With Sheets(NextVal).Shapes.Range("Goto_Overview_Button")
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = BGFontColor
            .Name = BFontName
        End With
    End With
    
    With Sheets(NextVal).ChartObjects("Bar_Chart").Chart
        With .Axes(xlValue)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .MajorGridlines.Format.Line.ForeColor.RGB = BGFontColor
            .MajorGridlines.Format.Line.Transparency = 0.5
        End With
        With .Axes(xlCategory)
            .TickLabels.Font.Name = BGFontName
            .TickLabels.Font.Color = BGFontColor
            .Format.Line.ForeColor.RGB = BGFontColor
            .MajorGridlines.Format.Line.Transparency = 0.3
        End With
    End With
    
    Sheets(NextVal).ChartObjects("Pie_Chart").Chart.FullSeriesCollection(1).Format.Line.ForeColor.RGB _
            = BGColor
            
    ButtonPos = (Sheets(periods(UBound(periods))).Range("A2:P2").Width / 2) - _
        (Sheets(periods(UBound(periods))).Shapes.Range("Add_Period_Button").Width / 2)
    ' Position Add Period Button
    Sheets(periods(UBound(periods))).Shapes.Range("Add_Period_Button").Left = 19
    

    
End Function


