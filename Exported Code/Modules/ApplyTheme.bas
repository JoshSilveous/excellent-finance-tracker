Attribute VB_Name = "ApplyTheme"
Sub Apply_Theme_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Control").Select
    With ActiveSheet.Shapes.Range(Array("Apply_Theme_Button"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    Call applyTheme

    Sheets("Control").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Apply_Theme_Button"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With

End Sub
Sub applyTheme()
'   Create seperate render functions for each type of page
'   Create a function that gets category and accounts, then updates The entire right section of a period sheet

'   Prompt for confirmation
    UserInput = MsgBox("Are you sure you would like to apply this theme?" & vbNewLine _
            & "This can take a while, depending on your machine's speed. " _
            & "Please do not close Excel.", vbYesNo, _
            "Apply Theme")
    If UserInput = 6 Then
'        Call applyThemeWelcome
        Call applyThemeControl
        Call applyThemeOverview
        Call applyThemePeriods
        Call applyThemePresets
    End If
End Sub
'Sub applyThemeWelcome()
'    Sheets("Welcome").Range("A1").Interior.Color = t.getBGColor
'    Sheets("Welcome").Range("A1").Font.Name = t.getBGFontName
'    Sheets("Welcome").Range("A1").Font.Color = t.getBGFontColor
'
'    Sheets("Welcome").Shapes.Range("Welcome_Begin_Button").Fill.ForeColor.RGB _
'            = t.getBColor
'    Sheets("Welcome").Shapes.Range("Welcome_Begin_Button").Left _
'            = (Sheets("Welcome").Range("A1").Width - Sheets("Welcome").Shapes.Range("Welcome_Begin_Button").Width) / 2
'
'    Sheets("Welcome").Shapes.Range(Array("Welcome_Begin_Button")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB _
'            = t.getBFontColor
'    Sheets("Welcome").Shapes.Range(Array("Welcome_Begin_Button")).TextFrame2.TextRange.Font.Name _
'            = t.getBFontName
'End Sub
Sub applyThemeControl()
    RowCount = f.getRowCount
    
    BColor = t.getBColor
    
    Sheets("Control").Range("A1:H2,A3:A" & RowCount & ",B5:E" & RowCount & ",E4,C4,C3,E3,H3:H" & RowCount & ",F20:G" & RowCount) _
        .Interior.Color = t.getBGColor
    Sheets("Control").Range("A1:H2,A3:A" & RowCount & ",B5:E" & RowCount & ",E4,C4,C3,E3,H3:H" & RowCount & ",F20:G" & RowCount & "") _
        .Font.Name = t.getBGFontName
    Sheets("Control").Range("A1:H2,A3:A" & RowCount & ",B5:E" & RowCount & ",E4,C4,C3,E3,H3:H" & RowCount & ",F20:G" & RowCount & "") _
        .Font.Color = t.getBGFontColor
            
    Sheets("Control").Shapes.Range(Array("Edit_Category_Button", _
        "Edit_Account_Button", "Apply_Theme_Button", "Presets_Button")) _
        .Fill.ForeColor.RGB = BColor
    Sheets("Control").Shapes.Range(Array("Edit_Category_Button", _
        "Edit_Account_Button", "Apply_Theme_Button", "Presets_Button")) _
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = t.getBFontColor
    Sheets("Control").Shapes.Range(Array("Edit_Category_Button", _
        "Edit_Account_Button", "Apply_Theme_Button", "Presets_Button")) _
        .TextFrame2.TextRange.Font.Name = t.getBFontName
        
    With Sheets("Control").Shapes.Range("Apply_Theme_Button")
        .Width = Sheets("Control").Range("F2:G2").Width - 8
        .Left = Sheets("Control").Range("F2").Left + 4
    End With
    Sheets("Control").Range("F5:F19").Interior.Color = t.getP1Color
    Sheets("Control").Range("F5:F19").Font.Name = t.getP1FontName
    Sheets("Control").Range("F5:F19").Font.Color = t.getP1FontColor
    
    Sheets("Control").Range("B3:B4,D3:D4,F3:G4").Interior.Color = t.getP2Color
    Sheets("Control").Range("B3:B4,D3:D4,F3:G4").Font.Name = t.getP2FontName
    Sheets("Control").Range("B3:B4,D3:D4,F3:G4").Font.Color = t.getP2FontColor

    
    Sheets("Control").Shapes.Range("Apply_Theme_Button").Top = _
        Range("F20").Top + 4
    
    With Sheets("Control").Range("G5:G19").Borders(xlEdgeLeft)
        .Color = BColor
        .Weight = xlMedium
    End With
    With Sheets("Control").Range("G5:G19").Borders(xlEdgeTop)
        .Color = BColor
        .Weight = xlMedium
    End With
    With Sheets("Control").Range("G5:G19").Borders(xlEdgeBottom)
        .Color = BColor
        .Weight = xlMedium
    End With
    With Sheets("Control").Range("G5:G19").Borders(xlEdgeRight)
        .Color = BColor
        .Weight = xlMedium
    End With
    
    ControlPage.renderCat
    ControlPage.renderAct
End Sub
Sub applyThemeOverview()
    OverviewPage.render
    
End Sub

Sub applyThemePeriods()
For Each PeriodName In f.getPerArray
    
    P1Color = t.getP1Color
    P1FontName = t.getP1FontName
    P1FontColor = t.getP1FontColor
    P2Color = t.getP2Color
    P2FontName = t.getP2FontName
    P2FontColor = t.getP2FontColor
    BGColor = t.getBGColor
    BGFontName = t.getBGFontName
    BGFontColor = t.getBGFontColor
    BColor = t.getBColor
    BFontName = t.getBFontName
    BFontColor = t.getBFontColor

    Dim PeriodNameString As String
    PeriodNameString = PeriodName
    endRow = f.getRowCount(PeriodNameString)
    
    BGRange = "A1:P2,A3:A" & endRow & ",B" & endRow & ":G" & endRow _
                & ",G3:G" & endRow - 1 & ",P3:P" & endRow
    P1Range = "B4:F" & endRow - 1
    
    With Sheets(PeriodName).Range(BGRange)
        .Interior.Color = BGColor
        .Font.Name = BGFontName
        .Font.Color = BGFontColor
    End With
    
    With Sheets(PeriodName).Range("B3:F3")
        .Interior.Color = P2Color
        .Font.Name = P2FontName
        .Font.Color = P2FontColor
    End With
    
    With Sheets(PeriodName).Range(P1Range)
        .Interior.Color = P1Color
        .Font.Name = P1FontName
        .Font.Color = P1FontColor
    End With
    
    Sheets(PeriodName).Range("B2:F2").Borders(xlEdgeBottom).Color = P2Color
    Sheets(PeriodName).Range("A3:A" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(PeriodName).Range("B3:B" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(PeriodName).Range("C3:C" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(PeriodName).Range("D3:D" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(PeriodName).Range("E3:E" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(PeriodName).Range("F3:F" & endRow - 1).Borders(xlEdgeRight).Color = P2Color
    Sheets(PeriodName).Range("B" & endRow - 1 & ":F" & endRow - 1).Borders(xlEdgeBottom).Color = P2Color
    
    On Error Resume Next
        Sheets(PeriodName).Rows("2").RowHeight = _
            Sheets(PeriodName).Shapes.Range("Add_Period_Button").Height + 8
        With Sheets(PeriodName).Shapes.Range("Add_Period_Button")
            .Fill.ForeColor.RGB = BColor
            With .TextFrame2.TextRange.Font
                .Fill.ForeColor.RGB = BFontColor
                .Name = BFontName
            End With
            .Top = Sheets(PeriodName).Range("A2").Top + 4
        End With
        If Err.Number <> 0 Then
            Sheets(PeriodSheet).Rows("2").RowHeight = 15
        End If
    On Error GoTo 0
    
    With Sheets(PeriodName).Shapes.Range("Add_Row_Button")
        .Fill.ForeColor.RGB = BColor
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = BFontColor
            .Name = BFontName
        End With
    End With
    
    With Sheets(PeriodName).Shapes.Range("Goto_Overview_Button")
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = BGFontColor
            .Name = BFontName
        End With
    End With
    
    With Sheets(PeriodName).ChartObjects("Pie_Chart").Chart.FullSeriesCollection(1).Format
        .Line.ForeColor.RGB = BGColor
    End With
    With Sheets(PeriodName).ChartObjects("Bar_Chart").Chart
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
    
Next
PeriodSheets.render
End Sub

Sub applyThemePresets()

BColor = t.getBColor

Sheets("Theme Presets").Visible = True
Sheets("Theme Presets").Select

With ActiveSheet.Range("1:2,H3:H10,A9:G10,A3:A8")
    .Interior.Color = t.getBGColor
    .Font.Name = t.getBGFontName
    .Font.Color = t.getBGFontColor
End With
With ActiveSheet.Range("B3:G3")
    .Interior.Color = t.getP1Color
    .Font.Name = t.getP1FontName
    .Font.Color = t.getP1FontColor
End With

Dim i As Integer
For i = 2 To 7 Step 1
    With ActiveSheet.Range(f.numToLet(i) & "4:" & f.numToLet(i) & "8")
        With .Borders(xlEdgeLeft)
            .Color = BColor
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .Color = BColor
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .Color = BColor
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .Color = BColor
            .Weight = xlMedium
        End With
    End With
    
    
    With ActiveSheet.Shapes.Range("Select_Theme_" & i - 1)
        .Fill.ForeColor.RGB = BColor
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = t.getBFontColor
            .Name = t.getBFontName
        End With
        .Width = ActiveSheet.Range(f.numToLet(i) & "3").Width - 6
        .Left = ActiveSheet.Range(f.numToLet(i) & "3").Left + 3
        .Top = ActiveSheet.Range(f.numToLet(i) & "9").Top + 3
    End With
    
Next

With ActiveSheet.Shapes.Range("Exit_Button")
    .Fill.ForeColor.RGB = BColor
    With .TextFrame2.TextRange.Font
        .Fill.ForeColor.RGB = t.getBFontColor
        .Name = t.getBFontName
    End With
End With


ActiveSheet.Rows("9:9").RowHeight = ActiveSheet.Shapes.Range("Select_Theme_1").Height + 6
Sheets("Theme Presets").Visible = False
End Sub

Sub Presets_Button()
    Application.ScreenUpdating = False
    Sheets("Control").Select
    With ActiveSheet.Shapes.Range(Array("Presets_Button"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    Sheets("Theme Presets").Visible = True

    Sheets("Control").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Presets_Button"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    Sheets("Theme Presets").Select
End Sub


