Attribute VB_Name = "Begin"
Sub Welcome_Begin_Button()
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = True
    Sheets("Welcome").Select
    With ActiveSheet.Shapes.Range(Array("Welcome_Begin_Button"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    Call welcomeBegin

    Application.ScreenUpdating = True
    
    On Error Resume Next ' If welcomeBegin goes through successfully, this button will be deleted
        With ActiveSheet.Shapes.Range(Array("Welcome_Begin_Button"))
            With .Shadow
                .OffsetX = 1.2246467991E-16
                .OffsetY = 2
            End With
            .ThreeD.BevelTopInset = 1
            .IncrementTop -1.2
            .ThreeD.BevelTopDepth = 0.5
        End With
    On Error GoTo 0
    
End Sub
Sub welcomeBegin()

    BeginForm.Show
    StartInterval = BeginForm.ReturnStartInterval
    
    Categories = Split(BeginForm.ReturnCategoriesStr, "|!DELIM!|")
    
    Accounts = Split(BeginForm.ReturnAccountsStr, "|!DELIM!|")
    
    StartingBalances = Split(BeginForm.ReturnStartingBalancesStr, "|!DELIM!|")
    
    PeriodType = BeginForm.ReturnPeriodTypeStr
    
    
    If StartInterval = "" Then Exit Sub
    
    
    
    ' Unhide / Update Controls Page
    Sheets("Control").Visible = True
    
    ' Run Controls Page Render Function
    Sheets("Control").Range("B5:B2000, D5:D2000").ClearContents
    
    Index = 0
    For Each Category In Categories
        Sheets("Control").Range("B" & Index + 5).Value = Category
        Index = Index + 1
    Next
    Index = 0
    For Each Account In Accounts
        Sheets("Control").Range("D" & Index + 5).Value = Account
        Index = Index + 1
    Next
    ControlPage.renderAct
    ControlPage.renderCat
    
    
    ' Update "add" buttons
    ButtonStr = ""
    If PeriodType = "Monthly" Then ButtonStr = "Month"
    If PeriodType = "Weekly" Then ButtonStr = "Week"
    If PeriodType = "Bi-Weekly" Then ButtonStr = "2 Weeks"
    Sheets("Overview").Shapes.Range(Array("Add_Period_Button")).TextFrame2.TextRange.Characters.Text _
    = "+     Add Next " & ButtonStr
    Sheets("Interval").Shapes.Range(Array("Add_Period_Button")).TextFrame2.TextRange.Characters.Text _
    = "+     Add Next " & ButtonStr
    
    ' Add Period Sheet and render
    Sheets("Overview").Visible = True
    NewSheetIndex = Sheets("Overview").Index
    
    Sheets("Interval").Visible = True
    Sheets("Interval").Select
    Sheets("Interval").Copy After:=Sheets(NewSheetIndex)
    Sheets("Interval (2)").Select
    Sheets("Interval (2)").Name = StartInterval
    Range("A1:P1").Select
    ActiveCell.FormulaR1C1 = "'" & StartInterval
    Range("A2").Select
    Sheets("Interval").Visible = False
    
    ' Update overview page
    Sheets("Overview").Range("C2").FormulaR1C1 = "'" & StartInterval
    PeriodSheets.render
    
    ' Change first values on period sheet to account starting balances
    Index = 0
    For Each Item In StartingBalances
        With Sheets(StartInterval).Range("I" & Index + 4)
            .FormulaR1C1 = Item
            .Style = "Currency"
        End With
        Index = Index + 1
    Next
    ' Update overview page
    OverviewPage.render
    applyTheme.applyThemePeriods
    
    Sheets(StartInterval).Select
    Sheets("Welcome").Shapes.Range("Welcome_Begin_Button").Delete
    Sheets("Welcome").Range("A13").ClearContents
    Sheets("Welcome").Visible = False

End Sub


