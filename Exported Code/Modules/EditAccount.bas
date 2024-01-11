Attribute VB_Name = "EditAccount"
Dim NewActBalanceArray() As String
Dim ChangesMade As Boolean
Sub Edit_Account_Button()
'
' DO NOT TOUCH unless you know what you're doing.
'
    On Error Resume Next
        Set SelectedCell = Selection.Address
        If Err.Number <> 0 Then SelectedCell = "A1"
    On Error GoTo 0

    Application.ScreenUpdating = True
    Sheets("Control").Select
    With ActiveSheet.Shapes.Range(Array("Edit_Account_Button"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ChangesMade = False
    AccountsPage.Show
    
    If ChangesMade Then
        PeriodSheets.render
        ControlPage.renderAct
        OverviewPage.render
        Call applyBalanceArray
    End If
    

    Sheets("Control").Select
    Range(SelectedCell).Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Edit_Account_Button"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .ThreeD.BevelTopDepth = 0.5
        .IncrementTop -1.2
    End With

End Sub
Function removeAct(ActName As String)

    addOne = False
    For i = 1 To f.getActCount() Step 1
        ThisRow = i + 4
        If Sheets("Control").Range("D" & ThisRow).Value = ActName Then
            Sheets("Control").Range("D" & ThisRow).Value = Sheets("Control").Range("D" & ThisRow + 1).Value
            addOne = True
        ElseIf addOne = True Then
            Sheets("Control").Range("D" & ThisRow).Value = Sheets("Control").Range("D" & ThisRow + 1).Value
        End If
    Next
    
    For Each Period In f.getPerArray
        Sheets(Period).Select
        Call PeriodSheets.renameAct(ActName, "")
    Next
    
End Function
Function renameAct(OldName As String, NewName As String)

    For i = 1 To f.getActCount() Step 1
        ThisRange = "D" & i + 4
        If Sheets("Control").Range(ThisRange).Value = OldName Then
            Sheets("Control").Range(ThisRange).Value = NewName
        End If
    Next
    
    For Each Period In f.getPerArray
        Sheets(Period).Select
        Call PeriodSheets.renameAct(OldName, NewName)
    Next

    
End Function
Function addAct(ActName As String)
    NewCell = "D" & f.getActCount + 5
    Sheets("Control").Range(NewCell).Value = ActName
End Function
Function applyArray(ArrayInput() As String)
    ChangesMade = True
    Let RowIndex = 5
    For Each Item In ArrayInput
        Sheets("Control").Range("D" & RowIndex).Value = Item
        RowIndex = RowIndex + 1
    Next
    Sheets("Control").Range("D" & RowIndex & ":D50").ClearContents
End Function
Function applyBalanceArray()
    
    On Error Resume Next
        FirstSheet = f.getPerArray()(0)
        RowIndex = 4
        For Each Item In NewActBalanceArray
            With Sheets(FirstSheet).Range("I" & RowIndex)
                .Value = Item
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End With
            RowIndex = RowIndex + 1
        Next
    On Error GoTo 0
End Function
Function setBalanceArray(ArrayInput() As String)
    NewActBalanceArray = ArrayInput
End Function



