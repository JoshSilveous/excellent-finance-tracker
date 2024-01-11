Attribute VB_Name = "EditCategory"
Dim ChangesMade As Boolean
Sub Edit_Category_Button()
'
' DO NOT TOUCH unless you know what you're doing.
'
    On Error Resume Next
        Set SelectedCell = Selection.Address
        If Err.Number <> 0 Then SelectedCell = "A1"
    On Error GoTo 0

    Application.ScreenUpdating = False
    Sheets("Control").Select
    With ActiveSheet.Shapes.Range(Array("Edit_Category_Button"))
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
    CategoriesPage.Show
    If ChangesMade Then
        PeriodSheets.render
        ControlPage.renderCat
        OverviewPage.render
    End If

    Sheets("Control").Select
    Range(SelectedCell).Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Edit_Category_Button"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .ThreeD.BevelTopDepth = 0.5
        .IncrementTop -1.2
    End With

End Sub
Function removeCat(CatName As String)

    addOne = False
    For i = 1 To f.getCatCount() Step 1
        ThisRow = i + 4
        If Sheets("Control").Range("B" & ThisRow).Value = CatName Then
            Sheets("Control").Range("B" & ThisRow).Value = Sheets("Control").Range("B" & ThisRow + 1).Value
            addOne = True
        ElseIf addOne = True Then
            Sheets("Control").Range("B" & ThisRow).Value = Sheets("Control").Range("B" & ThisRow + 1).Value
        End If
    Next
    For Each Period In f.getPerArray
        Sheets(Period).Select
        Call PeriodSheets.renameCat(CatName, "")
    Next
    

End Function
Function renameCat(OldName As String, NewName As String)

    For i = 1 To f.getCatCount() Step 1
        ThisRange = "B" & i + 4
        If Sheets("Control").Range(ThisRange).Value = OldName Then
            Sheets("Control").Range(ThisRange).Value = NewName
        End If
    Next
    
    For Each Period In f.getPerArray
        Sheets(Period).Select
        Call PeriodSheets.renameCat(OldName, NewName)
    Next

    
End Function
Function addCat(CatName As String)
    NewCell = "B" & f.getCatCount + 5
    Sheets("Control").Range(NewCell).Value = CatName
End Function
Function applyArray(ArrayInput() As String)
    ChangesMade = True
    Let RowIndex = 5
    For Each Item In ArrayInput
        Sheets("Control").Range("B" & RowIndex).Value = Item
        RowIndex = RowIndex + 1
    Next
    Sheets("Control").Range("B" & RowIndex & ":B50").ClearContents
End Function



