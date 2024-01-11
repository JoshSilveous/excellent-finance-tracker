VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CategoriesPage 
   Caption         =   "Categories"
   ClientHeight    =   6252
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4920
   OleObjectBlob   =   "CategoriesPage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CategoriesPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentActiveValues() As String
Dim StartingValues() As String
Dim CurrentAddNewLabelLocation As Integer
Dim CurrentlyHandlingUpdate As Boolean


Private Sub UserForm_Initialize()
    CurrentActiveValues = f.getCatArray()
    StartingValues = CurrentActiveValues
    CurrentlyHandlingUpdate = False
    ' Fill in values
    Dim ItemIndex As Integer
    ItemIndex = 1
    For Each Item In CurrentActiveValues
        Call cf.setCategoryValue(ItemIndex, CStr(Item))
        cf.unmaskCategory (ItemIndex)
        cf.makeCategoryVisible (ItemIndex)
        ItemIndex = ItemIndex + 1
    Next

    
    cf.makeCategoryVisible (ItemIndex)
    cf.maskCategory (ItemIndex)
    
    ' Set scroll height

    positionAddCaption (ItemIndex)
    
    For i = ItemIndex + 1 To 30 Step 1
        Call cf.makeCategoryInvisible(CInt(i))
    Next
    
    CategoriesFrame.ScrollHeight = cf.getCategoryBoxScrollHeight
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' Check if any changes are left to apply
    ' If so, render them
    
    
    
    
    If UBound(CurrentActiveValues) <> UBound(StartingValues) Then
        Call EditCategory.applyArray(CurrentActiveValues)
        Exit Sub
    End If
    
    CurIndex = 0
    For Each Item In CurrentActiveValues
        If Item <> StartingValues(CurIndex) Then
            Call EditCategory.applyArray(CurrentActiveValues)
            Exit Sub
        End If
        CurIndex = CurIndex + 1
    Next
    
    
End Sub
Private Sub Close_Button_Click()
    Unload Me
End Sub


Function positionAddCaption(CatNum As Integer)
    
    LabelTopValue = 24 + (24 * (CatNum - 1)) + 4
    AddCategoryLabel.Top = LabelTopValue
    
    CurrentAddNewLabelLocation = CatNum
    
End Function
Private Sub AddCategoryLabel_Click()
    If cf.validateInput(cf.getCategoryValue(CurrentAddNewLabelLocation - 1)) Then
        focusCategory (CurrentAddNewLabelLocation)
    Else
        handleCategoryUpdate (CurrentAddNewLabelLocation - 1)
        focusCategory (CurrentAddNewLabelLocation - 1)
    End If
End Sub



Function handleCategoryUpdate(CatNum As Integer)
    CurrentlyHandlingUpdate = True
    CategoriesFrame.ScrollHeight = cf.getCategoryBoxScrollHeight()
    
    InputStr = cf.getCategoryValue(CatNum)
    
    ' Remove trailing spaces
    If InputStr <> Trim(InputStr) Then
        Call cf.setCategoryValue(CatNum, Trim(InputStr))
        InputStr = Trim(InputStr)
    End If
    
    ' Check to make sure input isn't numerical
    If IsNumeric(InputStr) Then
        Call MsgBox("Category names cannot be only numbers.", vbOKOnly, "Invalid Input")
        Call deleteCategory(CatNum, True)
        CurrentlyHandlingUpdate = False
        Exit Function
        
    ' Check to make sure input isn't a duplicate
    ElseIf InputStr <> "" Then
        DuplicateFound = False
        For i = 1 To cf.getVisibleCategoryCount() Step 1
            If CatNum <> i And InputStr = cf.getCategoryValue(CInt(i)) Then
                DuplicateFound = True
            End If
        Next
        If DuplicateFound Then
            Call MsgBox("Duplicate category name." & vbNewLine & _
                    Chr(34) & InputStr & Chr(34) & " already exists.", _
                    vbOKOnly, "Invalid Input")
            
            ' If the user is renaming a category to a duplicate category name
            If UBound(CurrentActiveValues) + 1 >= CatNum Then
                Call cf.setCategoryValue(CatNum, CurrentActiveValues(CatNum - 1))
                
            Else
                Call deleteCategory(CatNum, True)
            End If
            CurrentlyHandlingUpdate = False
            Exit Function
        End If
    End If
    
    ' If input is already in CurrentActiveValues array (new inputs)
    If UBound(CurrentActiveValues) + 1 >= CatNum Then
        
        Dim OldName As String
        OldName = CurrentActiveValues(CatNum - 1)
        
        ' Check if this is an old value being set to ""
        ' Prompt to delete if so
        If InputStr = "" And OldName <> "" Then
            
            ' If this is the only category left, do not allow deletion
            If cf.getVisibleCategoryCount = 2 Then
                Call MsgBox("You cannot have less than 1 category.", vbExclamation, "Error")
                Call cf.setCategoryValue(CatNum, OldName)
        CurrentlyHandlingUpdate = False
                Exit Function
            End If
            
            Confirmation = MsgBox("Are you sure you would like to delete " & Chr(34) _
                & OldName & Chr(34) & "?", vbYesNo, "Delete")
            If Confirmation = vbYes Then
                Call deleteCategory(CatNum, True)
            Else
                Call cf.setCategoryValue(CatNum, OldName)
            End If
        CurrentlyHandlingUpdate = False
            Exit Function
        End If
        
        ' Run rename functions
        If OldName <> InputStr And OldName <> "" Then
            Confirmation = MsgBox("Are you sure you would like to rename " & Chr(34) & OldName _
                                & Chr(34) & " to " & Chr(34) & InputStr & Chr(34) & "?", _
                                vbYesNo, "Rename Category")
            If Confirmation = vbYes Then
                
                Call EditCategory.renameCat(OldName, CStr(InputStr))
                
                CurrentActiveValues(CatNum - 1) = InputStr

                
            Else
                Call cf.setCategoryValue(CatNum, OldName)
            End If
        
        End If
    ' Else, new category
    Else
        ' If new category name is blank, remove it
        If InputStr = "" Then
            
            Call focusCategory(CurrentAddNewLabelLocation)

            Call deleteCategory(CatNum, True)
            CurrentlyHandlingUpdate = False
            Exit Function
        End If
        
        Call EditCategory.addCat(CStr(InputStr))
        
        NewValues = cf.getCategoryValue(1)
        
        VisCats = cf.getVisibleCategoryCount
        Dim LoopEnd As Integer
        If VisCats = 30 And cf.getCategoryValue(30) <> "" Then
            LoopEnd = 30
        Else
            LoopEnd = cf.getVisibleCategoryCount - 1
        End If
        
        For i = 2 To LoopEnd Step 1
            NewValues = NewValues + "|!DELIM!|" & cf.getCategoryValue(CInt(i))
        Next
        CurrentActiveValues = Split(NewValues, "|!DELIM!|")
    End If
    
        CurrentlyHandlingUpdate = False
End Function

Function handleCategoryKeypress(CatNum As Integer)
    If CatNum = cf.getVisibleCategoryCount Then
        cf.maskCategory (CatNum)
        
        If CatNum <> 30 Then
            AddCategoryLabel.Visible = True
            positionAddCaption (CatNum + 1)
            cf.makeCategoryVisible (CatNum + 1)
            cf.maskCategory (CatNum + 1)
        Else
            AddCategoryLabel.Visible = False
        End If
        cf.unmaskCategory (CatNum)
    End If
    
    
End Function
Function shiftCategoryUp(CatNum As Integer)
    If CatNum = 1 Then Exit Function
    
    Dim ThisValue As String
    Dim AboveValue As String
    
    ThisValue = cf.getCategoryValue(CatNum)
    AboveValue = cf.getCategoryValue(CatNum - 1)
    
    Call cf.setCategoryValue(CatNum - 1, ThisValue)
    Call cf.setCategoryValue(CatNum, AboveValue)
    CurrentActiveValues(CatNum - 2) = ThisValue
    CurrentActiveValues(CatNum - 1) = AboveValue
    
End Function
Function shiftCategoryDown(CatNum As Integer)

    Dim LowestAcceptableCatNum As Integer
    
    VisCats = cf.getVisibleCategoryCount
    If VisCats = 30 And cf.getCategoryValue(30) <> "" Then
        LowestAcceptableCatNum = 29
    ElseIf VisCats = 30 Then
        LowestAcceptableCatNum = 28
    Else
        LowestAcceptableCatNum = VisCats - 2
    End If
    
    If CatNum > LowestAcceptableCatNum Then Exit Function
    
    Dim ThisValue As String
    Dim BelowValue As String
    
    ThisValue = cf.getCategoryValue(CatNum)
    BelowValue = cf.getCategoryValue(CatNum + 1)
    
    Call cf.setCategoryValue(CatNum, BelowValue)
    Call cf.setCategoryValue(CatNum + 1, ThisValue)
    
    CurrentActiveValues(CatNum - 1) = BelowValue
    CurrentActiveValues(CatNum) = ThisValue

    
End Function
Function deleteCategory(CatNum As Integer, Optional SkipPrompt = False)
    Dim CatCount As Integer
    CatCount = cf.getVisibleCategoryCount
    
    ' Break if user is trying to delete last category
    If CatCount = 2 Then
        Call MsgBox("You cannot have less than 1 category.", vbExclamation, "Error")
        Exit Function
    End If
    
    ' Prompt for confirmation
    If SkipPrompt = False Then
        Confirmation = MsgBox("Are you sure you would like to delete " & Chr(34) _
                & cf.getCategoryValue(CatNum) & Chr(34) & "?", vbYesNo, "Delete")
    Else
        Confirmation = vbYes
    End If


    If Confirmation = vbYes Then
        Call EditCategory.removeCat(cf.getCategoryValue(CatNum))
        
        ' Set the current value to ""
        Call cf.setCategoryValue(CatNum, "")
        
        ' If user is deleting the 30th category, just mask it and put the label on it
        If CatNum = 30 Then
            cf.maskCategory (30)
            AddCategoryLabel.Visible = True
            positionAddCaption (30)
            
        ' Otherwise, loop through every value below this and move them all up one
        Else
            
            If CatCount - 1 > CatNum Then
                BelowCategoryValue = cf.getCategoryValue(CatNum + 1)
                For i = CatNum To CatCount - 1 Step 1
                    Call cf.setCategoryValue(CInt(i), CStr(BelowCategoryValue))
                    BelowCategoryValue = cf.getCategoryValue(i + 2)
                Next
            End If
            
            AddCategoryLabel.Visible = True
            
            ' If every category was full, including the 30th, clear and mask the 30th category
            If CatCount = 30 And cf.getCategoryValue(CatCount) <> "" Then
                Call cf.setCategoryValue(30, "")
                cf.maskCategory (30)
                positionAddCaption (30)
                AddCategoryLabel.Visible = True
                
            ' Otherwise, hide the last category visible and mask the second-last
            Else
                If CatCount - 1 >= CatNum Then
                    positionAddCaption (CatCount - 1)
                    cf.maskCategory (CatCount - 1)
                    cf.makeCategoryInvisible (CatCount)
                End If
            End If
            
        End If
        
        ' Adjust the ScrollHeight
        CategoriesFrame.ScrollHeight = cf.getCategoryBoxScrollHeight()
        
        ' Update CurrentActiveValues array to match current values
        Dim NewActiveValuesStr As String
        NewActiveValuesStr = cf.getCategoryValue(1)
        For j = 2 To cf.getVisibleCategoryCount - 1 Step 1
            NewActiveValuesStr = NewActiveValuesStr & "|!DELIM!|" & cf.getCategoryValue(CInt(j))
        Next
        CurrentActiveValues = Split(NewActiveValuesStr, "|!DELIM!|")
        
        ' Scroll back up to top to prevent visual bugs (InputBoxes selectable even when made invisible).
        CategoriesFrame.ScrollTop = 0
    End If
    
    
    
End Function

Function focusCategory(CatNum As Integer)
    If CatNum = 1 Then
        Category1Input.SetFocus
        Exit Function
    ElseIf CatNum = 2 Then
        Category2Input.SetFocus
        Exit Function
    ElseIf CatNum = 3 Then
        Category3Input.SetFocus
        Exit Function
    ElseIf CatNum = 4 Then
        Category4Input.SetFocus
        Exit Function
    ElseIf CatNum = 5 Then
        Category5Input.SetFocus
        Exit Function
    ElseIf CatNum = 6 Then
        Category6Input.SetFocus
        Exit Function
    ElseIf CatNum = 7 Then
        Category7Input.SetFocus
        Exit Function
    ElseIf CatNum = 8 Then
        Category8Input.SetFocus
        Exit Function
    ElseIf CatNum = 9 Then
        Category9Input.SetFocus
        Exit Function
    ElseIf CatNum = 10 Then
        Category10Input.SetFocus
        Exit Function
    ElseIf CatNum = 11 Then
        Category11Input.SetFocus
        Exit Function
    ElseIf CatNum = 12 Then
        Category12Input.SetFocus
        Exit Function
    ElseIf CatNum = 13 Then
        Category13Input.SetFocus
        Exit Function
    ElseIf CatNum = 14 Then
        Category14Input.SetFocus
        Exit Function
    ElseIf CatNum = 15 Then
        Category15Input.SetFocus
        Exit Function
    ElseIf CatNum = 16 Then
        Category16Input.SetFocus
        Exit Function
    ElseIf CatNum = 17 Then
        Category17Input.SetFocus
        Exit Function
    ElseIf CatNum = 18 Then
        Category18Input.SetFocus
        Exit Function
    ElseIf CatNum = 19 Then
        Category19Input.SetFocus
        Exit Function
    ElseIf CatNum = 20 Then
        Category20Input.SetFocus
        Exit Function
    ElseIf CatNum = 21 Then
        Category21Input.SetFocus
        Exit Function
    ElseIf CatNum = 22 Then
        Category22Input.SetFocus
        Exit Function
    ElseIf CatNum = 23 Then
        Category23Input.SetFocus
        Exit Function
    ElseIf CatNum = 24 Then
        Category24Input.SetFocus
        Exit Function
    ElseIf CatNum = 25 Then
        Category25Input.SetFocus
        Exit Function
    ElseIf CatNum = 26 Then
        Category26Input.SetFocus
        Exit Function
    ElseIf CatNum = 27 Then
        Category27Input.SetFocus
        Exit Function
    ElseIf CatNum = 28 Then
        Category28Input.SetFocus
        Exit Function
    ElseIf CatNum = 29 Then
        Category29Input.SetFocus
        Exit Function
    ElseIf CatNum = 30 Then
        Category30Input.SetFocus
        Exit Function
    End If
End Function

Private Sub Category1Delete_Click()
    deleteCategory (1)
End Sub
Private Sub Category2Delete_Click()
    deleteCategory (2)
End Sub
Private Sub Category3Delete_Click()
    deleteCategory (3)
End Sub
Private Sub Category4Delete_Click()
    deleteCategory (4)
End Sub
Private Sub Category5Delete_Click()
    deleteCategory (5)
End Sub
Private Sub Category6Delete_Click()
    deleteCategory (6)
End Sub
Private Sub Category7Delete_Click()
    deleteCategory (7)
End Sub
Private Sub Category8Delete_Click()
    deleteCategory (8)
End Sub
Private Sub Category9Delete_Click()
    deleteCategory (9)
End Sub
Private Sub Category10Delete_Click()
    deleteCategory (10)
End Sub
Private Sub Category11Delete_Click()
    deleteCategory (11)
End Sub
Private Sub Category12Delete_Click()
    deleteCategory (12)
End Sub
Private Sub Category13Delete_Click()
    deleteCategory (13)
End Sub
Private Sub Category14Delete_Click()
    deleteCategory (14)
End Sub
Private Sub Category15Delete_Click()
    deleteCategory (15)
End Sub
Private Sub Category16Delete_Click()
    deleteCategory (16)
End Sub
Private Sub Category17Delete_Click()
    deleteCategory (17)
End Sub
Private Sub Category18Delete_Click()
    deleteCategory (18)
End Sub
Private Sub Category19Delete_Click()
    deleteCategory (19)
End Sub
Private Sub Category20Delete_Click()
    deleteCategory (20)
End Sub
Private Sub Category21Delete_Click()
    deleteCategory (21)
End Sub
Private Sub Category22Delete_Click()
    deleteCategory (22)
End Sub
Private Sub Category23Delete_Click()
    deleteCategory (23)
End Sub
Private Sub Category24Delete_Click()
    deleteCategory (24)
End Sub
Private Sub Category25Delete_Click()
    deleteCategory (25)
End Sub
Private Sub Category26Delete_Click()
    deleteCategory (26)
End Sub
Private Sub Category27Delete_Click()
    deleteCategory (27)
End Sub
Private Sub Category28Delete_Click()
    deleteCategory (28)
End Sub
Private Sub Category29Delete_Click()
    deleteCategory (29)
End Sub
Private Sub Category30Delete_Click()
    deleteCategory (30)
End Sub
Private Sub Category1Spin_SpinUp()
    shiftCategoryUp (1)
End Sub
Private Sub Category2Spin_SpinUp()
    shiftCategoryUp (2)
End Sub
Private Sub Category3Spin_SpinUp()
    shiftCategoryUp (3)
End Sub
Private Sub Category4Spin_SpinUp()
    shiftCategoryUp (4)
End Sub
Private Sub Category5Spin_SpinUp()
    shiftCategoryUp (5)
End Sub
Private Sub Category6Spin_SpinUp()
    shiftCategoryUp (6)
End Sub
Private Sub Category7Spin_SpinUp()
    shiftCategoryUp (7)
End Sub
Private Sub Category8Spin_SpinUp()
    shiftCategoryUp (8)
End Sub
Private Sub Category9Spin_SpinUp()
    shiftCategoryUp (9)
End Sub
Private Sub Category10Spin_SpinUp()
    shiftCategoryUp (10)
End Sub
Private Sub Category11Spin_SpinUp()
    shiftCategoryUp (11)
End Sub
Private Sub Category12Spin_SpinUp()
    shiftCategoryUp (12)
End Sub
Private Sub Category13Spin_SpinUp()
    shiftCategoryUp (13)
End Sub
Private Sub Category14Spin_SpinUp()
    shiftCategoryUp (14)
End Sub
Private Sub Category15Spin_SpinUp()
    shiftCategoryUp (15)
End Sub
Private Sub Category16Spin_SpinUp()
    shiftCategoryUp (16)
End Sub
Private Sub Category17Spin_SpinUp()
    shiftCategoryUp (17)
End Sub
Private Sub Category18Spin_SpinUp()
    shiftCategoryUp (18)
End Sub
Private Sub Category19Spin_SpinUp()
    shiftCategoryUp (19)
End Sub
Private Sub Category20Spin_SpinUp()
    shiftCategoryUp (20)
End Sub
Private Sub Category21Spin_SpinUp()
    shiftCategoryUp (21)
End Sub
Private Sub Category22Spin_SpinUp()
    shiftCategoryUp (22)
End Sub
Private Sub Category23Spin_SpinUp()
    shiftCategoryUp (23)
End Sub
Private Sub Category24Spin_SpinUp()
    shiftCategoryUp (24)
End Sub
Private Sub Category25Spin_SpinUp()
    shiftCategoryUp (25)
End Sub
Private Sub Category26Spin_SpinUp()
    shiftCategoryUp (26)
End Sub
Private Sub Category27Spin_SpinUp()
    shiftCategoryUp (27)
End Sub
Private Sub Category28Spin_SpinUp()
    shiftCategoryUp (28)
End Sub
Private Sub Category29Spin_SpinUp()
    shiftCategoryUp (29)
End Sub
Private Sub Category30Spin_SpinUp()
    shiftCategoryUp (30)
End Sub
Private Sub Category1Spin_SpinDown()
    shiftCategoryDown (1)
End Sub
Private Sub Category2Spin_SpinDown()
    shiftCategoryDown (2)
End Sub
Private Sub Category3Spin_SpinDown()
    shiftCategoryDown (3)
End Sub
Private Sub Category4Spin_SpinDown()
    shiftCategoryDown (4)
End Sub
Private Sub Category5Spin_SpinDown()
    shiftCategoryDown (5)
End Sub
Private Sub Category6Spin_SpinDown()
    shiftCategoryDown (6)
End Sub
Private Sub Category7Spin_SpinDown()
    shiftCategoryDown (7)
End Sub
Private Sub Category8Spin_SpinDown()
    shiftCategoryDown (8)
End Sub
Private Sub Category9Spin_SpinDown()
    shiftCategoryDown (9)
End Sub
Private Sub Category10Spin_SpinDown()
    shiftCategoryDown (10)
End Sub
Private Sub Category11Spin_SpinDown()
    shiftCategoryDown (11)
End Sub
Private Sub Category12Spin_SpinDown()
    shiftCategoryDown (12)
End Sub
Private Sub Category13Spin_SpinDown()
    shiftCategoryDown (13)
End Sub
Private Sub Category14Spin_SpinDown()
    shiftCategoryDown (14)
End Sub
Private Sub Category15Spin_SpinDown()
    shiftCategoryDown (15)
End Sub
Private Sub Category16Spin_SpinDown()
    shiftCategoryDown (16)
End Sub
Private Sub Category17Spin_SpinDown()
    shiftCategoryDown (17)
End Sub
Private Sub Category18Spin_SpinDown()
    shiftCategoryDown (18)
End Sub
Private Sub Category19Spin_SpinDown()
    shiftCategoryDown (19)
End Sub
Private Sub Category20Spin_SpinDown()
    shiftCategoryDown (20)
End Sub
Private Sub Category21Spin_SpinDown()
    shiftCategoryDown (21)
End Sub
Private Sub Category22Spin_SpinDown()
    shiftCategoryDown (22)
End Sub
Private Sub Category23Spin_SpinDown()
    shiftCategoryDown (23)
End Sub
Private Sub Category24Spin_SpinDown()
    shiftCategoryDown (24)
End Sub
Private Sub Category25Spin_SpinDown()
    shiftCategoryDown (25)
End Sub
Private Sub Category26Spin_SpinDown()
    shiftCategoryDown (26)
End Sub
Private Sub Category27Spin_SpinDown()
    shiftCategoryDown (27)
End Sub
Private Sub Category28Spin_SpinDown()
    shiftCategoryDown (28)
End Sub
Private Sub Category29Spin_SpinDown()
    shiftCategoryDown (29)
End Sub
Private Sub Category30Spin_SpinDown()
    shiftCategoryDown (30)
End Sub
Private Sub Category1Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (1)
End Sub
Private Sub Category2Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (2)
End Sub
Private Sub Category3Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (3)
End Sub
Private Sub Category4Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (4)
End Sub
Private Sub Category5Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (5)
End Sub
Private Sub Category6Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (6)
End Sub
Private Sub Category7Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (7)
End Sub
Private Sub Category8Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (8)
End Sub
Private Sub Category9Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (9)
End Sub
Private Sub Category10Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (10)
End Sub
Private Sub Category11Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (11)
End Sub
Private Sub Category12Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (12)
End Sub
Private Sub Category13Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (13)
End Sub
Private Sub Category14Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (14)
End Sub
Private Sub Category15Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (15)
End Sub
Private Sub Category16Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (16)
End Sub
Private Sub Category17Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (17)
End Sub
Private Sub Category18Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (18)
End Sub
Private Sub Category19Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (19)
End Sub
Private Sub Category20Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (20)
End Sub
Private Sub Category21Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (21)
End Sub
Private Sub Category22Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (22)
End Sub
Private Sub Category23Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (23)
End Sub
Private Sub Category24Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (24)
End Sub
Private Sub Category25Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (25)
End Sub
Private Sub Category26Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (26)
End Sub
Private Sub Category27Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (27)
End Sub
Private Sub Category28Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (28)
End Sub
Private Sub Category29Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (29)
End Sub
Private Sub Category30Input_Change()
    If Not CurrentlyHandlingUpdate Then handleCategoryKeypress (30)
End Sub
Private Sub Category1Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (1)
End Sub
Private Sub Category2Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (2)
End Sub
Private Sub Category3Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (3)
End Sub
Private Sub Category4Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (4)
End Sub
Private Sub Category5Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (5)
End Sub
Private Sub Category6Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (6)
End Sub
Private Sub Category7Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (7)
End Sub
Private Sub Category8Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (8)
End Sub
Private Sub Category9Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (9)
End Sub
Private Sub Category10Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (10)
End Sub
Private Sub Category11Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (11)
End Sub
Private Sub Category12Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (12)
End Sub
Private Sub Category13Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (13)
End Sub
Private Sub Category14Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (14)
End Sub
Private Sub Category15Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (15)
End Sub
Private Sub Category16Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (16)
End Sub
Private Sub Category17Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (17)
End Sub
Private Sub Category18Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (18)
End Sub
Private Sub Category19Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (19)
End Sub
Private Sub Category20Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (20)
End Sub
Private Sub Category21Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (21)
End Sub
Private Sub Category22Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (22)
End Sub
Private Sub Category23Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (23)
End Sub
Private Sub Category24Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (24)
End Sub
Private Sub Category25Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (25)
End Sub
Private Sub Category26Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (26)
End Sub
Private Sub Category27Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (27)
End Sub
Private Sub Category28Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (28)
End Sub
Private Sub Category29Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (29)
End Sub
Private Sub Category30Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleCategoryUpdate (30)
End Sub





