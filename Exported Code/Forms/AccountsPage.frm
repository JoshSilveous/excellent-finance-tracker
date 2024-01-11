VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccountsPage 
   Caption         =   "Accounts"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   195
   ClientWidth     =   6495
   OleObjectBlob   =   "AccountsPage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AccountsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentActiveValues() As String
Dim StartingValues() As String
Dim CurrentStartBalValues() As String
Dim StartingStartBalValues() As String
Dim CurrentAddNewLabelLoAction As Integer
Dim CurrentlyHandlingUpdate As Boolean






Private Sub UserForm_Initialize()
    CurrentActiveValues = f.getActArray()
    StartingValues = CurrentActiveValues
    CurrentStartBalValues = f.getActStartBalArray() ' ****
    StartingStartBalValues = CurrentStartBalValues
    CurrentlyHandlingUpdate = False
    ' Fill in values
    Dim ItemIndex As Integer
    ItemIndex = 1
    For Each Item In CurrentActiveValues
        af.unmaskAccount (ItemIndex)
        af.makeAccountVisible (ItemIndex)
        Call af.setAccountValue(ItemIndex, CStr(Item))
        
        BalVal = Format(CurrentStartBalValues(ItemIndex - 1), "$##,#0.00")
        If BalVal <> "" Then Call af.setStartBalValue(ItemIndex, CStr(BalVal))
        
        ItemIndex = ItemIndex + 1
    Next

    
    af.makeAccountVisible (ItemIndex)
    af.maskAccount (ItemIndex)
    
    ' Set scroll height

    positionAddCaption (ItemIndex)
    
    For i = ItemIndex + 1 To 30 Step 1
        Call af.makeAccountInvisible(CInt(i))
    Next
    
AccountsFrame.ScrollHeight = af.getAccountBoxScrollHeight
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If UBound(CurrentActiveValues) <> UBound(StartingValues) Then
        Call EditAccount.applyArray(CurrentActiveValues)
        Call EditAccount.setBalanceArray(CurrentStartBalValues)
        Exit Sub
    End If
    
    CurIndex = 0
    For Each Item In CurrentActiveValues
        If Item <> StartingValues(CurIndex) Then
            Call EditAccount.applyArray(CurrentActiveValues)
            Call EditAccount.setBalanceArray(CurrentStartBalValues)
            Exit Sub
        End If
        CurIndex = CurIndex + 1
    Next
    
        CurIndex = 0
    For Each Item In CurrentStartBalValues
        If Item <> StartingStartBalValues(CurIndex) Then
            Call EditAccount.applyArray(CurrentActiveValues)
            Call EditAccount.setBalanceArray(CurrentStartBalValues)
            Exit Sub
        End If
        CurIndex = CurIndex + 1
    Next
End Sub
Private Sub Close_Button_Click()
    Unload Me
End Sub


Function positionAddCaption(ActNum As Integer)
    
    LabelTopValue = 24 + (24 * (ActNum - 1)) + 4
    AddAccountLabel.Top = LabelTopValue
    
    CurrentAddNewLabelLoAction = ActNum
    
End Function
Private Sub AddAccountLabel_Click()
    If af.validateInput(af.getAccountValue(CurrentAddNewLabelLoAction - 1)) Then
        focusAccount (CurrentAddNewLabelLoAction)
    Else
        handleAccountUpdate (CurrentAddNewLabelLoAction - 1)
        focusAccount (CurrentAddNewLabelLoAction - 1)
    End If
End Sub



Function handleAccountUpdate(ActNum As Integer)
    CurrentlyHandlingUpdate = True
   AccountsFrame.ScrollHeight = af.getAccountBoxScrollHeight()
    
    InputStr = af.getAccountValue(ActNum)
    
    ' Remove trailing spaces
    If InputStr <> Trim(InputStr) Then
        Call af.setAccountValue(ActNum, Trim(InputStr))
        InputStr = Trim(InputStr)
    End If
    
    ' Check to make sure input isn't numerical
    If IsNumeric(InputStr) Then
        Call MsgBox("Account names cannot be only numbers.", vbOKOnly, "Invalid Input")
        Call deleteAccount(ActNum, True)
        CurrentlyHandlingUpdate = False
        Exit Function
        
    ' Check to make sure input isn't a Duplicate
    ElseIf InputStr <> "" Then
        DuplicateFound = False
        For i = 1 To af.getVisibleAccountCount() Step 1
            If ActNum <> i And InputStr = af.getAccountValue(CInt(i)) Then
                DuplicateFound = True
            End If
        Next
        If DuplicateFound Then
            Call MsgBox("DuplicateAccount name." & vbNewLine & _
                    Chr(34) & InputStr & Chr(34) & " already exists.", _
                    vbOKOnly, "Invalid Input")
            If UBound(CurrentActiveValues) + 1 >= ActNum Then
                Call af.setAccountValue(ActNum, CurrentActiveValues(ActNum - 1))
            Else
                Call deleteAccount(ActNum, True)
            End If
            CurrentlyHandlingUpdate = False
            Exit Function
        End If
    End If
    
    ' If input is already in CurrentActiveValues array (renaming)
    If UBound(CurrentActiveValues) + 1 >= ActNum Then
        
        Dim OldName As String
        OldName = CurrentActiveValues(ActNum - 1)
        
        ' Check if this is an old value being set to ""
        ' Prompt to delete if so
        If InputStr = "" And OldName <> "" Then
            
            ' If this is the onlyAccount left, do not allow deletion
            If af.getVisibleAccountCount = 2 Then
                Call MsgBox("You cannot have less than 1Account.", vbExclamation, "Error")
                Call af.setAccountValue(ActNum, OldName)
                CurrentlyHandlingUpdate = False
                Exit Function
            End If
            
            Confirmation = MsgBox("Are you sure you would like to delete " & Chr(34) _
                & OldName & Chr(34) & "?", vbYesNo, "Delete")
            If Confirmation = vbYes Then
                Call deleteAccount(ActNum, True)
            Else
                Call af.setAccountValue(ActNum, OldName)
            End If
        CurrentlyHandlingUpdate = False
            Exit Function
        End If
        
        ' Run rename functions
        If OldName <> InputStr And OldName <> "" Then
            Confirmation = MsgBox("Are you sure you would like to rename " & Chr(34) & OldName _
                                & Chr(34) & " to " & Chr(34) & InputStr & Chr(34) & "?", _
                                vbYesNo, "RenameAccount")
            If Confirmation = vbYes Then
                
                Call EditAccount.renameAct(OldName, CStr(InputStr))
                
                CurrentActiveValues(ActNum - 1) = InputStr
                
            Else
                Call af.setAccountValue(ActNum, OldName)
            End If
        
        End If
    ' Else, newAccount
    Else
        ' If newAccount name is blank, remove it
        If InputStr = "" Then
            
            Call focusAccount(CurrentAddNewLabelLoAction)

            Call deleteAccount(ActNum, True)
            CurrentlyHandlingUpdate = False
            Exit Function
        End If
        
        Call EditAccount.addAct(CStr(InputStr))
        
        NewValues = af.getAccountValue(1)
        NewStartBalValues = af.getStartBalValue(1)
        
        VisActs = af.getVisibleAccountCount
        Dim LoopEnd As Integer
        If VisActs = 30 And af.getAccountValue(30) <> "" Then
            LoopEnd = 30
        Else
            LoopEnd = af.getVisibleAccountCount - 1
        End If
        
        For i = 2 To LoopEnd Step 1
            NewValues = NewValues & "|!DELIM!|" & af.getAccountValue(CInt(i))
            NewStartBalValues = NewStartBalValues & "|!DELIM!|" & af.getStartBalValue(CInt(i))
        Next
        CurrentActiveValues = Split(NewValues, "|!DELIM!|")
        CurrentStartBalValues = Split(NewStartBalValues, "|!DELIM!|")
    End If
    
        CurrentlyHandlingUpdate = False
End Function
Function handleAccountKeypress(ActNum As Integer)
    If ActNum = af.getVisibleAccountCount Then
        af.maskAccount (ActNum)
        
        If ActNum <> 30 Then
            AddAccountLabel.Visible = True
            positionAddCaption (ActNum + 1)
            af.makeAccountVisible (ActNum + 1)
            af.maskAccount (ActNum + 1)
        Else
            AddAccountLabel.Visible = False
        End If
        af.unmaskAccount (ActNum)
    End If
    
    
End Function
Function handleStartBalUpdate(ActNum As Integer)
    InputStr = af.getStartBalValue(ActNum) ' ***
    InputStr = Trim(InputStr)
    If InputStr = "" Then
        Call af.setStartBalValue(ActNum, "$0.00")
        Call updateStartBalArray(ActNum - 1, "0")
    ElseIf Not IsNumeric(InputStr) Then
        Call MsgBox("Please enter a numeric input.", vbOKOnly, "Invalid Input")
        Call af.setStartBalValue(ActNum, "$0.00")
        Call updateStartBalArray(ActNum - 1, "0")
    Else
        Call af.setStartBalValue(ActNum, Format(InputStr, "$#,##0.00"))
        Call updateStartBalArray(ActNum - 1, CStr(InputStr))
    End If
End Function
Function updateStartBalArray(ActIndex As Integer, InputStr As String)
    ' if new value
    If ActIndex > UBound(CurrentStartBalValues) Then
        Dim StringValues As String
        ItemIndex = 0
        For Each Item In CurrentStartBalValues
            If ItemIndex = 0 Then
                StringValues = Item
            Else
                StringValues = StringValues & "|!DELIM!|" & Item
            End If
        Next
        StringValues = StringValues & "|!DELIM!|" & InputStr
        CurrentStartBalValues = Split(StringValues, "|!DELIM!|")
    ' if updating previous value
    Else
        CurrentStartBalValues(ActIndex) = InputStr
    End If
End Function
Function shiftAccountUp(ActNum As Integer)
    If ActNum = 1 Then Exit Function
    
    Dim ThisValue As String
    Dim ThisBalValue As String
    Dim AboveValue As String
    Dim AboveBalValue As String
    
    ThisValue = af.getAccountValue(ActNum)
    ThisBalValue = af.getStartBalValue(ActNum)
    AboveValue = af.getAccountValue(ActNum - 1)
    AboveBalValue = af.getStartBalValue(ActNum - 1)
    
    Call af.setAccountValue(ActNum - 1, ThisValue)
    Call af.setStartBalValue(ActNum - 1, ThisBalValue)
    Call af.setAccountValue(ActNum, AboveValue)
    Call af.setStartBalValue(ActNum, AboveBalValue)
    
    CurrentActiveValues(ActNum - 2) = ThisValue
    Call updateStartBalArray(ActNum - 2, ThisBalValue)
    CurrentActiveValues(ActNum - 1) = AboveValue
    Call updateStartBalArray(ActNum - 1, AboveBalValue)
End Function
Function shiftAccountDown(ActNum As Integer)

    Dim LowestAcceptableActNum As Integer
    
    VisActs = af.getVisibleAccountCount
    If VisActs = 30 And af.getAccountValue(30) <> "" Then
        LowestAcceptableActNum = 29
    ElseIf VisActs = 30 Then
        LowestAcceptableActNum = 28
    Else
        LowestAcceptableActNum = VisActs - 2
    End If
    
    If ActNum > LowestAcceptableActNum Then Exit Function
    
    Dim ThisValue As String
    Dim ThisBalValue As String
    Dim BelowValue As String
    Dim BelowBalValue As String
    
    ThisValue = af.getAccountValue(ActNum)
    ThisBalValue = af.getStartBalValue(ActNum)
    BelowValue = af.getAccountValue(ActNum + 1)
    BelowBalValue = af.getStartBalValue(ActNum + 1)
    
    Call af.setAccountValue(ActNum, BelowValue)
    Call af.setStartBalValue(ActNum, BelowBalValue)
    Call af.setAccountValue(ActNum + 1, ThisValue)
    Call af.setStartBalValue(ActNum + 1, ThisBalValue)
    
    CurrentActiveValues(ActNum - 1) = BelowValue
    Call updateStartBalArray(ActNum - 1, BelowBalValue)
    CurrentActiveValues(ActNum) = ThisValue
    Call updateStartBalArray(ActNum, ThisBalValue)

End Function
Function deleteAccount(ActNum As Integer, Optional SkipPrompt = False)
    Dim ActCount As Integer
    ActCount = af.getVisibleAccountCount
    
    ' Break if user is trying to delete lastAccount
    If ActCount = 2 Then
        Call MsgBox("You cannot have less than 1Account.", vbExclamation, "Error")
        Exit Function
    End If
    
    ' Prompt for confirmation
    If SkipPrompt = False Then
        Confirmation = MsgBox("Are you sure you would like to delete " & Chr(34) _
                & af.getAccountValue(ActNum) & Chr(34) & "?", vbYesNo, "Delete")
    Else
        Confirmation = vbYes
    End If


    If Confirmation = vbYes Then
        Call EditAccount.removeAct(af.getAccountValue(ActNum))
        
        ' Set the current value to ""
        Call af.setAccountValue(ActNum, "")
        Call af.setStartBalValue(ActNum, "$0.00")
        
        ' If user is deleting the 30thAccount, just mask it and put the label on it
        If ActNum = 30 Then
            af.maskAccount (30)
            AddAccountLabel.Visible = True
            positionAddCaption (30)
            
        ' Otherwise, loop through every value below this and move them all up one
        Else
            
            If ActCount - 1 > ActNum Then
                BelowAccountValue = af.getAccountValue(ActNum + 1)
                BelowStartBalValue = af.getStartBalValue(ActNum + 1)
                For i = ActNum To ActCount - 1 Step 1
                    Call af.setAccountValue(CInt(i), CStr(BelowAccountValue))
                    Call af.setStartBalValue(CInt(i), CStr(BelowStartBalValue))
                    BelowAccountValue = af.getAccountValue(i + 2)
                    BelowStartBalValue = af.getStartBalValue(i + 2)
                Next
            End If
            
            AddAccountLabel.Visible = True
            
            ' If everyAccount was full, including the 30th, clear and mask the 30thAccount
            If ActCount = 30 And af.getAccountValue(ActCount) <> "" Then
                Call af.setAccountValue(30, "")
                Call af.setStartBalValue(30, "$0.00")
                af.maskAccount (30)
                positionAddCaption (30)
                AddAccountLabel.Visible = True
                
            ' Otherwise, hide the lastAccount visible and mask the second-last
            Else
                If ActCount - 1 >= ActNum Then
                    positionAddCaption (ActCount - 1)
                    af.maskAccount (ActCount - 1)
                    af.makeAccountInvisible (ActCount)
                End If
            End If
            
        End If
        
        ' Adjust the ScrollHeight
       AccountsFrame.ScrollHeight = af.getAccountBoxScrollHeight()
        
        ' Update CurrentActiveValues array to match current values
        Dim NewActiveValuesStr As String
        Dim NewStartBalValuesStr As String
        
        NewActiveValuesStr = af.getAccountValue(1)
        NewStartBalValuesStr = af.getStartBalValue(1)
        
        For j = 2 To af.getVisibleAccountCount - 1 Step 1
            NewActiveValuesStr = NewActiveValuesStr & "|!DELIM!|" & af.getAccountValue(CInt(j))
            NewStartBalValuesStr = NewStartBalValuesStr & "|!DELIM!|" & af.getStartBalValue(CInt(j))
        Next
        CurrentActiveValues = Split(NewActiveValuesStr, "|!DELIM!|")
        CurrentStartBalValues = Split(NewStartBalValuesStr, "|!DELIM!|")
        
        ' Scroll back up to top to prevent visual bugs (InputBoxes selectable even when made invisible).
       AccountsFrame.ScrollTop = 0
    End If
    
    
    
End Function




Function focusAccount(ActNum As Integer)
    If ActNum = 1 Then
        Account1Input.SetFocus
        Exit Function
    ElseIf ActNum = 2 Then
        Account2Input.SetFocus
        Exit Function
    ElseIf ActNum = 3 Then
        Account3Input.SetFocus
        Exit Function
    ElseIf ActNum = 4 Then
        Account4Input.SetFocus
        Exit Function
    ElseIf ActNum = 5 Then
        Account5Input.SetFocus
        Exit Function
    ElseIf ActNum = 6 Then
        Account6Input.SetFocus
        Exit Function
    ElseIf ActNum = 7 Then
        Account7Input.SetFocus
        Exit Function
    ElseIf ActNum = 8 Then
        Account8Input.SetFocus
        Exit Function
    ElseIf ActNum = 9 Then
        Account9Input.SetFocus
        Exit Function
    ElseIf ActNum = 10 Then
        Account10Input.SetFocus
        Exit Function
    ElseIf ActNum = 11 Then
        Account11Input.SetFocus
        Exit Function
    ElseIf ActNum = 12 Then
        Account12Input.SetFocus
        Exit Function
    ElseIf ActNum = 13 Then
        Account13Input.SetFocus
        Exit Function
    ElseIf ActNum = 14 Then
        Account14Input.SetFocus
        Exit Function
    ElseIf ActNum = 15 Then
        Account15Input.SetFocus
        Exit Function
    ElseIf ActNum = 16 Then
        Account16Input.SetFocus
        Exit Function
    ElseIf ActNum = 17 Then
        Account17Input.SetFocus
        Exit Function
    ElseIf ActNum = 18 Then
        Account18Input.SetFocus
        Exit Function
    ElseIf ActNum = 19 Then
        Account19Input.SetFocus
        Exit Function
    ElseIf ActNum = 20 Then
        Account20Input.SetFocus
        Exit Function
    ElseIf ActNum = 21 Then
        Account21Input.SetFocus
        Exit Function
    ElseIf ActNum = 22 Then
        Account22Input.SetFocus
        Exit Function
    ElseIf ActNum = 23 Then
        Account23Input.SetFocus
        Exit Function
    ElseIf ActNum = 24 Then
        Account24Input.SetFocus
        Exit Function
    ElseIf ActNum = 25 Then
        Account25Input.SetFocus
        Exit Function
    ElseIf ActNum = 26 Then
        Account26Input.SetFocus
        Exit Function
    ElseIf ActNum = 27 Then
        Account27Input.SetFocus
        Exit Function
    ElseIf ActNum = 28 Then
        Account28Input.SetFocus
        Exit Function
    ElseIf ActNum = 29 Then
        Account29Input.SetFocus
        Exit Function
    ElseIf ActNum = 30 Then
        Account30Input.SetFocus
        Exit Function
    End If
End Function


Private Sub Account1Delete_Click()
    deleteAccount (1)
End Sub
Private Sub Account2Delete_Click()
    deleteAccount (2)
End Sub
Private Sub Account3Delete_Click()
    deleteAccount (3)
End Sub
Private Sub Account4Delete_Click()
    deleteAccount (4)
End Sub
Private Sub Account5Delete_Click()
    deleteAccount (5)
End Sub
Private Sub Account6Delete_Click()
    deleteAccount (6)
End Sub
Private Sub Account7Delete_Click()
    deleteAccount (7)
End Sub
Private Sub Account8Delete_Click()
    deleteAccount (8)
End Sub
Private Sub Account9Delete_Click()
    deleteAccount (9)
End Sub
Private Sub Account10Delete_Click()
    deleteAccount (10)
End Sub
Private Sub Account11Delete_Click()
    deleteAccount (11)
End Sub
Private Sub Account12Delete_Click()
    deleteAccount (12)
End Sub
Private Sub Account13Delete_Click()
    deleteAccount (13)
End Sub
Private Sub Account14Delete_Click()
    deleteAccount (14)
End Sub
Private Sub Account15Delete_Click()
    deleteAccount (15)
End Sub
Private Sub Account16Delete_Click()
    deleteAccount (16)
End Sub
Private Sub Account17Delete_Click()
    deleteAccount (17)
End Sub
Private Sub Account18Delete_Click()
    deleteAccount (18)
End Sub
Private Sub Account19Delete_Click()
    deleteAccount (19)
End Sub
Private Sub Account20Delete_Click()
    deleteAccount (20)
End Sub
Private Sub Account21Delete_Click()
    deleteAccount (21)
End Sub
Private Sub Account22Delete_Click()
    deleteAccount (22)
End Sub
Private Sub Account23Delete_Click()
    deleteAccount (23)
End Sub
Private Sub Account24Delete_Click()
    deleteAccount (24)
End Sub
Private Sub Account25Delete_Click()
    deleteAccount (25)
End Sub
Private Sub Account26Delete_Click()
    deleteAccount (26)
End Sub
Private Sub Account27Delete_Click()
    deleteAccount (27)
End Sub
Private Sub Account28Delete_Click()
    deleteAccount (28)
End Sub
Private Sub Account29Delete_Click()
    deleteAccount (29)
End Sub
Private Sub Account30Delete_Click()
    deleteAccount (30)
End Sub
Private Sub Account1Spin_SpinUp()
    shiftAccountUp (1)
End Sub
Private Sub Account2Spin_SpinUp()
    shiftAccountUp (2)
End Sub
Private Sub Account3Spin_SpinUp()
    shiftAccountUp (3)
End Sub
Private Sub Account4Spin_SpinUp()
    shiftAccountUp (4)
End Sub
Private Sub Account5Spin_SpinUp()
    shiftAccountUp (5)
End Sub
Private Sub Account6Spin_SpinUp()
    shiftAccountUp (6)
End Sub
Private Sub Account7Spin_SpinUp()
    shiftAccountUp (7)
End Sub
Private Sub Account8Spin_SpinUp()
    shiftAccountUp (8)
End Sub
Private Sub Account9Spin_SpinUp()
    shiftAccountUp (9)
End Sub
Private Sub Account10Spin_SpinUp()
    shiftAccountUp (10)
End Sub
Private Sub Account11Spin_SpinUp()
    shiftAccountUp (11)
End Sub
Private Sub Account12Spin_SpinUp()
    shiftAccountUp (12)
End Sub
Private Sub Account13Spin_SpinUp()
    shiftAccountUp (13)
End Sub
Private Sub Account14Spin_SpinUp()
    shiftAccountUp (14)
End Sub
Private Sub Account15Spin_SpinUp()
    shiftAccountUp (15)
End Sub
Private Sub Account16Spin_SpinUp()
    shiftAccountUp (16)
End Sub
Private Sub Account17Spin_SpinUp()
    shiftAccountUp (17)
End Sub
Private Sub Account18Spin_SpinUp()
    shiftAccountUp (18)
End Sub
Private Sub Account19Spin_SpinUp()
    shiftAccountUp (19)
End Sub
Private Sub Account20Spin_SpinUp()
    shiftAccountUp (20)
End Sub
Private Sub Account21Spin_SpinUp()
    shiftAccountUp (21)
End Sub
Private Sub Account22Spin_SpinUp()
    shiftAccountUp (22)
End Sub
Private Sub Account23Spin_SpinUp()
    shiftAccountUp (23)
End Sub
Private Sub Account24Spin_SpinUp()
    shiftAccountUp (24)
End Sub
Private Sub Account25Spin_SpinUp()
    shiftAccountUp (25)
End Sub
Private Sub Account26Spin_SpinUp()
    shiftAccountUp (26)
End Sub
Private Sub Account27Spin_SpinUp()
    shiftAccountUp (27)
End Sub
Private Sub Account28Spin_SpinUp()
    shiftAccountUp (28)
End Sub
Private Sub Account29Spin_SpinUp()
    shiftAccountUp (29)
End Sub
Private Sub Account30Spin_SpinUp()
    shiftAccountUp (30)
End Sub
Private Sub Account1Spin_SpinDown()
    shiftAccountDown (1)
End Sub
Private Sub Account2Spin_SpinDown()
    shiftAccountDown (2)
End Sub
Private Sub Account3Spin_SpinDown()
    shiftAccountDown (3)
End Sub
Private Sub Account4Spin_SpinDown()
    shiftAccountDown (4)
End Sub
Private Sub Account5Spin_SpinDown()
    shiftAccountDown (5)
End Sub
Private Sub Account6Spin_SpinDown()
    shiftAccountDown (6)
End Sub
Private Sub Account7Spin_SpinDown()
    shiftAccountDown (7)
End Sub
Private Sub Account8Spin_SpinDown()
    shiftAccountDown (8)
End Sub
Private Sub Account9Spin_SpinDown()
    shiftAccountDown (9)
End Sub
Private Sub Account10Spin_SpinDown()
    shiftAccountDown (10)
End Sub
Private Sub Account11Spin_SpinDown()
    shiftAccountDown (11)
End Sub
Private Sub Account12Spin_SpinDown()
    shiftAccountDown (12)
End Sub
Private Sub Account13Spin_SpinDown()
    shiftAccountDown (13)
End Sub
Private Sub Account14Spin_SpinDown()
    shiftAccountDown (14)
End Sub
Private Sub Account15Spin_SpinDown()
    shiftAccountDown (15)
End Sub
Private Sub Account16Spin_SpinDown()
    shiftAccountDown (16)
End Sub
Private Sub Account17Spin_SpinDown()
    shiftAccountDown (17)
End Sub
Private Sub Account18Spin_SpinDown()
    shiftAccountDown (18)
End Sub
Private Sub Account19Spin_SpinDown()
    shiftAccountDown (19)
End Sub
Private Sub Account20Spin_SpinDown()
    shiftAccountDown (20)
End Sub
Private Sub Account21Spin_SpinDown()
    shiftAccountDown (21)
End Sub
Private Sub Account22Spin_SpinDown()
    shiftAccountDown (22)
End Sub
Private Sub Account23Spin_SpinDown()
    shiftAccountDown (23)
End Sub
Private Sub Account24Spin_SpinDown()
    shiftAccountDown (24)
End Sub
Private Sub Account25Spin_SpinDown()
    shiftAccountDown (25)
End Sub
Private Sub Account26Spin_SpinDown()
    shiftAccountDown (26)
End Sub
Private Sub Account27Spin_SpinDown()
    shiftAccountDown (27)
End Sub
Private Sub Account28Spin_SpinDown()
    shiftAccountDown (28)
End Sub
Private Sub Account29Spin_SpinDown()
    shiftAccountDown (29)
End Sub
Private Sub Account30Spin_SpinDown()
    shiftAccountDown (30)
End Sub
Private Sub Account1Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (1)
End Sub
Private Sub Account2Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (2)
End Sub
Private Sub Account3Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (3)
End Sub
Private Sub Account4Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (4)
End Sub
Private Sub Account5Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (5)
End Sub
Private Sub Account6Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (6)
End Sub
Private Sub Account7Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (7)
End Sub
Private Sub Account8Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (8)
End Sub
Private Sub Account9Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (9)
End Sub
Private Sub Account10Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (10)
End Sub
Private Sub Account11Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (11)
End Sub
Private Sub Account12Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (12)
End Sub
Private Sub Account13Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (13)
End Sub
Private Sub Account14Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (14)
End Sub
Private Sub Account15Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (15)
End Sub
Private Sub Account16Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (16)
End Sub
Private Sub Account17Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (17)
End Sub
Private Sub Account18Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (18)
End Sub
Private Sub Account19Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (19)
End Sub
Private Sub Account20Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (20)
End Sub
Private Sub Account21Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (21)
End Sub
Private Sub Account22Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (22)
End Sub
Private Sub Account23Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (23)
End Sub
Private Sub Account24Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (24)
End Sub
Private Sub Account25Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (25)
End Sub
Private Sub Account26Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (26)
End Sub
Private Sub Account27Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (27)
End Sub
Private Sub Account28Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (28)
End Sub
Private Sub Account29Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (29)
End Sub
Private Sub Account30Input_Change()
    If Not CurrentlyHandlingUpdate Then handleAccountKeypress (30)
End Sub
Private Sub Account1Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (1)
End Sub
Private Sub Account2Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (2)
End Sub
Private Sub Account3Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (3)
End Sub
Private Sub Account4Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (4)
End Sub
Private Sub Account5Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (5)
End Sub
Private Sub Account6Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (6)
End Sub
Private Sub Account7Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (7)
End Sub
Private Sub Account8Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (8)
End Sub
Private Sub Account9Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (9)
End Sub
Private Sub Account10Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (10)
End Sub
Private Sub Account11Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (11)
End Sub
Private Sub Account12Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (12)
End Sub
Private Sub Account13Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (13)
End Sub
Private Sub Account14Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (14)
End Sub
Private Sub Account15Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (15)
End Sub
Private Sub Account16Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (16)
End Sub
Private Sub Account17Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (17)
End Sub
Private Sub Account18Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (18)
End Sub
Private Sub Account19Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (19)
End Sub
Private Sub Account20Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (20)
End Sub
Private Sub Account21Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (21)
End Sub
Private Sub Account22Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (22)
End Sub
Private Sub Account23Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (23)
End Sub
Private Sub Account24Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (24)
End Sub
Private Sub Account25Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (25)
End Sub
Private Sub Account26Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (26)
End Sub
Private Sub Account27Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (27)
End Sub
Private Sub Account28Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (28)
End Sub
Private Sub Account29Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (29)
End Sub
Private Sub Account30Input_AfterUpdate()
    If Not CurrentlyHandlingUpdate Then handleAccountUpdate (30)
End Sub
Private Sub Account1StartBal_AfterUpdate()
    handleStartBalUpdate (1)
End Sub
Private Sub Account2StartBal_AfterUpdate()
    handleStartBalUpdate (2)
End Sub
Private Sub Account3StartBal_AfterUpdate()
    handleStartBalUpdate (3)
End Sub
Private Sub Account4StartBal_AfterUpdate()
    handleStartBalUpdate (4)
End Sub
Private Sub Account5StartBal_AfterUpdate()
    handleStartBalUpdate (5)
End Sub
Private Sub Account6StartBal_AfterUpdate()
    handleStartBalUpdate (6)
End Sub
Private Sub Account7StartBal_AfterUpdate()
    handleStartBalUpdate (7)
End Sub
Private Sub Account8StartBal_AfterUpdate()
    handleStartBalUpdate (8)
End Sub
Private Sub Account9StartBal_AfterUpdate()
    handleStartBalUpdate (9)
End Sub
Private Sub Account10StartBal_AfterUpdate()
    handleStartBalUpdate (10)
End Sub
Private Sub Account11StartBal_AfterUpdate()
    handleStartBalUpdate (11)
End Sub
Private Sub Account12StartBal_AfterUpdate()
    handleStartBalUpdate (12)
End Sub
Private Sub Account13StartBal_AfterUpdate()
    handleStartBalUpdate (13)
End Sub
Private Sub Account14StartBal_AfterUpdate()
    handleStartBalUpdate (14)
End Sub
Private Sub Account15StartBal_AfterUpdate()
    handleStartBalUpdate (15)
End Sub
Private Sub Account16StartBal_AfterUpdate()
    handleStartBalUpdate (16)
End Sub
Private Sub Account17StartBal_AfterUpdate()
    handleStartBalUpdate (17)
End Sub
Private Sub Account18StartBal_AfterUpdate()
    handleStartBalUpdate (18)
End Sub
Private Sub Account19StartBal_AfterUpdate()
    handleStartBalUpdate (19)
End Sub
Private Sub Account20StartBal_AfterUpdate()
    handleStartBalUpdate (20)
End Sub
Private Sub Account21StartBal_AfterUpdate()
    handleStartBalUpdate (21)
End Sub
Private Sub Account22StartBal_AfterUpdate()
    handleStartBalUpdate (22)
End Sub
Private Sub Account23StartBal_AfterUpdate()
    handleStartBalUpdate (23)
End Sub
Private Sub Account24StartBal_AfterUpdate()
    handleStartBalUpdate (24)
End Sub
Private Sub Account25StartBal_AfterUpdate()
    handleStartBalUpdate (25)
End Sub
Private Sub Account26StartBal_AfterUpdate()
    handleStartBalUpdate (26)
End Sub
Private Sub Account27StartBal_AfterUpdate()
    handleStartBalUpdate (27)
End Sub
Private Sub Account28StartBal_AfterUpdate()
    handleStartBalUpdate (28)
End Sub
Private Sub Account29StartBal_AfterUpdate()
    handleStartBalUpdate (29)
End Sub
Private Sub Account30StartBal_AfterUpdate()
    handleStartBalUpdate (30)
End Sub






