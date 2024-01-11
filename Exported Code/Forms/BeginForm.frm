VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BeginForm 
   Caption         =   "Begin New Workbook"
   ClientHeight    =   9000.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7215
   OleObjectBlob   =   "BeginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BeginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DateIsEntered As Boolean
Dim CurrentAddCategoryLabelPosition As Integer
Dim CurrentAddAccountLabelPosition As Integer
Private Sub AddCategoryLabel_Click()
    focusCategory (CurrentAddCategoryLabelPosition)
End Sub
Private Sub AddAccountLabel_Click()
    focusAccount (CurrentAddAccountLabelPosition)
End Sub
Private Sub MultiPage1_Change()
    If MultiPage1.SelectedItem.Index = 0 Then
        MultiPageText.Caption = _
            "Below, you can enter the categories you would like to track. " & _
            "Things like Food, Rent, Gas, etc." & vbNewLine & _
            "These can be changed at any time."
    Else
        MultiPageText.Caption = _
            "Below, you can enter the accounts you would like to track. " & _
            "Things like Checking, Credit, Savings, etc." & vbNewLine & _
            "You should also enter the starting balances at the time you begin tracking. " & _
            "Use negative values where money is owed." & vbNewLine & _
            "These can be changed at any time."
    End If
End Sub
Function positionAddCategoryCaption(CatNum As Integer)
    
    LabelTopValue = 24 + (24 * (CatNum - 1)) + 4
    AddCategoryLabel.Top = LabelTopValue
    
    CurrentAddCategoryLabelPosition = CatNum
    
End Function
Function positionAddAccountCaption(ActNum As Integer)
    
    LabelTopValue = 24 + (24 * (ActNum - 1)) + 4
    AddAccountLabel.Top = LabelTopValue
    
    CurrentAddAccountLabelPosition = ActNum
    
End Function
Private Sub FinishButton_Click()
    
    ' ReturnStartInterval.Caption is already set up
    ReturnCatStr = ""
    For Each Item In bf.getFinalCategoryArray
        If ReturnCatStr = "" Then
            ReturnCatStr = Item
        Else
            ReturnCatStr = ReturnCatStr & "|!DELIM!|" & Item
        End If
    Next
    ReturnCategoriesStr.Caption = ReturnCatStr
    
    ReturnActStr = ""
    For Each Item In bf.getFinalAccountArray
        If ReturnActStr = "" Then
            ReturnActStr = Item
        Else
            ReturnActStr = ReturnActStr & "|!DELIM!|" & Item
        End If
    Next
    ReturnAccountsStr.Caption = ReturnActStr
    
    ReturnActBalStr = ""
    For Each Item In bf.getFinalAccountBalanceArray
        If ReturnActBalStr = "" Then
            ReturnActBalStr = Item
        Else
            ReturnActBalStr = ReturnActBalStr & "|!DELIM!|" & Item
        End If
    Next
    ReturnStartingBalancesStr.Caption = ReturnActBalStr
    
    Me.Hide
    
End Sub
Function getInterval() As String
    If OptWeekly.Value = True Then
        getInterval = "Weekly"
    ElseIf OptBiweekly.Value = True Then
        getInterval = "Bi-Weekly"
    ElseIf OptMonthly.Value = True Then
        getInterval = "Monthly"
    Else
        getInterval = "NA"
    End If
End Function
Function setInterval()

    ' Check if previous interval is weekly or biweekly
    Dim UsedToBeMonthly As Boolean
    UsedToBeMonthly = True
    If txtStartingDay.Value <> "" Then
        PreviousFirst = Split(txtStartingDay.Value, ",")(0)
        If PreviousFirst = "Sunday" Or PreviousFirst = "Monday" Or PreviousFirst = "Tuesday" Or _
                PreviousFirst = "Wednesday" Or PreviousFirst = "Thursday" Or _
                PreviousFirst = "Friday" Or PreviousFirst = "Saturday" Then
            UsedToBeMonthly = False
            
        Else
        End If
    End If
    
    cur = getInterval()
    
    If UsedToBeMonthly And cur = "Weekly" Or _
            UsedToBeMonthly And cur = "Bi-Weekly" Or _
            UsedToBeMonthly = False And cur = "Monthly" Then
        txtStartingDay.Value = ""
        DateIsEntered = False
    End If

    If cur = "Weekly" Or cur = "Bi-Weekly" Then
        DateLabel.Visible = True
        SelectDate.Visible = True
        
        DateLabel.Caption = "What day would you like to begin on?"
        SelectDate.Caption = "Select Date"
        If cur = "Weekly" Then
            ReturnPeriodTypeStr.Caption = "Weekly"
        Else
            ReturnPeriodTypeStr.Caption = "Bi-Weekly"
        End If
    ElseIf cur = "Monthly" Then
        DateLabel.Visible = True
        SelectDate.Visible = True
        
        DateLabel.Caption = "What month would you like to begin on?"
        SelectDate.Caption = "Select Month"
        ReturnPeriodTypeStr.Caption = "Monthly"
    End If
    
    Call updateDates
End Function


Private Sub ConfirmDateButton_Click()
    ConfirmDateButton.Visible = False
    MultiPage1.Visible = True
    FinishButton.Visible = True
    
    OptWeekly.Enabled = False
    OptBiweekly.Enabled = False
    OptMonthly.Enabled = False
    SelectDate.Enabled = False
    IntervalLabel.Enabled = False
    DateLabel.Enabled = False
    MultiPageText.Visible = True
    MultiPageText.Caption = _
            "Below, you can enter the categories you would like to track. " & _
            "Things like Food, Rent, Gas, etc." & vbNewLine & _
            "These can be changed at any time."
End Sub


Private Sub RangeConfirmLabel_Click()

End Sub

Private Sub SelectDate_Click()
    If getInterval() = "Monthly" Then
        MonthSelector.Show
        UserInput = MonthSelector.ReturnValue
        If UserInput <> "" Then
            SelectedMonth = Split(UserInput, "/")(0)
            SelectedYear = Split(UserInput, "/")(1)
            SelectedMonthFormatted = Format(CDate(SelectedMonth & "/1/2000"), "mmmm")
            txtStartingDay.Text = SelectedMonthFormatted & ", " & SelectedYear
            
            ReturnStartInterval.Caption = Format(CDate(SelectedMonth & "/1/2000"), "mmm") & " " & SelectedYear
            DateIsEntered = True
            Call updateDates
        End If
    Else
        DateSelector.Show
        UserInput = DateSelector.ReturnValue
        If UserInput <> "" Then
            SelectedDate = CDate(UserInput)
            SelectedDateFormatted = WeekdayName(Format(SelectedDate, "w")) & ", " & Format(SelectedDate, "mmmm d yyyy")
            txtStartingDay.Text = SelectedDateFormatted
            txtStartingDay.Tag = SelectedDate
            If getInterval() = "Weekly" Then
                ReturnStartInterval.Caption = Format(SelectedDate, "m-d") & " to " & Format(SelectedDate + 6, "m-d")
            Else
                ReturnStartInterval.Caption = Format(SelectedDate, "m-d") & " to " & Format(SelectedDate + 13, "m-d")
            End If
            DateIsEntered = True
            Call updateDates
        End If
    End If
    
End Sub
Function getStartDate() As Date
    On Error Resume Next
        TextInput = CDate(txtStartingDay.Tag)
    On Error GoTo 0

    getStartDate = TextInput
End Function
Function getEndDate() As Date
    Interval = getInterval()
    If Interval = "Weekly" Then
        getEndDate = getStartDate() + 6
    ElseIf Interval = "Bi-Weekly" Then
        getEndDate = getStartDate() + 13
    End If
End Function
Function updateDates()
    ' If using weekly or biweekly
    If DateIsEntered And getInterval() <> "Monthly" Then
        StartDate = getStartDate()
        EndDate = getEndDate()
        SelectedDateFormatted = WeekdayName(Format(StartDate, "w")) & ", " & Format(StartDate, "mmmm d yyyy")
        RangeEndDateFormatted = WeekdayName(Format(EndDate, "w")) & ", " & Format(EndDate, "mmmm d yyyy")
        RangeConfirmLabel.Visible = True
        RangeConfirmLabel.Caption = "Your first sheet will cover the following range:" & vbNewLine _
            & SelectedDateFormatted & vbNewLine & "to" & vbNewLine & RangeEndDateFormatted
        ConfirmDateButton.Visible = True
    ' If using monthly
    ElseIf DateIsEntered And getInterval() = "Monthly" Then
        RangeConfirmLabel.Visible = True
        RangeConfirmLabel.Caption = "Your first sheet will cover:" & vbNewLine _
            & txtStartingDay.Value
        ConfirmDateButton.Visible = True
    Else
        RangeConfirmLabel.Visible = False
        ConfirmDateButton.Visible = False
    End If
End Function









Function deleteCategory(CatNum As Integer)
    CatCount = bf.getVisibleCategoryCount
    
    Call bf.setCategoryValue(CatNum, "")
    
    If CatNum = 30 Then
        bf.maskCategory (30)
    Else
        BelowCategoryValue = bf.getCategoryValue(CatNum + 1)
        For i = CatNum To CatCount - 1 Step 1
            Call bf.setCategoryValue(CInt(i), CStr(BelowCategoryValue))
            BelowCategoryValue = bf.getCategoryValue(i + 2)
        Next
        bf.maskCategory (CatCount - 1)
        positionAddCategoryCaption (CatCount - 1)
        bf.makeCategoryInvisible (CatCount)
    End If
    MultiPage1.CategoriesPage.ScrollHeight = bf.getCategoryBoxScrollHeight()
    
End Function
Function deleteAccount(ActNum As Integer)
    ActCount = bf.getVisibleAccountCount
    
    Call bf.setAccountValue(ActNum, "")
    Call bf.setAccountBalanceValue(ActNum, "")
    
    If ActNum = 30 Then
        bf.maskAccount (30)
    Else
        BelowAccountValue = bf.getAccountValue(ActNum + 1)
        BelowAccountBalanceValue = bf.getAccountBalanceValue(ActNum + 1)
        For i = ActNum To ActCount - 1 Step 1
            Call bf.setAccountValue(CInt(i), CStr(BelowAccountValue))
            Call bf.setAccountBalanceValue(CInt(i), CStr(BelowAccountBalanceValue))
            BelowAccountValue = bf.getAccountValue(i + 2)
            BelowAccountBalanceValue = bf.getAccountBalanceValue(i + 2)
        Next
        bf.maskAccount (ActCount - 1)
        positionAddAccountCaption (ActCount - 1)
        bf.makeAccountInvisible (ActCount)
    End If
    MultiPage1.AccountsPage.ScrollHeight = bf.getAccountBoxScrollHeight()
    
End Function


Function shiftCategoryUp(CatNum As Integer)
    If CatNum = 1 Then Exit Function
    
    Dim ThisValue As String
    Dim AboveValue As String
    
    ThisValue = bf.getCategoryValue(CatNum)
    AboveValue = bf.getCategoryValue(CatNum - 1)
    
    Call bf.setCategoryValue(CatNum - 1, ThisValue)
    Call bf.setCategoryValue(CatNum, AboveValue)
    
End Function
Function shiftAccountUp(ActNum As Integer)
    If ActNum = 1 Then Exit Function
    
    Dim ThisValue As String
    Dim AboveValue As String
    Dim ThisBalanceValue As String
    Dim AboveBalanceValue As String
    
    ThisValue = bf.getAccountValue(ActNum)
    AboveValue = bf.getAccountValue(ActNum - 1)
    ThisBalanceValue = bf.getAccountBalanceValue(ActNum)
    AboveBalanceValue = bf.getAccountBalanceValue(ActNum - 1)
    
    Call bf.setAccountValue(ActNum - 1, ThisValue)
    Call bf.setAccountValue(ActNum, AboveValue)
    Call bf.setAccountBalanceValue(ActNum - 1, ThisBalanceValue)
    Call bf.setAccountBalanceValue(ActNum, AboveBalanceValue)
    
End Function


Function shiftCategoryDown(CatNum As Integer)
    If CatNum >= bf.getVisibleCategoryCount - 1 And bf.getVisibleCategoryCount <> 30 Then Exit Function
    
    Dim ThisValue As String
    Dim BelowValue As String
    
    ThisValue = bf.getCategoryValue(CatNum)
    BelowValue = bf.getCategoryValue(CatNum + 1)
    
    Call bf.setCategoryValue(CatNum, BelowValue)
    Call bf.setCategoryValue(CatNum + 1, ThisValue)
    
End Function
Function shiftAccountDown(ActNum As Integer)
    If ActNum >= bf.getVisibleAccountCount - 1 And bf.getVisibleAccountCount <> 30 Then Exit Function
    
    Dim ThisValue As String
    Dim BelowValue As String
    Dim ThisBalanceValue As String
    Dim BelowBalanceValue As String
    
    ThisValue = bf.getAccountValue(ActNum)
    BelowValue = bf.getAccountValue(ActNum + 1)
    ThisBalanceValue = bf.getAccountBalanceValue(ActNum)
    BelowBalanceValue = bf.getAccountBalanceValue(ActNum + 1)
    
    Call bf.setAccountValue(ActNum, BelowValue)
    Call bf.setAccountValue(ActNum + 1, ThisValue)
    Call bf.setAccountBalanceValue(ActNum, BelowBalanceValue)
    Call bf.setAccountBalanceValue(ActNum + 1, ThisBalanceValue)
    
End Function


Function handleCategoryClick(CatNum As Integer)
    If CatNum = bf.getVisibleCategoryCount Then
        bf.maskCategory (CatNum)
        
        If CatNum <> 30 Then
            bf.makeCategoryVisible (CatNum + 1)
            bf.maskCategory (CatNum + 1)
            positionAddCategoryCaption (CatNum + 1)
        End If
        bf.unmaskCategory (CatNum)
    End If
End Function
Function handleAccountClick(ActNum As Integer)
    If ActNum = bf.getVisibleAccountCount Then
        bf.maskAccount (ActNum)
        
        If ActNum <> 30 Then
            bf.makeAccountVisible (ActNum + 1)
            bf.maskAccount (ActNum + 1)
            positionAddAccountCaption (ActNum + 1)
            
        End If
        bf.unmaskAccount (ActNum)
    End If
End Function

Function enableFinishIfReady()
    If UBound(bf.getFinalCategoryArray) >= 0 And UBound(bf.getFinalAccountArray) >= 0 Then
        FinishButton.Enabled = True
    Else
        FinishButton.Enabled = False
    End If
    
End Function

Function handleCategoryUpdate(CatNum As Integer)
    ' Cannot loop due to being unable to set focus
    ' Must not be numerical, must not be a duplicate
    MultiPage1.CategoriesPage.ScrollHeight = bf.getCategoryBoxScrollHeight()
    InputStr = bf.getCategoryValue(CatNum)
    
    ' Remove trailing spaces
    If InputStr <> Trim(InputStr) Then
        Call bf.setCategoryValue(CatNum, Trim(InputStr))
        InputStr = Trim(InputStr)
    End If
    
    If InputStr = "" Then
        Exit Function
    ElseIf IsNumeric(InputStr) Then
        Call MsgBox("Category names cannot be only numbers.", vbOKOnly, "Invalid Input")
        Call bf.setCategoryValue(CatNum, "")
    Else
        DuplicateFound = False
        For i = 1 To bf.getVisibleCategoryCount() Step 1
            If CatNum <> i And InputStr = bf.getCategoryValue(CInt(i)) Then
                DuplicateFound = True
            End If
        Next
        If DuplicateFound Then
            Call MsgBox("Duplicate category name." & vbNewLine & _
                    Chr(34) & InputStr & Chr(34) & " already exists.", _
                    vbOKOnly, "Invalid Input")
            Call bf.setCategoryValue(CatNum, "")
        End If
    End If
    Call enableFinishIfReady
End Function
Function handleAccountUpdate(ActNum As Integer)
    ' Cannot loop due to being unable to set focus
    ' Must not be numerical, must not be a duplicate
    MultiPage1.CategoriesPage.ScrollHeight = bf.getAccountBoxScrollHeight()
    InputStr = bf.getAccountValue(ActNum)
    
    ' Remove trailing spaces
    If InputStr <> Trim(InputStr) Then
        Call bf.setAccountValue(ActNum, Trim(InputStr))
        InputStr = Trim(InputStr)
    End If
    
    If InputStr = "" Then
        Exit Function
    ElseIf IsNumeric(InputStr) Then
        Call MsgBox("Account names cannot be only numbers.", vbOKOnly, "Invalid Input")
        Call bf.setAccountValue(ActNum, "")
    Else
        DuplicateFound = False
        For i = 1 To bf.getVisibleAccountCount() Step 1
            If ActNum <> i And InputStr = bf.getAccountValue(CInt(i)) Then
                DuplicateFound = True
            End If
        Next
        If DuplicateFound Then
            Call MsgBox("Duplicate Account name." & vbNewLine & _
                    Chr(34) & InputStr & Chr(34) & " already exists.", _
                    vbOKOnly, "Invalid Input")
            Call bf.setAccountValue(ActNum, "")
        End If
    End If
    Call enableFinishIfReady
End Function
Function handleAccountBalanceUpdate(ActNum As Integer)
    InputStr = bf.getAccountBalanceValue(ActNum)
    
    If InputStr <> Trim(InputStr) Then
        Call bf.setAccountBalanceValue(ActNum, Trim(InputStr))
        InputStr = Trim(InputStr)
    End If
    
    If InputStr = "" Then
        Call bf.setAccountBalanceValue(ActNum, "$0.00")
        Exit Function
    ElseIf Not IsNumeric(InputStr) Then
        Call MsgBox("Please enter a numeric input.", vbOKOnly, "Invalid Input")
        Call bf.setAccountBalanceValue(ActNum, "$0.00")
    Else
        Call bf.setAccountBalanceValue(ActNum, Format(InputStr, "$#,##0.00"))
    End If
    
End Function



















































' Unfortunately, no way to export these

Function focusCategory(CatNum As Integer)
    If CatNum = 1 Then MultiPage1.CategoriesPage.Category1Input.SetFocus
    If CatNum = 2 Then MultiPage1.CategoriesPage.Category2Input.SetFocus
    If CatNum = 3 Then MultiPage1.CategoriesPage.Category3Input.SetFocus
    If CatNum = 4 Then MultiPage1.CategoriesPage.Category4Input.SetFocus
    If CatNum = 5 Then MultiPage1.CategoriesPage.Category5Input.SetFocus
    If CatNum = 6 Then MultiPage1.CategoriesPage.Category6Input.SetFocus
    If CatNum = 7 Then MultiPage1.CategoriesPage.Category7Input.SetFocus
    If CatNum = 8 Then MultiPage1.CategoriesPage.Category8Input.SetFocus
    If CatNum = 9 Then MultiPage1.CategoriesPage.Category9Input.SetFocus
    If CatNum = 10 Then MultiPage1.CategoriesPage.Category10Input.SetFocus
    If CatNum = 11 Then MultiPage1.CategoriesPage.Category11Input.SetFocus
    If CatNum = 12 Then MultiPage1.CategoriesPage.Category12Input.SetFocus
    If CatNum = 13 Then MultiPage1.CategoriesPage.Category13Input.SetFocus
    If CatNum = 14 Then MultiPage1.CategoriesPage.Category14Input.SetFocus
    If CatNum = 15 Then MultiPage1.CategoriesPage.Category15Input.SetFocus
    If CatNum = 16 Then MultiPage1.CategoriesPage.Category16Input.SetFocus
    If CatNum = 17 Then MultiPage1.CategoriesPage.Category17Input.SetFocus
    If CatNum = 18 Then MultiPage1.CategoriesPage.Category18Input.SetFocus
    If CatNum = 19 Then MultiPage1.CategoriesPage.Category19Input.SetFocus
    If CatNum = 20 Then MultiPage1.CategoriesPage.Category20Input.SetFocus
    If CatNum = 21 Then MultiPage1.CategoriesPage.Category21Input.SetFocus
    If CatNum = 22 Then MultiPage1.CategoriesPage.Category22Input.SetFocus
    If CatNum = 23 Then MultiPage1.CategoriesPage.Category23Input.SetFocus
    If CatNum = 24 Then MultiPage1.CategoriesPage.Category24Input.SetFocus
    If CatNum = 25 Then MultiPage1.CategoriesPage.Category25Input.SetFocus
    If CatNum = 26 Then MultiPage1.CategoriesPage.Category26Input.SetFocus
    If CatNum = 27 Then MultiPage1.CategoriesPage.Category27Input.SetFocus
    If CatNum = 28 Then MultiPage1.CategoriesPage.Category28Input.SetFocus
    If CatNum = 29 Then MultiPage1.CategoriesPage.Category29Input.SetFocus
    If CatNum = 30 Then MultiPage1.CategoriesPage.Category30Input.SetFocus
End Function
Function focusAccount(ActNum As Integer)
    If ActNum = 1 Then MultiPage1.AccountsPage.Account1Input.SetFocus
    If ActNum = 2 Then MultiPage1.AccountsPage.Account2Input.SetFocus
    If ActNum = 3 Then MultiPage1.AccountsPage.Account3Input.SetFocus
    If ActNum = 4 Then MultiPage1.AccountsPage.Account4Input.SetFocus
    If ActNum = 5 Then MultiPage1.AccountsPage.Account5Input.SetFocus
    If ActNum = 6 Then MultiPage1.AccountsPage.Account6Input.SetFocus
    If ActNum = 7 Then MultiPage1.AccountsPage.Account7Input.SetFocus
    If ActNum = 8 Then MultiPage1.AccountsPage.Account8Input.SetFocus
    If ActNum = 9 Then MultiPage1.AccountsPage.Account9Input.SetFocus
    If ActNum = 10 Then MultiPage1.AccountsPage.Account10Input.SetFocus
    If ActNum = 11 Then MultiPage1.AccountsPage.Account11Input.SetFocus
    If ActNum = 12 Then MultiPage1.AccountsPage.Account12Input.SetFocus
    If ActNum = 13 Then MultiPage1.AccountsPage.Account13Input.SetFocus
    If ActNum = 14 Then MultiPage1.AccountsPage.Account14Input.SetFocus
    If ActNum = 15 Then MultiPage1.AccountsPage.Account15Input.SetFocus
    If ActNum = 16 Then MultiPage1.AccountsPage.Account16Input.SetFocus
    If ActNum = 17 Then MultiPage1.AccountsPage.Account17Input.SetFocus
    If ActNum = 18 Then MultiPage1.AccountsPage.Account18Input.SetFocus
    If ActNum = 19 Then MultiPage1.AccountsPage.Account19Input.SetFocus
    If ActNum = 20 Then MultiPage1.AccountsPage.Account20Input.SetFocus
    If ActNum = 21 Then MultiPage1.AccountsPage.Account21Input.SetFocus
    If ActNum = 22 Then MultiPage1.AccountsPage.Account22Input.SetFocus
    If ActNum = 23 Then MultiPage1.AccountsPage.Account23Input.SetFocus
    If ActNum = 24 Then MultiPage1.AccountsPage.Account24Input.SetFocus
    If ActNum = 25 Then MultiPage1.AccountsPage.Account25Input.SetFocus
    If ActNum = 26 Then MultiPage1.AccountsPage.Account26Input.SetFocus
    If ActNum = 27 Then MultiPage1.AccountsPage.Account27Input.SetFocus
    If ActNum = 28 Then MultiPage1.AccountsPage.Account28Input.SetFocus
    If ActNum = 29 Then MultiPage1.AccountsPage.Account29Input.SetFocus
    If ActNum = 30 Then MultiPage1.AccountsPage.Account30Input.SetFocus
End Function

Private Sub OptBiweekly_Click()
    setInterval
End Sub
Private Sub OptMonthly_Click()
    setInterval
End Sub
Private Sub OptWeekly_Click()
    setInterval
End Sub

Private Sub UserForm_Terminate()
    ReturnStartInterval.Caption = ""
    ReturnCategoriesStr.Caption = ""
    ReturnAccountsStr.Caption = ""
    ReturnStartingBalancesStr.Caption = ""
End Sub
Private Sub UserForm_Initialize()
    DateIsEntered = False
    CurrentAddCategoryLabelPosition = 2
    CurrentAddAccountLabelPosition = 2
    MultiPage1.CategoriesPage.ScrollHeight = 0
    MultiPage1.AccountsPage.ScrollHeight = 0
End Sub
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
    handleCategoryClick (1)
End Sub
Private Sub Category2Input_Change()
    handleCategoryClick (2)
End Sub
Private Sub Category3Input_Change()
    handleCategoryClick (3)
End Sub
Private Sub Category4Input_Change()
    handleCategoryClick (4)
End Sub
Private Sub Category5Input_Change()
    handleCategoryClick (5)
End Sub
Private Sub Category6Input_Change()
    handleCategoryClick (6)
End Sub
Private Sub Category7Input_Change()
    handleCategoryClick (7)
End Sub
Private Sub Category8Input_Change()
    handleCategoryClick (8)
End Sub
Private Sub Category9Input_Change()
    handleCategoryClick (9)
End Sub
Private Sub Category10Input_Change()
    handleCategoryClick (10)
End Sub
Private Sub Category11Input_Change()
    handleCategoryClick (11)
End Sub
Private Sub Category12Input_Change()
    handleCategoryClick (12)
End Sub
Private Sub Category13Input_Change()
    handleCategoryClick (13)
End Sub
Private Sub Category14Input_Change()
    handleCategoryClick (14)
End Sub
Private Sub Category15Input_Change()
    handleCategoryClick (15)
End Sub
Private Sub Category16Input_Change()
    handleCategoryClick (16)
End Sub
Private Sub Category17Input_Change()
    handleCategoryClick (17)
End Sub
Private Sub Category18Input_Change()
    handleCategoryClick (18)
End Sub
Private Sub Category19Input_Change()
    handleCategoryClick (19)
End Sub
Private Sub Category20Input_Change()
    handleCategoryClick (20)
End Sub
Private Sub Category21Input_Change()
    handleCategoryClick (21)
End Sub
Private Sub Category22Input_Change()
    handleCategoryClick (22)
End Sub
Private Sub Category23Input_Change()
    handleCategoryClick (23)
End Sub
Private Sub Category24Input_Change()
    handleCategoryClick (24)
End Sub
Private Sub Category25Input_Change()
    handleCategoryClick (25)
End Sub
Private Sub Category26Input_Change()
    handleCategoryClick (26)
End Sub
Private Sub Category27Input_Change()
    handleCategoryClick (27)
End Sub
Private Sub Category28Input_Change()
    handleCategoryClick (28)
End Sub
Private Sub Category29Input_Change()
    handleCategoryClick (29)
End Sub
Private Sub Category30Input_Change()
    handleCategoryClick (30)
End Sub
Private Sub Category1Input_AfterUpdate()
    handleCategoryUpdate (1)
End Sub
Private Sub Category2Input_AfterUpdate()
    handleCategoryUpdate (2)
End Sub
Private Sub Category3Input_AfterUpdate()
    handleCategoryUpdate (3)
End Sub
Private Sub Category4Input_AfterUpdate()
    handleCategoryUpdate (4)
End Sub
Private Sub Category5Input_AfterUpdate()
    handleCategoryUpdate (5)
End Sub
Private Sub Category6Input_AfterUpdate()
    handleCategoryUpdate (6)
End Sub
Private Sub Category7Input_AfterUpdate()
    handleCategoryUpdate (7)
End Sub
Private Sub Category8Input_AfterUpdate()
    handleCategoryUpdate (8)
End Sub
Private Sub Category9Input_AfterUpdate()
    handleCategoryUpdate (9)
End Sub
Private Sub Category10Input_AfterUpdate()
    handleCategoryUpdate (10)
End Sub
Private Sub Category11Input_AfterUpdate()
    handleCategoryUpdate (11)
End Sub
Private Sub Category12Input_AfterUpdate()
    handleCategoryUpdate (12)
End Sub
Private Sub Category13Input_AfterUpdate()
    handleCategoryUpdate (13)
End Sub
Private Sub Category14Input_AfterUpdate()
    handleCategoryUpdate (14)
End Sub
Private Sub Category15Input_AfterUpdate()
    handleCategoryUpdate (15)
End Sub
Private Sub Category16Input_AfterUpdate()
    handleCategoryUpdate (16)
End Sub
Private Sub Category17Input_AfterUpdate()
    handleCategoryUpdate (17)
End Sub
Private Sub Category18Input_AfterUpdate()
    handleCategoryUpdate (18)
End Sub
Private Sub Category19Input_AfterUpdate()
    handleCategoryUpdate (19)
End Sub
Private Sub Category20Input_AfterUpdate()
    handleCategoryUpdate (20)
End Sub
Private Sub Category21Input_AfterUpdate()
    handleCategoryUpdate (21)
End Sub
Private Sub Category22Input_AfterUpdate()
    handleCategoryUpdate (22)
End Sub
Private Sub Category23Input_AfterUpdate()
    handleCategoryUpdate (23)
End Sub
Private Sub Category24Input_AfterUpdate()
    handleCategoryUpdate (24)
End Sub
Private Sub Category25Input_AfterUpdate()
    handleCategoryUpdate (25)
End Sub
Private Sub Category26Input_AfterUpdate()
    handleCategoryUpdate (26)
End Sub
Private Sub Category27Input_AfterUpdate()
    handleCategoryUpdate (27)
End Sub
Private Sub Category28Input_AfterUpdate()
    handleCategoryUpdate (28)
End Sub
Private Sub Category29Input_AfterUpdate()
    handleCategoryUpdate (29)
End Sub
Private Sub Category30Input_AfterUpdate()
    handleCategoryUpdate (30)
End Sub

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
    handleAccountClick (1)
End Sub
Private Sub Account2Input_Change()
    handleAccountClick (2)
End Sub
Private Sub Account3Input_Change()
    handleAccountClick (3)
End Sub
Private Sub Account4Input_Change()
    handleAccountClick (4)
End Sub
Private Sub Account5Input_Change()
    handleAccountClick (5)
End Sub
Private Sub Account6Input_Change()
    handleAccountClick (6)
End Sub
Private Sub Account7Input_Change()
    handleAccountClick (7)
End Sub
Private Sub Account8Input_Change()
    handleAccountClick (8)
End Sub
Private Sub Account9Input_Change()
    handleAccountClick (9)
End Sub
Private Sub Account10Input_Change()
    handleAccountClick (10)
End Sub
Private Sub Account11Input_Change()
    handleAccountClick (11)
End Sub
Private Sub Account12Input_Change()
    handleAccountClick (12)
End Sub
Private Sub Account13Input_Change()
    handleAccountClick (13)
End Sub
Private Sub Account14Input_Change()
    handleAccountClick (14)
End Sub
Private Sub Account15Input_Change()
    handleAccountClick (15)
End Sub
Private Sub Account16Input_Change()
    handleAccountClick (16)
End Sub
Private Sub Account17Input_Change()
    handleAccountClick (17)
End Sub
Private Sub Account18Input_Change()
    handleAccountClick (18)
End Sub
Private Sub Account19Input_Change()
    handleAccountClick (19)
End Sub
Private Sub Account20Input_Change()
    handleAccountClick (20)
End Sub
Private Sub Account21Input_Change()
    handleAccountClick (21)
End Sub
Private Sub Account22Input_Change()
    handleAccountClick (22)
End Sub
Private Sub Account23Input_Change()
    handleAccountClick (23)
End Sub
Private Sub Account24Input_Change()
    handleAccountClick (24)
End Sub
Private Sub Account25Input_Change()
    handleAccountClick (25)
End Sub
Private Sub Account26Input_Change()
    handleAccountClick (26)
End Sub
Private Sub Account27Input_Change()
    handleAccountClick (27)
End Sub
Private Sub Account28Input_Change()
    handleAccountClick (28)
End Sub
Private Sub Account29Input_Change()
    handleAccountClick (29)
End Sub
Private Sub Account30Input_Change()
    handleAccountClick (30)
End Sub
Private Sub Account1Input_AfterUpdate()
    handleAccountUpdate (1)
End Sub
Private Sub Account2Input_AfterUpdate()
    handleAccountUpdate (2)
End Sub
Private Sub Account3Input_AfterUpdate()
    handleAccountUpdate (3)
End Sub
Private Sub Account4Input_AfterUpdate()
    handleAccountUpdate (4)
End Sub
Private Sub Account5Input_AfterUpdate()
    handleAccountUpdate (5)
End Sub
Private Sub Account6Input_AfterUpdate()
    handleAccountUpdate (6)
End Sub
Private Sub Account7Input_AfterUpdate()
    handleAccountUpdate (7)
End Sub
Private Sub Account8Input_AfterUpdate()
    handleAccountUpdate (8)
End Sub
Private Sub Account9Input_AfterUpdate()
    handleAccountUpdate (9)
End Sub
Private Sub Account10Input_AfterUpdate()
    handleAccountUpdate (10)
End Sub
Private Sub Account11Input_AfterUpdate()
    handleAccountUpdate (11)
End Sub
Private Sub Account12Input_AfterUpdate()
    handleAccountUpdate (12)
End Sub
Private Sub Account13Input_AfterUpdate()
    handleAccountUpdate (13)
End Sub
Private Sub Account14Input_AfterUpdate()
    handleAccountUpdate (14)
End Sub
Private Sub Account15Input_AfterUpdate()
    handleAccountUpdate (15)
End Sub
Private Sub Account16Input_AfterUpdate()
    handleAccountUpdate (16)
End Sub
Private Sub Account17Input_AfterUpdate()
    handleAccountUpdate (17)
End Sub
Private Sub Account18Input_AfterUpdate()
    handleAccountUpdate (18)
End Sub
Private Sub Account19Input_AfterUpdate()
    handleAccountUpdate (19)
End Sub
Private Sub Account20Input_AfterUpdate()
    handleAccountUpdate (20)
End Sub
Private Sub Account21Input_AfterUpdate()
    handleAccountUpdate (21)
End Sub
Private Sub Account22Input_AfterUpdate()
    handleAccountUpdate (22)
End Sub
Private Sub Account23Input_AfterUpdate()
    handleAccountUpdate (23)
End Sub
Private Sub Account24Input_AfterUpdate()
    handleAccountUpdate (24)
End Sub
Private Sub Account25Input_AfterUpdate()
    handleAccountUpdate (25)
End Sub
Private Sub Account26Input_AfterUpdate()
    handleAccountUpdate (26)
End Sub
Private Sub Account27Input_AfterUpdate()
    handleAccountUpdate (27)
End Sub
Private Sub Account28Input_AfterUpdate()
    handleAccountUpdate (28)
End Sub
Private Sub Account29Input_AfterUpdate()
    handleAccountUpdate (29)
End Sub
Private Sub Account30Input_AfterUpdate()
    handleAccountUpdate (30)
End Sub
Private Sub Account1BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (1)
End Sub
Private Sub Account2BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (2)
End Sub
Private Sub Account3BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (3)
End Sub
Private Sub Account4BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (4)
End Sub
Private Sub Account5BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (5)
End Sub
Private Sub Account6BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (6)
End Sub
Private Sub Account7BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (7)
End Sub
Private Sub Account8BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (8)
End Sub
Private Sub Account9BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (9)
End Sub
Private Sub Account10BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (10)
End Sub
Private Sub Account11BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (11)
End Sub
Private Sub Account12BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (12)
End Sub
Private Sub Account13BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (13)
End Sub
Private Sub Account14BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (14)
End Sub
Private Sub Account15BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (15)
End Sub
Private Sub Account16BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (16)
End Sub
Private Sub Account17BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (17)
End Sub
Private Sub Account18BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (18)
End Sub
Private Sub Account19BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (19)
End Sub
Private Sub Account20BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (20)
End Sub
Private Sub Account21BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (21)
End Sub
Private Sub Account22BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (22)
End Sub
Private Sub Account23BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (23)
End Sub
Private Sub Account24BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (24)
End Sub
Private Sub Account25BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (25)
End Sub
Private Sub Account26BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (26)
End Sub
Private Sub Account27BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (27)
End Sub
Private Sub Account28BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (28)
End Sub
Private Sub Account29BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (29)
End Sub
Private Sub Account30BalanceInput_AfterUpdate()
    handleAccountBalanceUpdate (30)
End Sub




