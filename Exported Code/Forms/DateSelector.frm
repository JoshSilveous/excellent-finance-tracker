VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateSelector 
   Caption         =   "Starting Date"
   ClientHeight    =   5676
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5310
   OleObjectBlob   =   "DateSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function selectCell(RowNum As Integer, ColNum As Integer)

    ReturnValue.Caption = getCellTag(RowNum, ColNum)
    Me.Hide
End Function
Private Function renderCalendar(MonthInt As Integer, YearInt As Integer)
    ' Set Month & Year Text
    labmonth.Caption = MonthName(MonthInt)
    labyear.Caption = YearInt
    
    ' Get DayOfWeek for first date in month
    FirstDay = CDate(MonthInt & "/1/" & YearInt)
    FirstDayCol = Weekday(FirstDay)
    
    ' Get PrevMonth Month/Year
    Dim PrevMonthInt As Integer
    PrevMonthInt = MonthInt - 1
    Dim PrevYearInt As Integer
    PrevYearInt = YearInt
    If PrevMonthInt = 0 Then
        PrevMonthInt = 12
        PrevYearInt = PrevYearInt - 1
    End If

    
    CurrentDay = CInt(Format(Date, "d"))
    CurrentMonth = CInt(Format(Date, "m"))
    CurrentYear = CInt(Format(Date, "yyyy"))
    
    DaysInMonth = getDaysInMonth(MonthInt, YearInt)
    DaysInPrevMonth = getDaysInMonth(PrevMonthInt, PrevYearInt)
    DayCount = 1
    DaysAfterCount = 1
    Dim ThisCellTag As String
    
    For ThisRow = 1 To 6 Step 1
        For ThisCol = 1 To 7 Step 1
        
            
            ' Handle days BEFORE DAY 1
            If ThisRow = 1 And ThisCol < FirstDayCol Then
            
                CellVal = ThisCol - FirstDayCol + DaysInPrevMonth + 1
                
                
                If MonthInt = 1 Then
                    ThisCellTag = "12/" & CellVal & "/" & YearInt - 1
                Else
                    ThisCellTag = MonthInt - 1 & "/" & CellVal & "/" & YearInt
                End If
            
                ' Highlight current date
                If MonthInt - 1 = CurrentMonth And YearInt = CurrentYear And CellVal = CurrentDay Then
                    Call setCellText(CInt(ThisRow), CInt(ThisCol), CStr(CellVal), ThisCellTag, &HFFC0C0)
                Else
                    Call setCellText(CInt(ThisRow), CInt(ThisCol), CStr(CellVal), ThisCellTag, &HC0C0C0)
                End If
            
                
            ' Handle days AFTER DAY 1
            ElseIf DayCount > DaysInMonth Then
            
                NextMonthInt = MonthInt + 1
                NextMonthYear = YearInt
                If NextMonthInt = 13 Then
                    NextMonthInt = 1
                    NextMonthYear = NextMonthYear + 1
                End If
                
                ThisCellTag = NextMonthInt & "/" & DaysAfterCount & "/" & NextMonthYear
                
                ' Highlight current date
            
                If NextMonthInt = CurrentMonth And NextMonthYear = CurrentYear And DaysAfterCount = CurrentDay Then
                    Call setCellText(CInt(ThisRow), CInt(ThisCol), CStr(DaysAfterCount), ThisCellTag, &HFFC0C0)
                Else
                    Call setCellText(CInt(ThisRow), CInt(ThisCol), CStr(DaysAfterCount), ThisCellTag, &HC0C0C0)
                End If
                DaysAfterCount = DaysAfterCount + 1
                
            ' Handle days INSIDE THE MONTH
            Else
                ThisCellTag = MonthInt & "/" & DayCount & "/" & YearInt
                
                ' Highlight current date
                If MonthInt = CurrentMonth And YearInt = CurrentYear And DayCount = CurrentDay Then
                    Call setCellText(CInt(ThisRow), CInt(ThisCol), CStr(DayCount), ThisCellTag, &HFFC0C0)
                Else
                    Call setCellText(CInt(ThisRow), CInt(ThisCol), CStr(DayCount), ThisCellTag)
                End If
                DayCount = DayCount + 1
            End If
            
            
            ' Overflows into a sixth row if applicable
            If ThisRow = 6 And ThisCol = 1 And DaysAfterCount > 1 Then
                datecell_6_1.Visible = False
                datecell_6_2.Visible = False
                datecell_6_3.Visible = False
                datecell_6_4.Visible = False
                datecell_6_5.Visible = False
                datecell_6_6.Visible = False
                datecell_6_7.Visible = False
            ElseIf ThisRow = 6 And ThisCol = 1 Then
                datecell_6_1.Visible = True
                datecell_6_2.Visible = True
                datecell_6_3.Visible = True
                datecell_6_4.Visible = True
                datecell_6_5.Visible = True
                datecell_6_6.Visible = True
                datecell_6_7.Visible = True
            End If
                
            
        Next
    Next
    
End Function
Private Function getDaysInMonth(MonthInt As Integer, YearInt As Integer) As Integer
    NxtMonth = MonthInt + 1
    NxtYear = YearInt
    If NxtMonth = 13 Then
        NxtMonth = 1
        NxtYear = NxtYear + 1
    End If
    
    getDaysInMonth = CDate(NxtMonth & "/1/" & NxtYear) - CDate(MonthInt & "/1/" & YearInt)
End Function

























Private Function setCellText(RowNum As Integer, ColNum As Integer, CellText As String, CellTag As String, Optional CellColor As Long)
    If RowNum = 1 Then
        If ColNum = 1 Then
            datecell_1_1.Caption = CellText
            datecell_1_1.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_1.BackColor = &H8000000F
            Else
                datecell_1_1.BackColor = CellColor
            End If
            
        ElseIf ColNum = 2 Then
            datecell_1_2.Caption = CellText
            datecell_1_2.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_2.BackColor = &H8000000F
            Else
                datecell_1_2.BackColor = CellColor
            End If
            
        ElseIf ColNum = 3 Then
            datecell_1_3.Caption = CellText
            datecell_1_3.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_3.BackColor = &H8000000F
            Else
                datecell_1_3.BackColor = CellColor
            End If
            
        ElseIf ColNum = 4 Then
            datecell_1_4.Caption = CellText
            datecell_1_4.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_4.BackColor = &H8000000F
            Else
                datecell_1_4.BackColor = CellColor
            End If
            
        ElseIf ColNum = 5 Then
            datecell_1_5.Caption = CellText
            datecell_1_5.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_5.BackColor = &H8000000F
            Else
                datecell_1_5.BackColor = CellColor
            End If
            
        ElseIf ColNum = 6 Then
            datecell_1_6.Caption = CellText
            datecell_1_6.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_6.BackColor = &H8000000F
            Else
                datecell_1_6.BackColor = CellColor
            End If
            
        ElseIf ColNum = 7 Then
            datecell_1_7.Caption = CellText
            datecell_1_7.Tag = CellTag
            If CellColor = 0 Then
                datecell_1_7.BackColor = &H8000000F
            Else
                datecell_1_7.BackColor = CellColor
            End If
            
        End If
    ElseIf RowNum = 2 Then
        If ColNum = 1 Then
            datecell_2_1.Caption = CellText
            datecell_2_1.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_1.BackColor = &H8000000F
            Else
                datecell_2_1.BackColor = CellColor
            End If
            
        ElseIf ColNum = 2 Then
            datecell_2_2.Caption = CellText
            datecell_2_2.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_2.BackColor = &H8000000F
            Else
                datecell_2_2.BackColor = CellColor
            End If
            
        ElseIf ColNum = 3 Then
            datecell_2_3.Caption = CellText
            datecell_2_3.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_3.BackColor = &H8000000F
            Else
                datecell_2_3.BackColor = CellColor
            End If
            
        ElseIf ColNum = 4 Then
            datecell_2_4.Caption = CellText
            datecell_2_4.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_4.BackColor = &H8000000F
            Else
                datecell_2_4.BackColor = CellColor
            End If
            
        ElseIf ColNum = 5 Then
            datecell_2_5.Caption = CellText
            datecell_2_5.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_5.BackColor = &H8000000F
            Else
                datecell_2_5.BackColor = CellColor
            End If
            
        ElseIf ColNum = 6 Then
            datecell_2_6.Caption = CellText
            datecell_2_6.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_6.BackColor = &H8000000F
            Else
                datecell_2_6.BackColor = CellColor
            End If
            
        ElseIf ColNum = 7 Then
            datecell_2_7.Caption = CellText
            datecell_2_7.Tag = CellTag
            If CellColor = 0 Then
                datecell_2_7.BackColor = &H8000000F
            Else
                datecell_2_7.BackColor = CellColor
            End If
            
        End If
    ElseIf RowNum = 3 Then
        If ColNum = 1 Then
            datecell_3_1.Caption = CellText
            datecell_3_1.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_1.BackColor = &H8000000F
            Else
                datecell_3_1.BackColor = CellColor
            End If
            
        ElseIf ColNum = 2 Then
            datecell_3_2.Caption = CellText
            datecell_3_2.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_2.BackColor = &H8000000F
            Else
                datecell_3_2.BackColor = CellColor
            End If
            
        ElseIf ColNum = 3 Then
            datecell_3_3.Caption = CellText
            datecell_3_3.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_3.BackColor = &H8000000F
            Else
                datecell_3_3.BackColor = CellColor
            End If
            
        ElseIf ColNum = 4 Then
            datecell_3_4.Caption = CellText
            datecell_3_4.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_4.BackColor = &H8000000F
            Else
                datecell_3_4.BackColor = CellColor
            End If
            
        ElseIf ColNum = 5 Then
            datecell_3_5.Caption = CellText
            datecell_3_5.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_5.BackColor = &H8000000F
            Else
                datecell_3_5.BackColor = CellColor
            End If
            
        ElseIf ColNum = 6 Then
            datecell_3_6.Caption = CellText
            datecell_3_6.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_6.BackColor = &H8000000F
            Else
                datecell_3_6.BackColor = CellColor
            End If
            
        ElseIf ColNum = 7 Then
            datecell_3_7.Caption = CellText
            datecell_3_7.Tag = CellTag
            If CellColor = 0 Then
                datecell_3_7.BackColor = &H8000000F
            Else
                datecell_3_7.BackColor = CellColor
            End If
            
        End If
    ElseIf RowNum = 4 Then
        If ColNum = 1 Then
            datecell_4_1.Caption = CellText
            datecell_4_1.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_1.BackColor = &H8000000F
            Else
                datecell_4_1.BackColor = CellColor
            End If
            
        ElseIf ColNum = 2 Then
            datecell_4_2.Caption = CellText
            datecell_4_2.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_2.BackColor = &H8000000F
            Else
                datecell_4_2.BackColor = CellColor
            End If
            
        ElseIf ColNum = 3 Then
            datecell_4_3.Caption = CellText
            datecell_4_3.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_3.BackColor = &H8000000F
            Else
                datecell_4_3.BackColor = CellColor
            End If
            
        ElseIf ColNum = 4 Then
            datecell_4_4.Caption = CellText
            datecell_4_4.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_4.BackColor = &H8000000F
            Else
                datecell_4_4.BackColor = CellColor
            End If
            
        ElseIf ColNum = 5 Then
            datecell_4_5.Caption = CellText
            datecell_4_5.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_5.BackColor = &H8000000F
            Else
                datecell_4_5.BackColor = CellColor
            End If
            
        ElseIf ColNum = 6 Then
            datecell_4_6.Caption = CellText
            datecell_4_6.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_6.BackColor = &H8000000F
            Else
                datecell_4_6.BackColor = CellColor
            End If
            
        ElseIf ColNum = 7 Then
            datecell_4_7.Caption = CellText
            datecell_4_7.Tag = CellTag
            If CellColor = 0 Then
                datecell_4_7.BackColor = &H8000000F
            Else
                datecell_4_7.BackColor = CellColor
            End If
            
        End If
    ElseIf RowNum = 5 Then
        If ColNum = 1 Then
            datecell_5_1.Caption = CellText
            datecell_5_1.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_1.BackColor = &H8000000F
            Else
                datecell_5_1.BackColor = CellColor
            End If
            
        ElseIf ColNum = 2 Then
            datecell_5_2.Caption = CellText
            datecell_5_2.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_2.BackColor = &H8000000F
            Else
                datecell_5_2.BackColor = CellColor
            End If
            
        ElseIf ColNum = 3 Then
            datecell_5_3.Caption = CellText
            datecell_5_3.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_3.BackColor = &H8000000F
            Else
                datecell_5_3.BackColor = CellColor
            End If
            
        ElseIf ColNum = 4 Then
            datecell_5_4.Caption = CellText
            datecell_5_4.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_4.BackColor = &H8000000F
            Else
                datecell_5_4.BackColor = CellColor
            End If
            
        ElseIf ColNum = 5 Then
            datecell_5_5.Caption = CellText
            datecell_5_5.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_5.BackColor = &H8000000F
            Else
                datecell_5_5.BackColor = CellColor
            End If
            
        ElseIf ColNum = 6 Then
            datecell_5_6.Caption = CellText
            datecell_5_6.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_6.BackColor = &H8000000F
            Else
                datecell_5_6.BackColor = CellColor
            End If
            
        ElseIf ColNum = 7 Then
            datecell_5_7.Caption = CellText
            datecell_5_7.Tag = CellTag
            If CellColor = 0 Then
                datecell_5_7.BackColor = &H8000000F
            Else
                datecell_5_7.BackColor = CellColor
            End If
        End If
    ElseIf RowNum = 6 Then
        If ColNum = 1 Then
            datecell_6_1.Caption = CellText
            datecell_6_1.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_1.BackColor = &H8000000F
            Else
                datecell_6_1.BackColor = CellColor
            End If
            
        ElseIf ColNum = 2 Then
            datecell_6_2.Caption = CellText
            datecell_6_2.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_2.BackColor = &H8000000F
            Else
                datecell_6_2.BackColor = CellColor
            End If
            
        ElseIf ColNum = 3 Then
            datecell_6_3.Caption = CellText
            datecell_6_3.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_3.BackColor = &H8000000F
            Else
                datecell_6_3.BackColor = CellColor
            End If
            
        ElseIf ColNum = 4 Then
            datecell_6_4.Caption = CellText
            datecell_6_4.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_4.BackColor = &H8000000F
            Else
                datecell_6_4.BackColor = CellColor
            End If
            
        ElseIf ColNum = 5 Then
            datecell_6_5.Caption = CellText
            datecell_6_5.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_5.BackColor = &H8000000F
            Else
                datecell_6_5.BackColor = CellColor
            End If
            
        ElseIf ColNum = 6 Then
            datecell_6_6.Caption = CellText
            datecell_6_6.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_6.BackColor = &H8000000F
            Else
                datecell_6_6.BackColor = CellColor
            End If
            
        ElseIf ColNum = 7 Then
            datecell_6_7.Caption = CellText
            datecell_6_7.Tag = CellTag
            If CellColor = 0 Then
                datecell_6_7.BackColor = &H8000000F
            Else
                datecell_6_7.BackColor = CellColor
            End If
            
        End If
    End If
End Function
Private Sub butDecrMonth_Click()
    CurMonth = Month(CDate(labmonth.Caption & " 1"))
    
    Dim PrevMonthInt As Integer
    Dim PrevYearInt As Integer
    
    PrevMonthInt = CurMonth - 1
    PrevYearInt = CInt(labyear.Caption)
    If PrevMonthInt = 0 Then
        PrevMonthInt = 12
        PrevYearInt = PrevYearInt - 1
    End If
    
    Call renderCalendar(PrevMonthInt, PrevYearInt)
End Sub
Private Sub butDecrYear_Click()
    CurMonth = Month(CDate(labmonth.Caption & " 1"))

    Dim NextYearInt As Integer
    NextYearInt = CInt(labyear.Caption) - 1
    
    Call renderCalendar(CInt(CurMonth), NextYearInt)
End Sub
Private Sub butIncrMonth_Click()
    CurMonth = Month(CDate(labmonth.Caption & " 1"))
    
    Dim NextMonthInt As Integer
    Dim NextYearInt As Integer
    
    NextMonthInt = CurMonth + 1
    NextYearInt = CInt(labyear.Caption)
    If NextMonthInt = 13 Then
        NextMonthInt = 1
        NextYearInt = NextYearInt + 1
    End If
    
    Call renderCalendar(NextMonthInt, NextYearInt)
End Sub
Private Sub butIncrYear_Click()
    CurMonth = Month(CDate(labmonth.Caption & " 1"))

    Dim NextYearInt As Integer
    NextYearInt = CInt(labyear.Caption) + 1
    
    Call renderCalendar(CInt(CurMonth), NextYearInt)
End Sub
Private Sub Label10_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim MonthNum As Integer
    MonthNum = Month(Date)
    Dim YearNum As Integer
    YearNum = Year(Date)
    Call renderCalendar(MonthNum, YearNum)
End Sub
Private Sub datecell_1_1_Click()
    Call selectCell(1, 1)
End Sub
Private Sub datecell_1_2_Click()
    Call selectCell(1, 2)
End Sub
Private Sub datecell_1_3_Click()
    Call selectCell(1, 3)
End Sub
Private Sub datecell_1_4_Click()
    Call selectCell(1, 4)
End Sub
Private Sub datecell_1_5_Click()
    Call selectCell(1, 5)
End Sub
Private Sub datecell_1_6_Click()
    Call selectCell(1, 6)
End Sub
Private Sub datecell_1_7_Click()
    Call selectCell(1, 7)
End Sub
Private Sub datecell_2_1_Click()
    Call selectCell(2, 1)
End Sub
Private Sub datecell_2_2_Click()
    Call selectCell(2, 2)
End Sub
Private Sub datecell_2_3_Click()
    Call selectCell(2, 3)
End Sub
Private Sub datecell_2_4_Click()
    Call selectCell(2, 4)
End Sub
Private Sub datecell_2_5_Click()
    Call selectCell(2, 5)
End Sub
Private Sub datecell_2_6_Click()
    Call selectCell(2, 6)
End Sub
Private Sub datecell_2_7_Click()
    Call selectCell(2, 7)
End Sub
Private Sub datecell_3_1_Click()
    Call selectCell(3, 1)
End Sub
Private Sub datecell_3_2_Click()
    Call selectCell(3, 2)
End Sub
Private Sub datecell_3_3_Click()
    Call selectCell(3, 3)
End Sub
Private Sub datecell_3_4_Click()
    Call selectCell(3, 4)
End Sub
Private Sub datecell_3_5_Click()
    Call selectCell(3, 5)
End Sub
Private Sub datecell_3_6_Click()
    Call selectCell(3, 6)
End Sub
Private Sub datecell_3_7_Click()
    Call selectCell(3, 7)
End Sub
Private Sub datecell_4_1_Click()
    Call selectCell(4, 1)
End Sub
Private Sub datecell_4_2_Click()
    Call selectCell(4, 2)
End Sub
Private Sub datecell_4_3_Click()
    Call selectCell(4, 3)
End Sub
Private Sub datecell_4_4_Click()
    Call selectCell(4, 4)
End Sub
Private Sub datecell_4_5_Click()
    Call selectCell(4, 5)
End Sub
Private Sub datecell_4_6_Click()
    Call selectCell(4, 6)
End Sub
Private Sub datecell_4_7_Click()
    Call selectCell(4, 7)
End Sub
Private Sub datecell_5_1_Click()
    Call selectCell(5, 1)
End Sub
Private Sub datecell_5_2_Click()
    Call selectCell(5, 2)
End Sub
Private Sub datecell_5_3_Click()
    Call selectCell(5, 3)
End Sub
Private Sub datecell_5_4_Click()
    Call selectCell(5, 4)
End Sub
Private Sub datecell_5_5_Click()
    Call selectCell(5, 5)
End Sub
Private Sub datecell_5_6_Click()
    Call selectCell(5, 6)
End Sub
Private Sub datecell_5_7_Click()
    Call selectCell(5, 7)
End Sub
Private Sub datecell_6_1_Click()
    Call selectCell(6, 1)
End Sub
Private Sub datecell_6_2_Click()
    Call selectCell(6, 2)
End Sub
Private Sub datecell_6_3_Click()
    Call selectCell(6, 3)
End Sub
Private Sub datecell_6_4_Click()
    Call selectCell(6, 4)
End Sub
Private Sub datecell_6_5_Click()
    Call selectCell(6, 5)
End Sub
Private Sub datecell_6_6_Click()
    Call selectCell(6, 6)
End Sub
Private Sub datecell_6_7_Click()
    Call selectCell(6, 7)
End Sub
Private Function getCellTag(RowNum As Integer, ColNum As Integer) As String
    
    If RowNum = 1 Then
        If ColNum = 1 Then
            getCellTag = datecell_1_1.Tag
        ElseIf ColNum = 2 Then
            getCellTag = datecell_1_2.Tag
        ElseIf ColNum = 3 Then
            getCellTag = datecell_1_3.Tag
        ElseIf ColNum = 4 Then
            getCellTag = datecell_1_4.Tag
        ElseIf ColNum = 5 Then
            getCellTag = datecell_1_5.Tag
        ElseIf ColNum = 6 Then
            getCellTag = datecell_1_6.Tag
        ElseIf ColNum = 7 Then
            getCellTag = datecell_1_7.Tag
        End If
    ElseIf RowNum = 2 Then
        If ColNum = 1 Then
            getCellTag = datecell_2_1.Tag
        ElseIf ColNum = 2 Then
            getCellTag = datecell_2_2.Tag
        ElseIf ColNum = 3 Then
            getCellTag = datecell_2_3.Tag
        ElseIf ColNum = 4 Then
            getCellTag = datecell_2_4.Tag
        ElseIf ColNum = 5 Then
            getCellTag = datecell_2_5.Tag
        ElseIf ColNum = 6 Then
            getCellTag = datecell_2_6.Tag
        ElseIf ColNum = 7 Then
            getCellTag = datecell_2_7.Tag
        End If
    ElseIf RowNum = 3 Then
        If ColNum = 1 Then
            getCellTag = datecell_3_1.Tag
        ElseIf ColNum = 2 Then
            getCellTag = datecell_3_2.Tag
        ElseIf ColNum = 3 Then
            getCellTag = datecell_3_3.Tag
        ElseIf ColNum = 4 Then
            getCellTag = datecell_3_4.Tag
        ElseIf ColNum = 5 Then
            getCellTag = datecell_3_5.Tag
        ElseIf ColNum = 6 Then
            getCellTag = datecell_3_6.Tag
        ElseIf ColNum = 7 Then
            getCellTag = datecell_3_7.Tag
        End If
    ElseIf RowNum = 4 Then
        If ColNum = 1 Then
            getCellTag = datecell_4_1.Tag
        ElseIf ColNum = 2 Then
            getCellTag = datecell_4_2.Tag
        ElseIf ColNum = 3 Then
            getCellTag = datecell_4_3.Tag
        ElseIf ColNum = 4 Then
            getCellTag = datecell_4_4.Tag
        ElseIf ColNum = 5 Then
            getCellTag = datecell_4_5.Tag
        ElseIf ColNum = 6 Then
            getCellTag = datecell_4_6.Tag
        ElseIf ColNum = 7 Then
            getCellTag = datecell_4_7.Tag
        End If
    ElseIf RowNum = 5 Then
        If ColNum = 1 Then
            getCellTag = datecell_5_1.Tag
        ElseIf ColNum = 2 Then
            getCellTag = datecell_5_2.Tag
        ElseIf ColNum = 3 Then
            getCellTag = datecell_5_3.Tag
        ElseIf ColNum = 4 Then
            getCellTag = datecell_5_4.Tag
        ElseIf ColNum = 5 Then
            getCellTag = datecell_5_5.Tag
        ElseIf ColNum = 6 Then
            getCellTag = datecell_5_6.Tag
        ElseIf ColNum = 7 Then
            getCellTag = datecell_5_7.Tag
        End If
    ElseIf RowNum = 6 Then
        If ColNum = 1 Then
            getCellTag = datecell_6_1.Tag
        ElseIf ColNum = 2 Then
            getCellTag = datecell_6_2.Tag
        ElseIf ColNum = 3 Then
            getCellTag = datecell_6_3.Tag
        ElseIf ColNum = 4 Then
            getCellTag = datecell_6_4.Tag
        ElseIf ColNum = 5 Then
            getCellTag = datecell_6_5.Tag
        ElseIf ColNum = 6 Then
            getCellTag = datecell_6_6.Tag
        ElseIf ColNum = 7 Then
            getCellTag = datecell_6_7.Tag
        End If
    End If

    
End Function

