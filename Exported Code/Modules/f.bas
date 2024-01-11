Attribute VB_Name = "f"
Sub forceScreenUpdate()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Wait Now + #12:00:01 AM#
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub
Function getCatCount() As Integer
    getCatCount = Application.WorksheetFunction.CountA(Sheets("Control").Range("B:B")) - 1
End Function
Function getCatArray() As String()
    CatCount = getCatCount()
    
    Dim StringValues
    StringValues = Sheets("Control").Range("B5").Value
    For i = 1 To CatCount - 1 Step 1
        StringValues = StringValues + "|!DELIM!|" + Sheets("Control").Range("B" & 5 + i).Value
    Next
    
    getCatArray = Split(StringValues, "|!DELIM!|")
End Function
Function getActCount() As Integer
    getActCount = Application.WorksheetFunction.CountA(Sheets("Control").Range("D:D")) - 1
End Function
Function getActArray() As String()
    ActCount = getActCount()
    
    Dim StringValues
    StringValues = Sheets("Control").Range("D5").Value
    For i = 1 To ActCount - 1 Step 1
        StringValues = StringValues + "|!DELIM!|" + Sheets("Control").Range("D" & 5 + i).Value
    Next
    
    getActArray = Split(StringValues, "|!DELIM!|")
End Function
Function getActStartBalArray() As String()
    ActCount = getActCount()
    StartSheet = f.getPerArray()(0)
    Dim StringValues As String
    StringValues = Sheets(StartSheet).Range("I4").Value
    If ActCount = 1 Then
        Dim RetArray(0) As String
        RetArray(0) = StringValues
        getActStartBalArray = RetArray
    Else
        For i = 2 To ActCount Step 1
            StringValues = StringValues & "|!DELIM!|" & Sheets(StartSheet).Range("I" & i + 3).Value
        Next
        getActStartBalArray = Split(StringValues, "|!DELIM!|")
    End If
End Function
Function getPerCount() As Integer
    getPerCount = Application.WorksheetFunction.CountA(Sheets("Overview").Range("2:2")) - 1

End Function
Function getPerArray() As String()
    PerCount = getPerCount()
    
    Dim StringValues
    StringValues = Sheets("Overview").Range("C2").Value
    For i = 1 To PerCount - 1 Step 1
        StringValues = StringValues & "|!DELIM!|" & Sheets("Overview").Range(numToLet(i + 3) & "2").Value
    Next
    
    getPerArray = Split(StringValues, "|!DELIM!|")

End Function
Function numToLet(ColumnNum As Integer) As String
    numToLet = Split(Cells(1, ColumnNum).Address, "$")(1)
End Function
Function getRowCount(Optional SheetName As String = "|!BLANK!|") As Integer
    If SheetName = "|!BLANK!|" Then
        For i = 1 To 10000 Step 1
            If Rows(i).Hidden = True Then
                getRowCount = i - 1
                Exit For
            Else
                getRowCount = -1
            End If
        Next
    Else
        For i = 1 To 10000 Step 1
            If Sheets(SheetName).Rows(i).Hidden = True Then
                getRowCount = i - 1
                Exit For
            Else
                getRowCount = -1
            End If
        Next
    End If
End Function

