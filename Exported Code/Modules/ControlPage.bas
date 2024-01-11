Attribute VB_Name = "ControlPage"
Function renderCat()
    Sheets("Control").Select
    Categories = f.getCatArray()
    
    Sheets("Control").Range("B5:B1048576").ClearContents
    Sheets("Control").Range("B5:B1048576").Interior.Color = t.getBGColor
    For Each Item In Sheets("Control").Range("B5:B1048576").Borders
        Item.LineStyle = xlNone
    Next
    Index = 0
    
    P1Color = t.getP1Color
    P1FontName = t.getP1FontName
    P1FontColor = t.getP1FontColor
    
    ControlRowCount = f.getRowCount("Control")
    For Each Item In Categories
    
        If Index > ControlRowCount - 10 - 5 Then
            Sheets("Control").Range(ControlRowCount & ":" & ControlRowCount).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        With Sheets("Control").Range("B" & Index + 5)
            .Font.Size = 14
            .Font.Color = P1FontColor
            .Font.Name = P1FontName
            .Interior.Color = P1Color
            .FormulaR1C1 = Item
        End With
        Index = Index + 1
        
    Next
    
    With Sheets("Control").Shapes.Range("Edit_Category_Button")
        .Top = Sheets("Control").Range("B" & Index + 5).Top + 4
        .Left = Sheets("Control").Range("B5").Left + 2
        .Width = Sheets("Control").Range("B5").Width - 4
    End With

    checkRowCount

End Function
Function renderAct()
    Sheets("Control").Select
    Accounts = f.getActArray()
    
    Sheets("Control").Range("D5:D1048576").ClearContents
    Sheets("Control").Range("D5:D1048576").Interior.Color = t.getBGColor
    For Each Item In Sheets("Control").Range("D5:D1048576").Borders
        Item.LineStyle = xlNone
    Next
    Index = 0
    
    P1Color = t.getP1Color
    P1FontName = t.getP1FontName
    P1FontColor = t.getP1FontColor
    
    ControlRowCount = f.getRowCount("Control")
    For Each Item In Accounts
    
        If Index > ControlRowCount - 10 - 5 Then
            Sheets("Control").Range(ControlRowCount & ":" & ControlRowCount).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    
        With Sheets("Control").Range("D" & Index + 5)
            .Font.Size = 14
            .Font.Color = P1FontColor
            .Font.Name = P1FontName
            .Interior.Color = P1Color
            .FormulaR1C1 = Item
        End With
        Index = Index + 1
        
    Next
    
    checkRowCount
    
    With Sheets("Control").Shapes.Range("Edit_Account_Button")
        .Top = Sheets("Control").Range("D" & Index + 5).Top + 4
        .Left = Sheets("Control").Range("D5").Left + 2
        .Width = Sheets("Control").Range("D5").Width - 4
    End With
    
End Function
Function checkRowCount()
    
    BottomPosCat = Sheets("Control").Range("B" & f.getCatCount + 4).Top _
                    + Sheets("Control").Shapes.Range("Edit_Category_Button").Height + 4
    BottomPosAct = Sheets("Control").Range("B" & f.getActCount + 4).Top _
                    + Sheets("Control").Shapes.Range("Edit_Account_Button").Height + 4
    If BottomPosCat > BottomPosAct Then
        BottomPos = BottomPosCat
    Else
        BottomPos = BottomPosAct
    End If
                
    LastCell = "B" & f.getRowCount
    CellHeight = Sheets("Control").Range(LastCell).Height
    SheetHeight = Sheets("Control").Range(LastCell).Top + CellHeight
    
    
    ' Add Rows as needed
    If BottomPos > SheetHeight - (CellHeight * 2) Then
        RowsToAdd = Int((-BottomPos + SheetHeight) / CellHeight)
        For i = 1 To RowsToAdd Step 1
            Sheets("Control").Range(f.getRowCount & ":" & f.getRowCount).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
    End If
    
    ' Remove Rows As Needed
    If BottomPos + (CellHeight * 2) < SheetHeight And f.getRowCount > 22 Then
        RowsToRemove = Int((SheetHeight - (BottomPos + (CellHeight * 2))) / CellHeight)
        For i = 1 To RowsToRemove Step 1
            If f.getRowCount > 22 Then
                Sheets("Control").Rows(f.getRowCount).EntireRow.Hidden = True
            End If
            
        Next
    End If
    
End Function


