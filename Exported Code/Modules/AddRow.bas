Attribute VB_Name = "AddRow"
Sub Add_Row_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    On Error Resume Next
        ReturnSelect = Selection.Address
        If Err.Number <> 0 Then
            ReturnSelect = "B4"
        End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    With ActiveSheet.Shapes.Range(Array("Add_Row_Button"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    Call addRow

    Range(ReturnSelect).Select

    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Add_Row_Button"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .ThreeD.BevelTopDepth = 0.5
        .IncrementTop -1.2
    End With

    

End Sub
Function addRow()
    ' Get # of rows
    Dim RowCount As Integer
    RowCount = f.getRowCount()
    Range(RowCount & ":" & RowCount).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    With Range("B" & RowCount & ":F" & RowCount)
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = t.getP2Color
            .Weight = xlThin
        End With
        .Borders(xlEdgeTop).LineStyle = xlNone
    End With
    
    With Range("B" & RowCount).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = t.getP2Color
        .Weight = xlThin
    End With
    With Range("C" & RowCount).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = t.getP2Color
        .Weight = xlThin
    End With
    With Range("D" & RowCount).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = t.getP2Color
        .Weight = xlThin
    End With
    With Range("E" & RowCount).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = t.getP2Color
        .Weight = xlThin
    End With
    With Range("F" & RowCount).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = t.getP2Color
        .Weight = xlThin
    End With
    With Range("F" & RowCount).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = t.getP2Color
        .Weight = xlThin
    End With
    
    PeriodSheets.UpdateValidation
    
    positionAddRowButton
    
End Function
Function positionAddRowButton()
    Dim RowCount As Integer
    RowCount = f.getRowCount()
    
    With ActiveSheet.Shapes.Range(Array("Add_Row_Button"))
        .Top = ActiveSheet.Range("H" & RowCount - 2).Top + 3.5
        .Left = ActiveSheet.Range("H" & RowCount - 2).Left + 2.5
        .Height = 25
    End With
End Function






