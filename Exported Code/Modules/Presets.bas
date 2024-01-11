Attribute VB_Name = "Presets"
Sub Theme1_Select_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Theme Presets").Select
    With ActiveSheet.Shapes.Range(Array("Select_Theme_1"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ' -------------------------------------
    With Sheets("Control").Range("G5:G7")
        .Interior.Color = 16777215
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G8:G10")
        .Interior.Color = 15917529
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G11:G13")
        .Interior.Color = 15189684
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G14:G16")
        .Interior.Color = 14395790
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G17:G19")
        .Interior.Color = 9851952
        .Font.Name = "Tw Cen MT"
        .Font.Color = 16777215
    End With
    ' -------------------------------------

    Sheets("Theme Presets").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Select_Theme_1"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    
    Sheets("Control").Select
    Range("F3:G4").Select
    Sheets("Theme Presets").Visible = False

End Sub
Sub Theme2_Select_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Theme Presets").Select
    With ActiveSheet.Shapes.Range(Array("Select_Theme_2"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ' -------------------------------------
    With Sheets("Control").Range("G5:G7")
        .Interior.Color = 16247773
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G8:G10")
        .Interior.Color = 15652797
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G11:G13")
        .Interior.Color = 15123099
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G14:G16")
        .Interior.Color = 11892015
        .Font.Name = "Tw Cen MT"
        .Font.Color = 16777215
    End With
    
    With Sheets("Control").Range("G17:G19")
        .Interior.Color = 14348258
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    ' -------------------------------------

    Sheets("Theme Presets").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Select_Theme_2"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    
    Sheets("Control").Select
    Range("F3:G4").Select
    Sheets("Theme Presets").Visible = False

End Sub
Sub Theme3_Select_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Theme Presets").Select
    With ActiveSheet.Shapes.Range(Array("Select_Theme_3"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ' -------------------------------------
    With Sheets("Control").Range("G5:G7")
        .Interior.Color = 16762606
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10165607
    End With
    
    With Sheets("Control").Range("G8:G10")
        .Interior.Color = 15109342
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10165607
    End With
    
    With Sheets("Control").Range("G11:G13")
        .Interior.Color = 16480169
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10165607
    End With
    
    With Sheets("Control").Range("G14:G16")
        .Interior.Color = 16404628
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10165607
    End With
    
    With Sheets("Control").Range("G17:G19")
        .Interior.Color = 10165607
        .Font.Name = "Tw Cen MT"
        .Font.Color = 16103651
    End With
    ' -------------------------------------

    Sheets("Theme Presets").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Select_Theme_3"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    
    Sheets("Control").Select
    Range("F3:G4").Select
    Sheets("Theme Presets").Visible = False

End Sub
Sub Theme4_Select_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Theme Presets").Select
    With ActiveSheet.Shapes.Range(Array("Select_Theme_4"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ' -------------------------------------
    With Sheets("Control").Range("G5:G7")
        .Interior.Color = 15663067
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10178333
    End With
    
    With Sheets("Control").Range("G8:G10")
        .Interior.Color = 14411441
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10178333
    End With
    
    With Sheets("Control").Range("G11:G13")
        .Interior.Color = 16499319
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10178333
    End With
    
    With Sheets("Control").Range("G14:G16")
        .Interior.Color = 16422992
        .Font.Name = "Tw Cen MT"
        .Font.Color = 10178333
    End With
    
    With Sheets("Control").Range("G17:G19")
        .Interior.Color = 10178333
        .Font.Name = "Tw Cen MT"
        .Font.Color = 15663067
    End With
    ' -------------------------------------

    Sheets("Theme Presets").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Select_Theme_4"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    
    Sheets("Control").Select
    Range("F3:G4").Select
    Sheets("Theme Presets").Visible = False

End Sub
Sub Theme5_Select_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Theme Presets").Select
    With ActiveSheet.Shapes.Range(Array("Select_Theme_5"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ' -------------------------------------
    With Sheets("Control").Range("G5:G7")
        .Interior.Color = 14022644
        .Font.Name = "Tw Cen MT"
        .Font.Color = 1925531
    End With
    
    With Sheets("Control").Range("G8:G10")
        .Interior.Color = 11657446
        .Font.Name = "Tw Cen MT"
        .Font.Color = 1925531
    End With
    
    With Sheets("Control").Range("G11:G13")
        .Interior.Color = 8174302
        .Font.Name = "Tw Cen MT"
        .Font.Color = 1925531
    End With
    
    With Sheets("Control").Range("G14:G16")
        .Interior.Color = 4821222
        .Font.Name = "Tw Cen MT"
        .Font.Color = 1925531
    End With
    
    With Sheets("Control").Range("G17:G19")
        .Interior.Color = 1925531
        .Font.Name = "Tw Cen MT"
        .Font.Color = 14417916
    End With
    ' -------------------------------------

    Sheets("Theme Presets").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Select_Theme_5"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    
    Sheets("Control").Select
    Range("F3:G4").Select
    Sheets("Theme Presets").Visible = False

End Sub
Sub Theme6_Select_Button()
'
'
' DO NOT TOUCH unless you know what you're doing.
'
    Application.ScreenUpdating = False
    Sheets("Theme Presets").Select
    With ActiveSheet.Shapes.Range(Array("Select_Theme_6"))
        .ThreeD.BevelTopInset = 0
        .ThreeD.BevelTopDepth = 0
        .IncrementTop 1.2
        With .Shadow
            .OffsetX = 0
            .OffsetY = 0
        End With
    End With
    Call f.forceScreenUpdate
    
    ' -------------------------------------
    With Sheets("Control").Range("G5:G7")
        .Interior.Color = 7628499
        .Font.Name = "Tw Cen MT"
        .Font.Color = 16777215
    End With
    
    With Sheets("Control").Range("G8:G10")
        .Interior.Color = 8953855
        .Font.Name = "Tw Cen MT"
        .Font.Color = 2104731
    End With
    
    With Sheets("Control").Range("G11:G13")
        .Interior.Color = 7829499
        .Font.Name = "Tw Cen MT"
        .Font.Color = 2104731
    End With
    
    With Sheets("Control").Range("G14:G16")
        .Interior.Color = 5263610
        .Font.Name = "Tw Cen MT"
        .Font.Color = 0
    End With
    
    With Sheets("Control").Range("G17:G19")
        .Interior.Color = 2104731
        .Font.Name = "Tw Cen MT"
        .Font.Color = 13823222
    End With
    ' -------------------------------------

    Sheets("Theme Presets").Select
    Application.ScreenUpdating = True
    With ActiveSheet.Shapes.Range(Array("Select_Theme_6"))
        With .Shadow
            .OffsetX = 1.2246467991E-16
            .OffsetY = 2
        End With
        .ThreeD.BevelTopInset = 1
        .IncrementTop -1.2
        .ThreeD.BevelTopDepth = 0.5
    End With
    
    Sheets("Control").Select
    Range("F3:G4").Select
    Sheets("Theme Presets").Visible = False

End Sub

Sub Exit_Button()
    Sheets("Control").Select
    Sheets("Theme Presets").Visible = False
End Sub


