VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthSelector 
   Caption         =   "UserForm4"
   ClientHeight    =   4092
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5895
   OleObjectBlob   =   "MonthSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonthSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function selectCell(MonthInt As Integer)
    ReturnValue.Caption = MonthInt & "/" & getCurrentYear()
    Me.Hide
End Function
Function getCurrentYear() As Integer
    getCurrentYear = CInt(labyear.Caption)
End Function

Private Sub UserForm_Initialize()
    labyear.Caption = Year(Date)
End Sub















Private Sub butIncrYear_Click()
    labyear.Caption = getCurrentYear + 1
End Sub
Private Sub butDecrYear_Click()
    labyear.Caption = getCurrentYear - 1
End Sub
Private Sub monthcell_1_Click()
    selectCell (1)
End Sub
Private Sub monthcell_2_Click()
    selectCell (2)
End Sub
Private Sub monthcell_3_Click()
    selectCell (3)
End Sub
Private Sub monthcell_4_Click()
    selectCell (4)
End Sub
Private Sub monthcell_5_Click()
    selectCell (5)
End Sub
Private Sub monthcell_6_Click()
    selectCell (6)
End Sub
Private Sub monthcell_7_Click()
    selectCell (7)
End Sub
Private Sub monthcell_8_Click()
    selectCell (8)
End Sub
Private Sub monthcell_9_Click()
    selectCell (9)
End Sub
Private Sub monthcell_10_Click()
    selectCell (10)
End Sub
Private Sub monthcell_11_Click()
    selectCell (11)
End Sub
Private Sub monthcell_12_Click()
    selectCell (12)
End Sub



