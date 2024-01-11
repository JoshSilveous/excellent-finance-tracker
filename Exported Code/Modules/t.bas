Attribute VB_Name = "t"
Function getBGColor() As Long
    getBGColor = toColorFormat(Sheets("Control").Range("G6").DisplayFormat.Interior.Color)
End Function
Function getBGFontColor() As Long
    getBGFontColor = toColorFormat(Sheets("Control").Range("G6").DisplayFormat.Font.Color)
End Function
Function getBGFontName() As String
    getBGFontName = Sheets("Control").Range("G6").DisplayFormat.Font.Name
End Function

Function getP1Color() As Long
    getP1Color = toColorFormat(Sheets("Control").Range("G9").DisplayFormat.Interior.Color)
End Function
Function getP1FontColor() As Long
    getP1FontColor = toColorFormat(Sheets("Control").Range("G9").DisplayFormat.Font.Color)
End Function
Function getP1FontName() As String
    getP1FontName = Sheets("Control").Range("G9").DisplayFormat.Font.Name
End Function

Function getP2Color() As Long
    getP2Color = toColorFormat(Sheets("Control").Range("G12").DisplayFormat.Interior.Color)
End Function
Function getP2FontColor() As Long
    getP2FontColor = toColorFormat(Sheets("Control").Range("G12").DisplayFormat.Font.Color)
End Function
Function getP2FontName() As String
    getP2FontName = Sheets("Control").Range("G12").DisplayFormat.Font.Name
End Function

Function getP3Color() As Long
    getP3Color = toColorFormat(Sheets("Control").Range("G15").DisplayFormat.Interior.Color)
End Function
Function getP3FontColor() As Long
    getP3FontColor = toColorFormat(Sheets("Control").Range("G15").DisplayFormat.Font.Color)
End Function
Function getP3FontName() As String
    getP3FontName = Sheets("Control").Range("G15").DisplayFormat.Font.Name
End Function

Function getBColor() As Long
    getBColor = toColorFormat(Sheets("Control").Range("G18").DisplayFormat.Interior.Color)
End Function
Function getBFontColor() As Long
    getBFontColor = toColorFormat(Sheets("Control").Range("G18").DisplayFormat.Font.Color)
End Function
Function getBFontName() As String
    getBFontName = Sheets("Control").Range("G18").DisplayFormat.Font.Name
End Function
Function toColorFormat(inputDouble As Double) As Long
    R = inputDouble Mod 256
    G = inputDouble \ 256 Mod 256
    B = inputDouble \ 65536 Mod 256
    toColorFormat = RGB(R, G, B)
End Function

