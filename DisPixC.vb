Option Explicit

Sub DisPixC()
    Dim i, iRow, iCol, var(256, 256, 3) As Integer
    i = 1
    Sheets("In").Select
    For iRow = 1 To 256
        For iCol = 1 To 256
            var(iRow, iCol, 1) = Range("A" & i).Value   ' Red
            var(iRow, iCol, 2) = Range("B" & i).Value   ' Green
            var(iRow, iCol, 3) = Range("C" & i).Value   ' Blue
            i = i + 1
        Next iCol
    Next iRow
    i = 1
    Sheets("Out").Select
    Range("A1:IV256").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    For iRow = 1 To 256
        For iCol = 1 To 256
            Range(Cells(iRow, iCol), Cells(iRow, iCol)).Select
            Selection.Interior.Color = RGB(var(iRow, iCol, 1), var(iRow, iCol, 2), var(iRow, iCol, 3))
            i = i + 1
        Next iCol
    Next iRow
End Sub
