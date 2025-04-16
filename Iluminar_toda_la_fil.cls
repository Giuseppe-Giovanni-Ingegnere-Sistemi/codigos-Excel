Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells.Interior.ColorIndex = xlNone
    Target.EntireRow.Interior.Color = RGB(200, 230, 255)
End Sub
