Dim celdaAnterior As Range

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Restaurar color de la celda anterior
    If Not celdaAnterior Is Nothing Then
        celdaAnterior.Interior.ColorIndex = xlNone
    End If

    ' Aplicar nuevo color a la celda activa
    Target.Interior.Color = RGB(255, 255, 0) ' Amarillo

    ' Guardar nueva celda como la anterior
    Set celdaAnterior = Target
End Sub
