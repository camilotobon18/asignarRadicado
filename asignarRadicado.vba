Sub asignar_radicado()
Dim listaFuncionarios() As String
Dim largoFuncionarios As Integer
Dim i As Integer
i = 1
Dim j As Integer
j = 1
Dim k As Integer
k = 1
Dim nombre As String
largoFuncionarios = Application.CountA(Worksheets("RADICADOS").Columns("E")) - 1
ReDim listaFuncionarios(1, largoFuncionarios)

For fila = 1 To largoFuncionarios
    listaFuncionarios(1, fila) = Worksheets("RADICADOS").Cells(fila + 1, 5).Value
Next fila




Do While (Worksheets("RADICADOS").Cells(i + 1, 1).Value <> "")
    
    If j > largoFuncionarios Then
        j = 1
    End If
    Worksheets("RADICADOS").Cells(i + 1, 2) = listaFuncionarios(1, j)
    i = i + 1
    j = j + 1
Loop


Do While (Worksheets("RADICADOS").Cells(k + 1, 5).Value <> "")
    
    nombre = Worksheets("RADICADOS").Cells(k + 1, 5).Value
    Worksheets("RADICADOS").Cells(k + 1, 6) = Application.CountIf(Worksheets("RADICADOS").Range("B2:B" & i), nombre)
    k = k + 1
Loop
End Sub
