Attribute VB_Name = "Módulo4"
Sub verificarEnvios()
    ' Define a planilha com a aba "Envios 2022" como ativa
    Worksheets("Envios 2022").Activate
    
    ' Define a última linha preenchida na coluna C
    lastRow = Cells(Rows.Count, "C").End(xlUp).Row
    
    ' Loop pelas linhas da coluna C, iniciando na linha 2
    For i = 2 To lastRow
        
        ' Pega os 3 primeiros dígitos da célula na coluna C
        tresDigitos = Left(Cells(i, "C"), 3)
        
        ' Verifica se os 3 primeiros dígitos são 300 ou 150
        If tresDigitos = "300" Then
            ' Insere a string "POSTAL" na coluna T
            Cells(i, "T") = "POSTAL"
        ElseIf tresDigitos = "150" Then
            ' Insere a string "ACESSO" na coluna T
            Cells(i, "T") = "ACESSO"
        End If
        
    Next i
    
End Sub

