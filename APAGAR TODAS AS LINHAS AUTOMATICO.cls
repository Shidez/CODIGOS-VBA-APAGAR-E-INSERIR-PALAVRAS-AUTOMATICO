
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim lin As Integer

'lin = Range("A7").CurrentRegion.Rows.Count

lin = Range("A7:A30").End(xlDown).Row

For linha = 7 To lin

If Cells(linha, "A").Value = "" Then
        Cells(linha, "H").Value = ""
End If

If Cells(linha, "A").Value = "" Then
        Cells(linha, "I").Value = ""
End If

Next
End Sub



