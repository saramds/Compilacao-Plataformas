Attribute VB_Name = "Módulo2"
Sub limpar_abas()

For Each aba In ThisWorkbook.Sheets

'Limpa os dados das abas dos meses

If aba.Index > 1 Then

    aba.Activate
    
    Range("B2:H1048576").ClearContents
    
End If

Next

'Volta para a aba principal

Sheets("Base").Activate

End Sub

