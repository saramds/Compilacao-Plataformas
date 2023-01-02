Attribute VB_Name = "M�dulo1"
Sub compilacao_plataformas()

'Recursos de otimiza��o da macro

Application.ScreenUpdating = False

Application.Calculation = xlCalculationManual

'Limpa os dados iniciais

For Each aba In ThisWorkbook.Sheets

    If aba.Index > 1 Then

        aba.Range("B2:H1048576").ClearContents
    
    End If

Next

Sheets("Base").Activate

linha = 2

'Captura as informa��es do m�s, plataforma e volume da aba 'Base'

Do Until Cells(linha, 1).Value = ""

    mes = Cells(linha, 1).Value
    
    plataforma = Cells(linha, 3).Value
    
    volume_extraido = Cells(linha, 4).Value

    'Descobre qual coluna vai ser preenchida, de acordo com a plataforma
    
    coluna_plataforma = Sheets(mes).Cells.Find(plataforma).Column
    
    'Descobre �ltima linha vazia, que ser� preenchida
    
    linha_plataforma = Sheets(mes).Cells(1048576, coluna_plataforma).End(xlUp).Row + 1
    
    'Escreve a informa��o do volume extra�do
    
    Sheets(mes).Cells(linha_plataforma, coluna_plataforma).Value = volume_extraido
    
    linha = linha + 1

Loop

Application.Calculation = xlCalculationAutomatic

Application.ScreenUpdating = True

'Avisa que a macro foi executada com sucesso

resposta = MsgBox("Macro executada com sucesso!", vbInformation, "Sucesso!")

End Sub
