Attribute VB_Name = "Módulo1"
Sub compilacao_plataformas()



For Each aba In ThisWorkbook.Sheets
If aba.Index > 1 Then


aba.Activate

Range("B2:H10000").ClearContents
End If

Next

Sheets("Base").Activate

linha = 2

Do Until Cells(linha, 1).Value = ""

mes = Cells(linha, 1).Value
plataforma = Cells(linha, 3).Value
volume = Cells(linha, 4).Value

Sheets(mes).Activate


coluna_plataforma = Cells.Find(plataforma).Column
linha_plataforma = Cells(100000, coluna_plataforma).End(xlUp).Row + 1

Cells(linha_plataforma, coluna_plataforma).Value = volume

Sheets("Base").Activate

linha = linha + 1

Loop

resposta = MsgBox("Macro executada com sucesso!", vbInformation, "Sucesso!")

End Sub










Sub compilacao_plataformas_otimizado()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


For Each aba In ThisWorkbook.Sheets
If aba.Index > 1 Then




aba.Range("B2:H10000").ClearContents
End If

Next

Sheets("Base").Activate

linha = 2

Do Until Cells(linha, 1).Value = ""

mes = Cells(linha, 1).Value
plataforma = Cells(linha, 3).Value
volume = Cells(linha, 4).Value

coluna_plataforma = Sheets(mes).Cells.Find(plataforma).Column
linha_plataforma = Sheets(mes).Cells(100000, coluna_plataforma).End(xlUp).Row + 1

Sheets(mes).Cells(linha_plataforma, coluna_plataforma).Value = volume



linha = linha + 1

Loop

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

resposta = MsgBox("Macro executada com sucesso!", vbInformation, "Sucesso!")

End Sub

Sub limpa_abas()


For Each aba In ThisWorkbook.Sheets
If aba.Index > 1 Then


aba.Activate

Range("B2:H10000").ClearContents
End If

Next
Sheets("Base").Activate


End Sub
