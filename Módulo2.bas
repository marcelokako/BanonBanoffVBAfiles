Attribute VB_Name = "M�dulo2"
Sub Atualizar_Clique()
Sheets("Sum�rio").Cells(2, 11) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
Sheets("Sum�rio").Cells(2, 1).ListObject.DataBodyRange.Rows.ClearContents

'lin serve para procurar na plan entregas e linha para escrever na plan sum�rio
lin = 1
LinhaFim = 2
Do Until Sheets("Entregas").Cells(LinhaFim, 1) = ""
LinhaFim = LinhaFim + 1
Loop
Do Until Sheets("Entregas").Cells(LinhaFim - lin, 7) <> Sheets("Sum�rio").Cells(2, 11)

entregador = Sheets("Entregas").Cells(LinhaFim - lin, 1)

linha = 2
Do Until Sheets("Sum�rio").Cells(linha, 1) = ""
linha = linha + 1
Loop

For ite = 2 To linha
    If Sheets("Sum�rio").Cells(ite, 1) = entregador Then
    TaNoRange = True
    Exit For
    Else
    TaNoRange = False
    End If
Next ite
If TaNoRange Then
    Sheets("Sum�rio").Cells(ite, 5) = Cells(ite, 5) + 1
    Sheets("Sum�rio").Cells(ite, 6) = Cells(ite, 6) & "," & Sheets("Entregas").Cells(LinhaFim - lin, 1).Offset(0, 1)
    Sheets("Sum�rio").Cells(ite, 7) = Cells(ite, 7) + Sheets("Entregas").Cells(LinhaFim - lin, 1).Offset(0, 2)
Else
    'nao ta no range
    Sheets("Sum�rio").Cells(linha, 1) = entregador
    Sheets("Sum�rio").Cells(linha, 2) = Sheets("Motoboys").Cells.Find(entregador).Offset(0, 1)
    Sheets("Sum�rio").Cells(linha, 3) = Sheets("Motoboys").Cells.Find(entregador).Offset(0, 2)
    Sheets("Sum�rio").Cells(linha, 4) = Sheets("Motoboys").Cells.Find(entregador).Offset(0, 3)
    Sheets("Sum�rio").Cells(linha, 5) = 1
    Sheets("Sum�rio").Cells(linha, 6) = Sheets("Entregas").Cells.Find(entregador, searchdirection:=xlPrevious).Offset(0, 1)
    Sheets("Sum�rio").Cells(linha, 7) = Sheets("Entregas").Cells.Find(entregador, searchdirection:=xlPrevious).Offset(0, 2)
    Sheets("Sum�rio").Cells(linha, 8) = 0
End If

lin = lin + 1
Loop
End Sub


