Attribute VB_Name = "Módulo3"
Sub Financeiro_Clique()

periodoInicio = Sheets("Financeiro").Cells(6, 2)
periodoFinal = Sheets("Financeiro").Cells(7, 2)
    
countTaxaIfood = 0
countTaxaCielo = 0
countTaxaFora = 0
countBrutoIfood = 0
countBrutoCielo = 0
countBrutoFora = 0
countLiquidoIfood = 0
countLiquidoCielo = 0
countLiquidoFora = 0
countFrete = 0

linhaInicio = 2
    Do Until Sheets("Entregas").Cells(linhaInicio, 7) = periodoInicio
    linhaInicio = linhaInicio + 1
    Loop
    
    Do Until Sheets("Entregas").Cells(linhaInicio, 7) = periodoFinal Or Sheets("Entregas").Cells(linhaInicio, 7) = ""
        freteTemp = Sheets("Entregas").Cells(linhaInicio, 3)
        plataformaTemp = Sheets("Entregas").Cells(linhaInicio, 5)
        precoTemp = Sheets("Entregas").Cells(linhaInicio, 6)
        pagamentoTemp = Sheets("Entregas").Cells(linhaInicio, 9)
                    
        Select Case pagamentoTemp
                    
            Case "Crédito Online"
                colunaFinanceiro = 2
            Case "Débito Online"
                colunaFinanceiro = 3
            Case "Pix"
                colunaFinanceiro = 4
            Case "Maquineta Crédito"
                colunaFinanceiro = 5
            Case "Maquineta Débito"
                colunaFinanceiro = 6
            Case "Dinheiro"
                colunaFinanceiro = 7
            Case Else
                colunaFinanceiro = 2
        End Select
        Select Case plataformaTemp
            Case "Ifood"
                linhaFinanceiro = 2
            Case Else
                linhaFinanceiro = 3
        End Select
    
    taxaTemp = Sheets("Financeiro").Cells(linhaFinanceiro, colunaFinanceiro)
    
    precoFinal = precoTemp * (1 - taxaTemp)
    
        Select Case plataformaTemp
            Case "Ifood"
                countBrutoIfood = countBrutoIfood + precoTemp
                countLiquidoIfood = countLiquidoIfood + precoFinal
                countTaxaIfood = countTaxaIfood + (precoTemp - precoFinal)
            Case Else
                Select Case pagamentoTemp
                    Case "Dinheiro", "Pix"
                        countBrutoFora = countBrutoFora + precoTemp
                        countLiquidoFora = countLiquidoFora + precoFinal
                        countTaxaFora = countTaxaFora + (precoTemp - precoFinal)
                    Case Else
                        countBrutoCielo = countBrutoCielo + precoTemp
                        countLiquidoCielo = countLiquidoCielo + precoFinal
                        countTaxaCielo = countTaxaCielo + (precoTemp - precoFinal)
                End Select
        End Select
    
    linhaInicio = linhaInicio + 1
    
    countFrete = countFrete + freteTemp
    
    Loop

Sheets("Financeiro").Cells(6, 5) = countBrutoIfood
Sheets("Financeiro").Cells(6, 7) = countLiquidoIfood
Sheets("Financeiro").Cells(6, 6) = countTaxaIfood
Sheets("Financeiro").Cells(7, 5) = countBrutoCielo
Sheets("Financeiro").Cells(7, 7) = countLiquidoCielo
Sheets("Financeiro").Cells(7, 6) = countTaxaCielo
Sheets("Financeiro").Cells(8, 5) = countBrutoFora
Sheets("Financeiro").Cells(8, 7) = countLiquidoFora
Sheets("Financeiro").Cells(8, 6) = countTaxaFora
Sheets("Financeiro").Cells(6, 8) = countFrete

End Sub


