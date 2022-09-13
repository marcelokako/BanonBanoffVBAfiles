Attribute VB_Name = "Trad"

Public Function Tradu(ByVal ped As String) As String

alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
resposta = ""
For ite1 = 2 To Len(ped) Step 2
    For ite2 = 1 To 17 'MESMO VALOR DA VARIAVEL qtd_lin EM PEDIDOS
        If InStr(Mid(ped, ite1, 1), Mid(alfabeto, ite2, 1)) Then
        resposta = resposta & " | " & Mid(ped, ite1 - 1, 1) & "x" & Sheets("Cardápio").Columns(1).Cells.Find(Mid(alfabeto, ite2, 1)).Offset(0, 1)
        End If
    Next
Next

Tradu = resposta

End Function
