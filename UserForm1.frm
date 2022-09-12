VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Registrar Pedido"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click()

botao = 1
Pedido.T_A.Value = 0
Pedido.T_B.Value = 0
Pedido.T_C.Value = 0
Pedido.T_D.Value = 0
Pedido.T_E.Value = 0
Pedido.T_F.Value = 0
Pedido.T_G.Value = 0
Pedido.T_H.Value = 0
Pedido.T_I.Value = 0
Pedido.T_J.Value = 0
Pedido.T_K.Value = 0
Pedido.T_L.Value = 0
Pedido.T_M.Value = 0
Pedido.T_N.Value = 0
Pedido.T_O.Value = 0
Pedido.Show
End Sub

Private Sub Button2_Click()

botao = 2
Pedido.T_A.Value = 0
Pedido.T_B.Value = 0
Pedido.T_C.Value = 0
Pedido.T_D.Value = 0
Pedido.T_E.Value = 0
Pedido.T_F.Value = 0
Pedido.T_G.Value = 0
Pedido.T_H.Value = 0
Pedido.T_I.Value = 0
Pedido.T_J.Value = 0
Pedido.T_K.Value = 0
Pedido.T_L.Value = 0
Pedido.T_M.Value = 0
Pedido.T_N.Value = 0
Pedido.T_O.Value = 0
Pedido.Show
End Sub

Private Sub Button3_Click()

botao = 3
Pedido.T_A.Value = 0
Pedido.T_B.Value = 0
Pedido.T_C.Value = 0
Pedido.T_D.Value = 0
Pedido.T_E.Value = 0
Pedido.T_F.Value = 0
Pedido.T_G.Value = 0
Pedido.T_H.Value = 0
Pedido.T_I.Value = 0
Pedido.T_J.Value = 0
Pedido.T_K.Value = 0
Pedido.T_L.Value = 0
Pedido.T_M.Value = 0
Pedido.T_N.Value = 0
Pedido.T_O.Value = 0
Pedido.Show
End Sub

Private Sub Button4_Click()
botao = 4
Pedido.T_A.Value = 0
Pedido.T_B.Value = 0
Pedido.T_C.Value = 0
Pedido.T_D.Value = 0
Pedido.T_E.Value = 0
Pedido.T_F.Value = 0
Pedido.T_G.Value = 0
Pedido.T_H.Value = 0
Pedido.T_I.Value = 0
Pedido.T_J.Value = 0
Pedido.T_K.Value = 0
Pedido.T_L.Value = 0
Pedido.T_M.Value = 0
Pedido.T_N.Value = 0
Pedido.T_O.Value = 0
Pedido.Show
End Sub

Private Sub Finalizar_Click()

lin = 2
Do Until Sheets("Entregas").Cells(lin, 1) = ""
lin = lin + 1
Loop
'LinhaFim = lin + entregas.Value - 1
'1Motoboy   2Bairro  3Frete   4pedido  5Plataforma  6Preço   7Data    8Horário  9pagamento

If Not Retirada.Value = True Then

Motoboy = ComboBox1.Value
Bairro1 = ComboBox2.Value
frete1 = Sheets("Bairros").Cells.Find(Bairro1).Offset(0, 1)
Pedido1 = TextBox1.Value

'Escrevendo motoboy
Sheets("Entregas").Cells(lin, 1) = Motoboy
'Escrevendo bairro
Sheets("Entregas").Cells(lin, 2) = Bairro1
'Escrevendo Frete
Sheets("Entregas").Cells(lin, 3) = frete1
'Escrevendo Pedido
Sheets("Entregas").Cells(lin, 4) = Pedido1
'Escrevendo Plataforma
Sheets("Entregas").Cells(lin, 5) = Plataforma1
'Escrevendo Preço
Sheets("Entregas").Cells(lin, 6) = Preco1 + frete1
'Escrevendo Data
Sheets("Entregas").Cells(lin, 7) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
'Escrevendo Horário
Sheets("Entregas").Cells(lin, 8) = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
'Escrevendo Horário
Sheets("Entregas").Cells(lin, 9) = Pagamento1

    If entregas.Value >= 2 Then
    Bairro2 = ComboBox3.Value
    frete2 = Sheets("Bairros").Cells.Find(Bairro2).Offset(0, 1)
    Pedido2 = TextBox2.Value
    
    Sheets("Entregas").Cells(lin + 1, 1) = Motoboy
    Sheets("Entregas").Cells(lin + 1, 2) = Bairro2
    Sheets("Entregas").Cells(lin + 1, 3) = frete2
    Sheets("Entregas").Cells(lin + 1, 4) = Pedido2
    Sheets("Entregas").Cells(lin + 1, 5) = Plataforma2
    Sheets("Entregas").Cells(lin + 1, 6) = Preco2 + frete2
    Sheets("Entregas").Cells(lin + 1, 7) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    Sheets("Entregas").Cells(lin + 1, 8) = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
    Sheets("Entregas").Cells(lin + 1, 9) = Pagamento2
    End If
    If entregas.Value >= 3 Then
    Bairro3 = ComboBox4.Value
    frete3 = Sheets("Bairros").Cells.Find(Bairro3).Offset(0, 1)
    Pedido3 = TextBox3.Value
    
    Sheets("Entregas").Cells(lin + 2, 1) = Motoboy
    Sheets("Entregas").Cells(lin + 2, 2) = Bairro3
    Sheets("Entregas").Cells(lin + 2, 3) = frete3
    Sheets("Entregas").Cells(lin + 2, 4) = Pedido3
    Sheets("Entregas").Cells(lin + 2, 5) = Plataforma3
    Sheets("Entregas").Cells(lin + 2, 6) = Preco3 + frete3
    Sheets("Entregas").Cells(lin + 2, 7) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    Sheets("Entregas").Cells(lin + 2, 8) = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
    Sheets("Entregas").Cells(lin + 2, 9) = Pagamento3
    End If
    If entregas.Value = 4 Then
    Bairro4 = ComboBox5.Value
    frete4 = Sheets("Bairros").Cells.Find(Bairro4).Offset(0, 1)
    Pedido4 = TextBox4.Value
    
    Sheets("Entregas").Cells(lin + 3, 1) = Motoboy
    Sheets("Entregas").Cells(lin + 3, 2) = Bairro4
    Sheets("Entregas").Cells(lin + 3, 3) = frete4
    Sheets("Entregas").Cells(lin + 3, 4) = Pedido4
    Sheets("Entregas").Cells(lin + 3, 5) = Plataforma4
    Sheets("Entregas").Cells(lin + 3, 6) = Preco4 + frete4
    Sheets("Entregas").Cells(lin + 3, 7) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    Sheets("Entregas").Cells(lin + 3, 8) = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
    Sheets("Entregas").Cells(lin + 3, 9) = Pagamento4
    End If
       
    
Else

    Motoboy = "Retirada"
    Sheets("Entregas").Cells(lin, 1) = "Retirada"
    Sheets("Entregas").Cells(lin, 2) = "-"
    Sheets("Entregas").Cells(lin, 3) = 0
    Sheets("Entregas").Cells(lin, 4) = TextBox1.Value
    Sheets("Entregas").Cells(lin, 5) = Plataforma1
    Sheets("Entregas").Cells(lin, 6) = Preco1
    Sheets("Entregas").Cells(lin, 7) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    Sheets("Entregas").Cells(lin, 8) = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
    Sheets("Entregas").Cells(lin, 9) = Pagamento1
    
End If
'Atualizando Planilha Motoboys(data e ultima entrega)
EntregasTotais = Sheets("Motoboys").Cells.Find(Motoboy).Offset(0, 3)
Sheets("Motoboys").Cells.Find(Motoboy).Offset(0, 3) = EntregasTotais + entregas.Value
Sheets("Motoboys").Cells.Find(Motoboy).Offset(0, 4) = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
Call Atualizar_Clique
Me.Hide
End Sub

Private Sub entregas_change()

Select Case entregas.Value
    Case 1
    ComboBox3.Visible = False
    ComboBox4.Visible = False
    ComboBox5.Visible = False
    
    Button2.Visible = False
    Button3.Visible = False
    Button4.Visible = False
    
    TextBox2.Visible = False
    TextBox3.Visible = False
    TextBox4.Visible = False
    
    Case 2
    ComboBox3.Visible = True
    ComboBox4.Visible = False
    ComboBox5.Visible = False
    
    Button2.Visible = True
    Button3.Visible = False
    Button4.Visible = False
    
    TextBox2.Visible = True
    TextBox3.Visible = False
    TextBox4.Visible = False
    
    Case 3
    ComboBox3.Visible = True
    ComboBox4.Visible = True
    ComboBox5.Visible = False
    
    Button2.Visible = True
    Button3.Visible = True
    Button4.Visible = False
    
    TextBox2.Visible = True
    TextBox3.Visible = True
    TextBox4.Visible = False
    
    Case 4
    ComboBox3.Visible = True
    ComboBox4.Visible = True
    ComboBox5.Visible = True
    
    Button2.Visible = True
    Button3.Visible = True
    Button4.Visible = True
    
    TextBox2.Visible = True
    TextBox3.Visible = True
    TextBox4.Visible = True

End Select

End Sub

Private Sub SpinButton1_SpinDown()

qtd_entregas = entregas.Value

If Not qtd_entregas = 1 Then
 qtd_entregas = qtd_entregas - 1
entregas.Value = qtd_entregas
End If


End Sub
 
Private Sub SpinButton1_SpinUp()

qtd_entregas = entregas.Value

If Not qtd_entregas = 4 Then
 qtd_entregas = qtd_entregas + 1
entregas.Value = qtd_entregas

End If


End Sub

Private Sub Retirada_Change()

If Retirada.Value = True Then
ComboBox1.Enabled = False
ComboBox2.Enabled = False
SpinButton1.Enabled = False
entregas.Enabled = False
entregas.Value = 1
Else
ComboBox1.Enabled = True
ComboBox2.Enabled = True
SpinButton1.Enabled = True
entregas.Enabled = True

End If

End Sub

Private Sub ComboBox1_MouseMove( _
                        ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
                HookListBoxScroll Me, Me.ComboBox1
End Sub
Private Sub ComboBox2_MouseMove( _
                        ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
                HookListBoxScroll Me, Me.ComboBox2
End Sub
Private Sub ComboBox3_MouseMove( _
                        ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
                HookListBoxScroll Me, Me.ComboBox3
End Sub
Private Sub ComboBox4_MouseMove( _
                        ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
                HookListBoxScroll Me, Me.ComboBox4
End Sub
Private Sub ComboBox5_MouseMove( _
                        ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
                HookListBoxScroll Me, Me.ComboBox5
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        UnhookListBoxScroll
End Sub

