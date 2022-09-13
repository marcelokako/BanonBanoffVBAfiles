VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pedido 
   Caption         =   "Registrar Pedido"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15270
   OleObjectBlob   =   "Pedido.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_SpinUp()
T_A.Value = T_A.Value + 1
End Sub
Private Sub B_SpinUp()
T_B.Value = T_B.Value + 1
End Sub
Private Sub C_SpinUp()
T_C.Value = T_C.Value + 1
End Sub
Private Sub D_SpinUp()
T_D.Value = T_D.Value + 1
End Sub
Private Sub E_SpinUp()
T_E.Value = T_E.Value + 1
End Sub
Private Sub F_SpinUp()
T_F.Value = T_F.Value + 1
End Sub
Private Sub G_SpinUp()
T_G.Value = T_G.Value + 1
End Sub
Private Sub H_SpinUp()
T_H.Value = T_H.Value + 1
End Sub
Private Sub I_SpinUp()
T_I.Value = T_I.Value + 1
End Sub
Private Sub J_SpinUp()
T_J.Value = T_J.Value + 1
End Sub
Private Sub K_SpinUp()
T_K.Value = T_K.Value + 1
End Sub
Private Sub L_SpinUp()
T_L.Value = T_L.Value + 1
End Sub
Private Sub M_SpinUp()
T_M.Value = T_M.Value + 1
End Sub
Private Sub N_SpinUp()
T_N.Value = T_N.Value + 1
End Sub
Private Sub O_SpinUp()
T_O.Value = T_O.Value + 1
End Sub



Private Sub A_SpinDown()
If Not T_A.Value = 0 Then
T_A.Value = T_A.Value - 1
End If
End Sub
Private Sub B_SpinDown()
If Not T_B.Value = 0 Then
T_B.Value = T_B.Value - 1
End If
End Sub
Private Sub C_SpinDown()
If Not T_C.Value = 0 Then
T_C.Value = T_C.Value - 1
End If
End Sub
Private Sub D_SpinDown()
If Not T_D.Value = 0 Then
T_D.Value = T_D.Value - 1
End If
End Sub
Private Sub E_SpinDown()
If Not T_E.Value = 0 Then
T_E.Value = T_E.Value - 1
End If
End Sub
Private Sub F_SpinDown()
If Not T_F.Value = 0 Then
T_F.Value = T_F.Value - 1
End If
End Sub
Private Sub G_SpinDown()
If Not T_G.Value = 0 Then
T_G.Value = T_G.Value - 1
End If
End Sub
Private Sub H_SpinDown()
If Not T_H.Value = 0 Then
T_H.Value = T_H.Value - 1
End If
End Sub
Private Sub I_SpinDown()
If Not T_I.Value = 0 Then
T_I.Value = T_I.Value - 1
End If
End Sub
Private Sub J_SpinDown()
If Not T_J.Value = 0 Then
T_J.Value = T_J.Value - 1
End If
End Sub
Private Sub K_SpinDown()
If Not T_K.Value = 0 Then
T_K.Value = T_K.Value - 1
End If
End Sub
Private Sub L_SpinDown()
If Not T_L.Value = 0 Then
T_L.Value = T_L.Value - 1
End If
End Sub
Private Sub M_SpinDown()
If Not T_M.Value = 0 Then
T_M.Value = T_M.Value - 1
End If
End Sub
Private Sub N_SpinDown()
If Not T_N.Value = 0 Then
T_N.Value = T_N.Value - 1
End If
End Sub
Private Sub O_SpinDown()
If Not T_O.Value = 0 Then
T_O.Value = T_O.Value - 1
End If
End Sub
Private Sub CommandButton1_Click()

Pedidotexto = ""

If Not Plataforma.Value = "" Or Pagamento.Value = "" Then
'SE MUDAR QUANTIDADE DE ITENS MUDAR AQUI
Dim itens(17, 17) As Double

qtd_lin = 17
alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

itens(1, 1) = T_A.Value
itens(2, 2) = T_B.Value
itens(3, 3) = T_C.Value
itens(4, 4) = T_D.Value
itens(5, 5) = T_E.Value
itens(6, 6) = T_F.Value
itens(7, 7) = T_G.Value
itens(8, 8) = T_H.Value
itens(9, 9) = T_I.Value
itens(10, 10) = T_J.Value
itens(11, 11) = T_K.Value
itens(12, 12) = T_L.Value
itens(13, 13) = T_M.Value
itens(14, 14) = T_N.Value
itens(15, 15) = T_O.Value
itens(16, 16) = T_P.Value
itens(17, 17) = T_Q.Value

preco = 0
If Plataforma.Value = "Ifood" Then
    incremento = 2
Else
    incremento = 3
End If

For ite = 1 To qtd_lin
    If Not itens(ite, ite) = 0 Then
    Pedidotexto = Pedidotexto & itens(ite, ite) & Mid(alfabeto, ite, 1)
    preco = preco + itens(ite, ite) * Sheets("Cardápio").Columns(1).Cells.Find(Mid(alfabeto, ite, 1)).Offset(0, incremento)
    End If
Next
Else
MsgBox "Preencha todos os dados"
End If


Select Case botao
    Case 1
    UserForm1.TextBox1.Value = Pedidotexto
    Plataforma1 = Plataforma.Value
    Pagamento1 = Pagamento.Value
    Preco1 = preco
    Pedido.Hide
    Case 2
    UserForm1.TextBox2.Value = Pedidotexto
    Plataforma2 = Plataforma.Value
    Pagamento2 = Pagamento.Value
    Preco2 = preco
    Pedido.Hide
    Case 3
    UserForm1.TextBox3.Value = Pedidotexto
    Plataforma3 = Plataforma.Value
    Pagamento3 = Pagamento.Value
    Preco3 = preco
    Pedido.Hide
    Case 4
    UserForm1.TextBox4.Value = Pedidotexto
    Plataforma4 = Plataforma.Value
    Pagamento4 = Pagamento.Value
    Preco4 = preco
    Pedido.Hide
End Select



End Sub

Private Sub T_A_Change()
If T_A.Value > 0 Then
T_A.BackColor = RGB(0, 210, 0)
Else
T_A.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_B_Change()
If T_B.Value > 0 Then
T_B.BackColor = RGB(0, 210, 0)
Else
T_B.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_C_Change()
If T_C.Value > 0 Then
T_C.BackColor = RGB(0, 210, 0)
Else
T_C.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_D_Change()
If T_D.Value > 0 Then
T_D.BackColor = RGB(0, 210, 0)
Else
T_D.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_E_Change()
If T_E.Value > 0 Then
T_E.BackColor = RGB(0, 210, 0)
Else
T_E.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_F_Change()
If T_F.Value > 0 Then
T_F.BackColor = RGB(0, 210, 0)
Else
T_F.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_G_Change()
If T_G.Value > 0 Then
T_G.BackColor = RGB(0, 210, 0)
Else
T_G.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_H_Change()
If T_H.Value > 0 Then
T_H.BackColor = RGB(0, 210, 0)
Else
T_H.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_I_Change()
If T_I.Value > 0 Then
T_I.BackColor = RGB(0, 210, 0)
Else
T_I.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_J_Change()
If T_J.Value > 0 Then
T_J.BackColor = RGB(0, 210, 0)
Else
T_J.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_K_Change()
If T_K.Value > 0 Then
T_K.BackColor = RGB(0, 210, 0)
Else
T_K.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_L_Change()
If T_L.Value > 0 Then
T_L.BackColor = RGB(0, 210, 0)
Else
T_L.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_M_Change()
If T_M.Value > 0 Then
T_M.BackColor = RGB(0, 210, 0)
Else
T_M.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_N_Change()
If T_N.Value > 0 Then
T_N.BackColor = RGB(0, 210, 0)
Else
T_N.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_O_Change()
If T_O.Value > 0 Then
T_O.BackColor = RGB(0, 210, 0)
Else
T_O.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_P_Change()
If T_P.Value > 0 Then
T_P.BackColor = RGB(0, 210, 0)
Else
T_P.BackColor = RGB(250, 250, 250)
End If
End Sub
Private Sub T_Q_Change()
If T_Q.Value > 0 Then
T_Q.BackColor = RGB(0, 210, 0)
Else
T_Q.BackColor = RGB(250, 250, 250)
End If
End Sub

Private Sub UserForm_Initialize()

Plataforma.AddItem "Ifood"
Plataforma.AddItem "Neemo"
Plataforma.AddItem "WhatsApp"
Plataforma.AddItem "Outro"

Pagamento.AddItem "Pix"
Pagamento.AddItem "Crédito Online"
Pagamento.AddItem "Débito Online"
Pagamento.AddItem "Maquineta Crédito"
Pagamento.AddItem "Maquineta Débito"
Pagamento.AddItem "Dinheiro"
End Sub

