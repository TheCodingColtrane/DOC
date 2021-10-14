VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmEscolhaFeriados 
   Caption         =   "Feriados"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   OleObjectBlob   =   "FrmEscolhaFeriados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEscolhaFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cbxs As New Collection
Private cbxCaptions As New Dictionary



Private Sub btnEscolherFeriados_Click()
Dim FeriadosSelecionados, aux As Integer
Dim FeriadosTrabalhados As New Dictionary
Dim QtdCbxs As Integer: QtdCbxs = cbxs.Count
aux = 0
For FeriadosSelecionados = 1 To QtdCbxs
If cbxs.Item(FeriadosSelecionados) = True Then
FeriadosTrabalhados.Add cbxCaptions.Keys(aux), cbxCaptions.Items(aux)
aux = aux + 1
End If
Next FeriadosSelecionados
Set SelecaoFeriados = FeriadosTrabalhados
frmEscolhaFeriados.Hide
End Sub


Private Sub UserForm_Initialize()
'Dim Feriados As New Dictionary: Set Feriados = GetFeriados()
'Set Feriados = FeriadoContextual(Feriados, Planilha_Aberta_Editar)
Dim Feriados As New Dictionary
Set Feriados = SelecaoFeriados

Dim estilo As Integer: estilo = 20
Dim FeriadoAtual, QtdFeriadoHomonimo As Integer
Dim QtdFeriados As Integer: QtdFeriados = Feriados.Count - 1
Dim Chk As Control
Dim ChkStyle As CheckBox
QtdFeriadoHomonimo = FeriadosHomonimos.Count
For FeriadoAtual = 0 To QtdFeriados
Set Chk = Me.Controls.Add("Forms.CheckBox.1", "chk" & FeriadoAtual)
With Chk
If QtdFeriadoHomonimo > 0 Then
If FeriadosHomonimos.Keys(FeriadoAtual) = Feriados.Keys(FeriadoAtual) Then
.Caption = CDate(Feriados.Keys(FeriadoAtual)) & " - " & FeriadosHomonimos.Items(FeriadoAtual)
Else
.Caption = CDate(Feriados.Keys(FeriadoAtual)) & " - " & Feriados.Items(FeriadoAtual)
End If
Else
.Caption = CDate(Feriados.Keys(FeriadoAtual)) & " - " & Feriados.Items(FeriadoAtual)
End If
.Value = True
cbxs.Add Chk, Chk.Name

If QtdFeriadoHomonimo > 0 Then
If FeriadosHomonimos.Keys(FeriadoAtual) = Feriados.Keys(FeriadoAtual) Then
cbxCaptions.Add CDate(Feriados.Keys(FeriadoAtual)), FeriadosHomonimos.Items(FeriadoAtual)
Else
cbxCaptions.Add CDate(Feriados.Keys(FeriadoAtual)), Feriados.Items(FeriadoAtual)
End If
Else
cbxCaptions.Add CDate(Feriados.Keys(FeriadoAtual)), Feriados.Items(FeriadoAtual)
End If
.Top = 40 + FeriadoAtual * estilo
.Left = 90
.Width = 150
End With
QtdFeriadoHomonimo = QtdFeriadoHomonimo - 1
Next FeriadoAtual

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
Dim resposta As Integer
resposta = MsgBox("Ao fechar esta janela, sem clicar no botão, todos os feriados " _
& "serão considerados para o trabalho. Deseja prosseguir?", vbYesNo + vbExclamation)
If resposta = vbYes Then
Unload frmEscolhaFeriados
Exit Sub
Else
Cancel = True
End If
Else
Cancel = False
End If
End Sub
