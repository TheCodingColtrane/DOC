VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDisparoEmail 
   Caption         =   "Disparo de E-mails"
   ClientHeight    =   11115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17955
   OleObjectBlob   =   "FrmDisparoEmail.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDisparoEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cbxs As New Collection
Private cbxCaptions As New Dictionary
Private AuxAnalista As New Dictionary
Private AnalistasSelecionados, aux, QtdCbxs, EmailsDisponiveis, EmailsMarcados As Integer
Private AnalistaAtual, EmailAtual  As Variant
Private Nome, Email As String

Private Sub btnEnviarEmail_Click()
aux = 1
EmailsMarcados = 0
For AnalistasSelecionados = 1 To QtdCbxs
aux = AnalistasSelecionados - 1
If cbxs.Item(AnalistasSelecionados).Value = True Then
cbxs.Item(AnalistasSelecionados).Enabled = False
Nome = cbxs.Item(AnalistasSelecionados).Caption
For Each AnalistaAtual In EmailsParaVisualizacao
Email = AuxAnalista.Item(Nome)
If Email = AnalistaAtual.Para Then
EnvioEmail AnalistaAtual.Para, AnalistaAtual.Copia, AnalistaAtual.Assunto, AnalistaAtual.Mensagem, , , True, AnalistaAtual.Anexo
If AnalistaAtual.Anexo <> "" Then
Kill (AnalistaAtual.Anexo)
End If
AnalistaAtual.Para = ""
AnalistaAtual.Copia = ""
AnalistaAtual.Assunto = ""
AnalistaAtual.Mensagem = ""
AnalistaAtual.Anexo = ""
EmailsDisponiveis = EmailsDisponiveis - 1
EmailsMarcados = EmailsMarcados + 1
End If
Next AnalistaAtual
End If
Next AnalistasSelecionados
If EmailsDisponiveis = 0 Then Unload FrmDisparoEmail
End Sub

Private Sub btnVerEmails_Click()
aux = 1
EmailsMarcados = 0
For AnalistasSelecionados = 1 To QtdCbxs
aux = AnalistasSelecionados - 1
If cbxs.Item(AnalistasSelecionados).Value = True Then
cbxs.Item(AnalistasSelecionados).Enabled = False
Nome = cbxs.Item(AnalistasSelecionados).Caption
For Each AnalistaAtual In EmailsParaVisualizacao
Email = AuxAnalista.Item(Nome)
If Email = AnalistaAtual.Para Then
EnvioEmail AnalistaAtual.Para, AnalistaAtual.Copia, AnalistaAtual.Assunto, AnalistaAtual.Mensagem, , , False, AnalistaAtual.Anexo
If AnalistaAtual.Anexo <> "" Then
Kill (AnalistaAtual.Anexo)
End If
AnalistaAtual.Para = ""
AnalistaAtual.Copia = ""
AnalistaAtual.Assunto = ""
AnalistaAtual.Mensagem = ""
AnalistaAtual.Anexo = ""
EmailsDisponiveis = EmailsDisponiveis - 1
EmailsMarcados = EmailsMarcados + 1
End If
Next AnalistaAtual
End If
Next AnalistasSelecionados
If EmailsDisponiveis = 0 Then Unload FrmDisparoEmail
End Sub

Private Sub chkAlternar_Click()
If EmailsMarcados = QtdCbxs Then
For Each EmailAtual In cbxs
EmailAtual.Value = False
Next EmailAtual
EmailsMarcados = 0
chkAlternar.Caption = "Marcar todos os E-mails"
Else
For Each EmailAtual In cbxs
EmailAtual.Value = True
Next EmailAtual
chkAlternar.Caption = "Desmarcar todos E-mails"
EmailsMarcados = QtdCbxs
End If
End Sub

Private Sub UserForm_Initialize()

Dim estilo As Integer
estilo = 25
Dim FeriadoAtual, espaco As Integer
Dim Chk As Control
Dim ChkStyle As CheckBox
Dim Email As Variant
Dim QtdEmails, EmailAtual As Integer
QtdEmails = UBound(NomesEmailsCriados, 2)
For EmailAtual = 0 To QtdEmails

If EmailsCriados.Exists(NomesEmailsCriados(1, EmailAtual)) Then
espaco = espaco + 1
Set Chk = Me.Controls.Add("Forms.CheckBox.1", "chk" & NomesEmailsCriados(0, EmailAtual))
AuxAnalista.Add NomesEmailsCriados(0, EmailAtual), NomesEmailsCriados(1, EmailAtual)
With Chk
.Caption = NomesEmailsCriados(0, EmailAtual)
.Value = True
cbxs.Add Chk, Chk.Name
cbxCaptions.Add NomesEmailsCriados(0, EmailAtual), ""
.Top = 40 + espaco * estilo
.Left = 325
.Width = 800
End With
End If

Next EmailAtual
espaco = espaco + 1
Set Chk = Me.Controls.Add("Forms.CheckBox.1", "chk" & EmailsCriados.Keys(EmailsCriados.Count - 1))
AuxAnalista.Add EmailsCriados.Keys(EmailsCriados.Count - 1), EmailsCriados.Items(EmailsCriados.Count - 1)
With Chk
.Caption = EmailsCriados.Keys(EmailsCriados.Count - 1)
.Value = True
cbxs.Add Chk, Chk.Name
cbxCaptions.Add EmailsCriados.Keys(EmailsCriados.Count - 1), ""
.Top = 40 + espaco * estilo
.Left = 325
.Width = 800
End With
QtdCbxs = cbxs.Count
EmailsMarcados = QtdCbxs
EmailsDisponiveis = QtdCbxs

End Sub
