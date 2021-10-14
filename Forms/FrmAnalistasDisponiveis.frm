VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAnalistasDisponiveis 
   Caption         =   "Analistas Disponíveis"
   ClientHeight    =   12555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21360
   OleObjectBlob   =   "FrmAnalistasDisponiveis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAnalistasDisponiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Analistas As Variant
Private cbxs As New Collection
Private QtdAnalistas, QtdNomes, QtdCbxs As Integer
Private cbxCaptions As New Dictionary

Private Sub BtnSelecionarAnalistasDisponiveis_Click()

Dim AnalistasSelecionados, aux As Integer
Dim AnalistaDisponiveis As New Dictionary
Dim QtdCbxs As Integer: QtdCbxs = cbxs.Count
Dim Analista(), AnalistaSelecionado As Variant
ReDim Analista(1, 0)
aux = 0
For AnalistasSelecionados = 1 To QtdCbxs
aux = AnalistasSelecionados - 1
If cbxs.Item(AnalistasSelecionados) = True Then
AnalistaDisponiveis.Add cbxCaptions.Keys(aux), cbxCaptions.Items(aux)
aux = aux + 1
Else
aux = aux + 1
'AnalistaDisponiveis.RemoveAll
End If

Next AnalistasSelecionados
If aux < 3 Then
MsgBox "Não é permitido selecionar menos de 3 analistas", vbExclamation
Exit Sub
MsgBox cbxCaptions.Keys(0)
End If
QtdNomes = AnalistaDisponiveis.Count - 1
ReDim Analista(1, 0 To QtdNomes)
aux = 0
For Each AnalistaSelecionado In AnalistaDisponiveis.Keys
Analista(0, aux) = AnalistaSelecionado
Analista(1, aux) = AnalistaDisponiveis.Item(AnalistaSelecionado)
aux = aux + 1
Next AnalistaSelecionado
Dim Preferencias As CRAVPreferencias
Set Preferencias = New CRAVPreferencias
Preferencias.celula = CStr(RAVPreferencias(1).celula)
Preferencias.Analista = RAVPreferencias(1).Analista
Preferencias.Cliente = RAVPreferencias(1).Cliente
Preferencias.ambiente = RAVPreferencias(1).ambiente
Preferencias.Analistas = Analista
RAVPreferencias.Remove (1)
RAVPreferencias.Add Preferencias
Unload FrmAnalistasDisponiveis
End Sub

Private Sub ChkDesmarcarTodosOsAnalistas_Click()
If QtdNomes = QtdCbxs Then
Dim NomeAtual
For Each NomeAtual In cbxs
NomeAtual.Value = False
Next NomeAtual
QtdNomes = 0
ChkDesmarcarTodosOsAnalistas.Caption = "Marcar todos os analistas"
Else
For Each NomeAtual In cbxs
NomeAtual.Value = True
Next NomeAtual
ChkDesmarcarTodosOsAnalistas.Caption = "Desmarcar todos os analistas"
QtdNomes = QtdCbxs
End If
End Sub

Private Sub ScrollBar1_Change()

End Sub

Private Sub UserForm_Initialize()

Dim QtdDocumentosAtrasados, QtdDocumentosPrazoFatal, QtdDocumentoEmDia, _
PorcentagemBloqueioAtrasado, PorcentagemComumAtrasado, _
PorcentagemBloqueioPrazoFatal, PorcentagemComumPrazoFatal, PrimeiraLinha, TotalDocumentos As Long

Dim UltimaLinha As Long

Dim PlanilhaEspelho As Excel.Worksheet
Dim PlanilhaRav As Excel.Worksheet

Set RAVColunas = New cRAVColunasXL
Set RAVColunas = GetColunasIndiceAlfabetico(PastadeTrabalhoRAV)

Dim estilo As Integer
estilo = 25
Dim FeriadoAtual, espaco, botaoStyle As Integer
Dim Chk As Control
Dim ChkStyle As CheckBox
Dim Email As Variant
Dim QtdEmails, EmailAtual As Integer

Set PlanilhaRav = PastadeTrabalhoRAV.Worksheets(1)
Set PlanilhaEspelho = PastadeTrabalhoRAV.Worksheets(2)
PrimeiraLinha = 2
UltimaLinha = PlanilhaEspelho.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row

QtdDocumentosAtrasados = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Atrasado", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não")

QtdDocumentosPrazoFatal = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Prazo Fatal", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não")

QtdDocumentoEmDia = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Em dia", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não")

PorcentagemBloqueioAtrasado = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Tipo & PrimeiraLinha & ":" & RAVColunas.Tipo & UltimaLinha), "Bloqueio", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não", _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Atrasado")

If QtdDocumentosAtrasados > 0 Then
PorcentagemBloqueioAtrasado = _
PlanilhaEspelho.Application.WorksheetFunction.Round(PorcentagemBloqueioAtrasado / QtdDocumentosAtrasados, 2)
Else
PorcentagemBloqueioAtrasado = 0
End If
PorcentagemComumAtrasado = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Tipo & PrimeiraLinha & ":" & RAVColunas.Tipo & UltimaLinha), "Comum", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não", _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Atrasado")

If QtdDocumentosAtrasados > 0 Then
PorcentagemComumAtrasado = _
PlanilhaEspelho.Application.WorksheetFunction.Round(PorcentagemComumAtrasado / QtdDocumentosAtrasados, 2)
Else
PorcentagemComumAtrasado = 0
End If

PorcentagemBloqueioPrazoFatal = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Tipo & PrimeiraLinha & ":" & RAVColunas.Tipo & UltimaLinha), "Bloqueio", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não", _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Prazo Fatal")

If QtdDocumentosPrazoFatal > 0 Then
PorcentagemBloqueioPrazoFatal = _
PlanilhaEspelho.Application.WorksheetFunction.Round(PorcentagemBloqueioPrazoFatal / QtdDocumentosPrazoFatal, 2)
Else
PorcentagemBloqueioPrazoFatal = 0
End If

PorcentagemComumPrazoFatal = PlanilhaEspelho.Application.WorksheetFunction.CountIfs( _
PlanilhaEspelho.Range(RAVColunas.Tipo & PrimeiraLinha & ":" & RAVColunas.Tipo & UltimaLinha), "Comum", _
PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não", _
PlanilhaEspelho.Range(RAVColunas.Status & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha), "Prazo Fatal")

If QtdDocumentosPrazoFatal > 0 Then
PorcentagemComumPrazoFatal = _
PlanilhaEspelho.Application.WorksheetFunction.Round(PorcentagemComumPrazoFatal / QtdDocumentosPrazoFatal, 2)
Else
PorcentagemComumPrazoFatal = 0
End If

LblMsgSituacao = "Hoje você possui " & QtdDocumentosAtrasados & " documento(s) atrasado(s) " & QtdDocumentosPrazoFatal _
& " documento(s) em prazo fatal e " & QtdDocumentoEmDia & " documento(s) em dia"

LblMsgAtrasadosPorcentagem = "Sobre o(s) documento(s) atrasado(s) " & PorcentagemBloqueioAtrasado & "% é/são de bloqueio e " _
& PorcentagemComumAtrasado & " é/são comum(ns)."

LblMsgPrazoFatalPorcentagem = "Sobre o(s) documento(s) em prazo fatal " & PorcentagemBloqueioPrazoFatal & "% é/são de bloqueio e " _
& PorcentagemComumPrazoFatal & " é/são comuns."


estilo = 20
Dim AnalistaAtual, AnalistaHomonimo, QtdAnalista, QtdAnalistaHomonimo As Integer
Analistas = RAVPreferencias(1).Analistas
QtdAnalistas = UBound(Analistas, 2)
LblQtdAnalista = QtdAnalistas + 1 & " analistas disponíveis"
For AnalistaAtual = 0 To QtdAnalistas
Set Chk = Me.Controls.Add("Forms.CheckBox.1", "chk" & AnalistaAtual)
With Chk
If AnalistaAtual > 0 And AnalistaAtual < QtdAnalista Then
If Analistas(0, AnalistaAtual) = Analistas(0, AnalistaAtual + 1) Then
QtdAnalistaHomonimo = QtdAnalistaHomonimo + 1
MsgBox "Como há mais de um(a) " & Analistas(0, AnalistaAtual) & _
" ele/ela se chamará agora em diante de " & Analistas(0, AnalistaAtual) & " " & QtdAnalistaHomonimo
cbxCaptions.Add Analistas(0, AnalistaAtual), Analistas(1, AnalistaAtual)
.Caption = Analistas(0, AnalistaAtual)
Else
cbxCaptions.Add Analistas(0, AnalistaAtual), Analistas(1, AnalistaAtual)
.Caption = Analistas(0, AnalistaAtual)
End If
Else
cbxCaptions.Add Analistas(0, AnalistaAtual), Analistas(1, AnalistaAtual)
.Caption = Analistas(0, AnalistaAtual)
End If
.Value = True
cbxs.Add Chk, Chk.Name

.Top = 60 + AnalistaAtual * estilo
.Left = 350
.Width = 300
botaoStyle = .Top + 60
End With
BtnSelecionarAnalistasDisponiveis.Top = botaoStyle
Next AnalistaAtual
QtdNomes = cbxs.Count
QtdCbxs = cbxs.Count
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim resposta As Integer
If CloseMode = 0 Then
resposta = MsgBox("Ao fechar esta janela, sem clicar no botão, todos os analistas " _
& "serão considerados aptos para o trabalho hoje. Deseja prosseguir?", vbYesNo + vbExclamation)
If resposta = vbYes Then
Unload FrmAnalistasDisponiveis
Exit Sub
Else
Cancel = True
End If
Else
Cancel = False
End If
End Sub

