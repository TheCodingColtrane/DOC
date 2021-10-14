VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCadastroDocumento 
   Caption         =   "Cadastro de Documentos"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12360
   OleObjectBlob   =   "FrmCadastroDocumento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCadastroDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub BtnCadastrar_Click()
Dim ct As Integer
Dim Documento As CDocumento
Set Documento = New CDocumento
Dim Erro As Boolean
Dim Campo As Object
For Each Campo In Me.Controls
If TypeName(Campo) = "TextBox" Then
If Campo.Text = Empty Then
Campo.BorderColor = vbRed
Erro = True
End If
ElseIf TypeName(Campo) = "ComboBox" Then
If Campo.Text = Empty Then
Campo.BorderColor = vbRed
Erro = True
End If
End If
Next Campo
If Erro = True Then Exit Sub

Documento.celula = CbbCelula.Value
Documento.Cliente = CbbCliente.Value
Documento.Complexidade = CbbComplexidade.Value
Documento.Nome = TxtDocumento.Value
If CbbTipo.Value = "COMUM" Then
Documento.Tipo = 0
Else
Documento.Tipo = 1
End If
Documento.TempoMedioAnalise = TxtTempoMedio.Value
Documento.PrazoMaximoAnalise = CbbPrazoMaximo.Value
Documento.celula = CbbCelula.Value
ct = PostDocumentoAPI(Documento)
Unload FrmCadastroDocumento
End Sub



Private Sub btnFechar_Click()
Unload FrmCadastroDocumento
End Sub

Private Sub txtTempoMedio_AfterUpdate()
TxtTempoMedio.Value = Format(TxtTempoMedio.Value, "hh:mm:ss")
End Sub

Private Sub UserForm_Initialize()

Dim Documento As CDocumento
Set Documento = New CDocumento
CbbTipo.AddItem "BLOQUEIO"
CbbTipo.AddItem "COMUM"
CbbTipo.Value = "BLOQUEIO"
Dim complexidades As Integer: complexidades = 4
Dim ComplexidadeAtual, PrazoAtual, QtdPrazos As Integer
For ComplexidadeAtual = 1 To complexidades
CbbComplexidade.AddItem ComplexidadeAtual
Next ComplexidadeAtual
CbbComplexidade.Value = CbbComplexidade.List(2)
CbbComplexidade.Value = CbbComplexidade.List(2)

TxtTempoMedio.Value = "00:02:00"

If DocumentoNovo.Count > 0 Then
Dim PrazoMaximo As Variant
TxtDocumento.Value = DocumentoNovo(DocumentoNovo.Count).Nome
TxtDocumento.Enabled = False
If Documento.Tipo = "BLOQUEIO" Then

CbbCelula.Value = DocumentoNovo(1).celula
CbbCelula.Enabled = False
CbbCliente.Value = DocumentoNovo(1).Cliente
CbbCliente.Enabled = False
PrazoMaximo = DocumentoNovo(DocumentoNovo.Count).PrazoMaximoAnalise

QtdPrazos = UBound(PrazoMaximo, 2)

For PrazoAtual = 0 To QtdPrazos
CbbPrazoMaximo.AddItem PrazoMaximo(0, PrazoAtual)
Next PrazoAtual

CbbPrazoMaximo.Value = PrazoMaximo(0, 0)

Else

CbbCelula.Value = DocumentoNovo(1).celula
CbbCelula.Enabled = False
CbbCliente.Value = DocumentoNovo(1).Cliente
CbbCliente.Enabled = False

PrazoMaximo = DocumentoNovo(DocumentoNovo.Count).PrazoMaximoAnalise
QtdPrazos = UBound(PrazoMaximo)

For PrazoAtual = 0 To QtdPrazos
CbbPrazoMaximo.AddItem PrazoMaximo(PrazoAtual)
Next PrazoAtual

CbbPrazoMaximo.Value = PrazoMaximo(0)

End If
End If
End Sub
