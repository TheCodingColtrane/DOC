VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDistribuiçãoInteligente 
   Caption         =   "Distribuição Inteligente"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16935
   OleObjectBlob   =   "FrmDistribuiçãoInteligente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDistribuiçãoInteligente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Documentos As Variant
Private ColunaCliente_ As Integer
Private ColunaFornecedor_ As Integer
Private ColunaUnidade_ As Integer
Private ColunaTipo_ As Integer
Private ColunaDocumento_ As Integer
Private ClienteCarregado As Boolean
Private FornecedorCarregado As Boolean
Private DocumentosFiltrados As New Dictionary

Private Sub CbbCliente_Change()
Dim LinhaFinal, LinhaAtual As Long
Dim Fornecedor As New Dictionary
CbbFornecedor.Clear
CbbUnidade.Clear
CbbTipo.Clear
CbbDocumento.Clear

LinhaFinal = UBound(Documentos, 2)
For LinhaAtual = 1 To LinhaFinal
If CbbCliente.Value = Documentos(LinhaAtual, ColunaCliente_) And Not Fornecedor.Exists(Documentos(LinhaAtual, ColunaCliente_)) Then
CbbFornecedor.AddItem Documentos(LinhaAtual, ColunaFornecedor_)
CbbUnidade.AddItem Documentos(LinhaAtual, ColunaUnidade_)
CbbTipo.AddItem Documentos(LinhaAtual, ColunaTipo_)
CbbDocumento.AddItem Documentos(LinhaAtual, ColunaDocumento_)
End If
CbbFornecedor.Value = CbbFornecedor.List(0)
CbbUnidade.Value = CbbUnidade.List(0)
CbbTipo.Value = CbbTipo.List(0)
CbbDocumento.Value = CbbDocumento.List(0)
Next LinhaAtual
ClienteCarregado = True
End Sub

Private Sub CbbFornecedor_Change()
If ClienteCarregado = True Then
Dim Unidade As New Dictionary
Dim Tipo As New Dictionary
Dim Documento As New Dictionary
Unidade.RemoveAll
Tipo.RemoveAll
Documento.RemoveAll
CbbUnidade.Clear
CbbTipo.Clear
CbbDocumento.Clear
If CbbFornecedor.Value = Documentos(LinhaAtual, ColunaFornecedor_) And Not Fornecedor.Exists(Documentos(LinhaAtual, ColunaCliente_)) Then
If Unidade.Exists(Documentos(LinhaAtual, ColunaUnidade_)) Then
CbbUnidade.AddItem Documentos(LinhaAtual, ColunaUnidade_)
End If
If Tipo.Exists(Documentos(LinhaAtual, ColunaTipo_)) Then
CbbTipo.AddItem Documentos(LinhaAtual, ColunaTipo_)
End If

If Documento.Exists(Documentos(LinhaAtual, ColunaDocumento_)) Then
CbbDocumento.AddItem Documentos(LinhaAtual, ColunaDocumento_)
End If
End If
CbbUnidade.Value = CbbUnidade.List(0)
CbbTipo.Value = CbbTipo.List(0)
CbbDocumento.Value = CbbDocumento.List(0)
End If
End Sub

Private Sub CbbTipoDocumento_Change()
If ClienteCarregado = True Then
Dim Unidade As New Dictionary
Dim Tipo As New Dictionary
Dim Documento As New Dictionary
Dim Fornecedor As New Dictionary
Unidade.RemoveAll
Tipo.RemoveAll
Documento.RemoveAll
CbbFornecedor.Clear
CbbUnidade.Clear
CbbTipo.Clear
CbbDocumento.Clear
If CbbTipo.Value = Documentos(LinhaAtual, ColunaTipo_) And Not Fornecedor.Exists(Documentos(LinhaAtual, ColunaTipo_)) Then
If Fornecedor.Exists(Documentos(LinhaAtual, ColunaFornecedor_)) Then
CbbFornecedor.AddItem Documentos(LinhaAtual, ColunaFornecedor_)
End If
If Unidade.Exists(Documentos(LinhaAtual, ColunaUnidade_)) Then
Unidade.AddItem Documentos(LinhaAtual, ColunaUnidade_)
End If

If Documento.Exists(Documentos(LinhaAtual, ColunaDocumento_)) Then
CbbDocumento.AddItem Documentos(LinhaAtual, ColunaDocumento_)
End If
End If
CbbUnidade.Value = CbbUnidade.List(0)
CbbTipo.Value = CbbTipo.List(0)
CbbDocumento.Value = CbbDocumento.List(0)
End If
End Sub

Private Sub CbbUnidade_Change()
If ClienteCarregado = True Then
Dim Unidade As New Dictionary
Dim Tipo As New Dictionary
Dim Documento As New Dictionary
Dim Fornecedor As New Dictionary
Unidade.RemoveAll
Tipo.RemoveAll
Documento.RemoveAll
CbbFornecedor.Clear
CbbUnidade.Clear
CbbTipo.Clear
CbbDocumento.Clear
If CbbUnidade.Value = Documentos(LinhaAtual, ColunaUnidade_) And Not Fornecedor.Exists(Documentos(LinhaAtual, ColunaUnidade_)) Then
If Fornecedor.Exists(Documentos(LinhaAtual, ColunaFornecedor_)) Then
CbbFornecedor.AddItem Documentos(LinhaAtual, ColunaFornecedor_)
End If
If Tipo.Exists(Documentos(LinhaAtual, ColunaTipo_)) Then
CbbTipo.AddItem Documentos(LinhaAtual, ColunaTipo_)
End If

If Documento.Exists(Documentos(LinhaAtual, ColunaDocumento_)) Then
CbbDocumento.AddItem Documentos(LinhaAtual, ColunaDocumento_)
End If
End If
CbbUnidade.Value = CbbUnidade.List(0)
CbbTipo.Value = CbbTipo.List(0)
CbbDocumento.Value = CbbDocumento.List(0)
End If
End Sub

Private Sub UserForm_Initialize()
 Dim PlanilhaTemp As Excel.Worksheet
Dim LinhaAtual, UltimaLinha As Long
Dim ColunaCliente, ColunaFornecedor, ColunaUnidade, ColunaTipo, ColunaDocumento As Integer
Dim Clientes As New Dictonary
Dim Fornecedores As New Dictionary
Dim ColunasRav As New cRAVColunasXL
ClienteCarregado = False
ColunaCliente = GetColunaIndiceNumerico(ColunasRav.Cliente)
ColunaFornecedor = GetColunaIndiceNumerico(ColunasRav.Fornecedor)
ColunaUnidade = GetColunaIndiceNumerico(ColunasRav.Unidade)
ColunaTipo = GetColunaIndiceNumerico(ColunasRav.Tipo)
ColunaDocumento = GetColunaIndiceNumerico(ColunasRav.Documento)

ColunaCliente_ = ColunaCliente
ColunaFornecedor_ = ColunaFornecedor
ColunaUnidade_ = ColunaUnidade
ColunaTipo_ = ColunaTipo
ColunaDocumento_ = ColunaDocumento

Set ColunasRav = GetColunasIndiceAlfabetico(PastadeTrabalhoRAV)
Set PlanilhaTemp = PastadeTrabalhoRAV.Worksheets(2)
UltimaLinha = PlanilhaTemp.Range("A:A").End(xlDown).Row
Documentos = PlanilhaTemp.Range("A2" & ColunasRav.Linha & UltimaLinha).Value2

For LinhaAtual = 1 To UltimaLinha
If Not Clientes.Exists(Documentos(LinhaAtual, ColunaCliente)) Then
Clientes.Add Documentos(LinhaAtual, ColunaCliente), 1
CbbCliente.AddItem Documentos(LinhaAtual, ColunaCliente)
End If
Next LinhaAtual
CbbCliente.Value = CbbCliente.List(0)

End Sub
