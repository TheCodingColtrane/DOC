VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCliente 
   Caption         =   "Cliente"
   ClientHeight    =   10530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17100
   OleObjectBlob   =   "FrmCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Cliente As CCliente
Private CelulaAnterior As String
Private Alterado, ClienteCarregado As Boolean
Private QtdCelulas As Integer
Private Sub ComboBox2_Change()

End Sub

Private Sub BtnAlterar_Click()

Set Cliente = New CCliente

'Set Analistas = New CAnalista
Dim QtdRegistroAtualizado As Integer
If TxtEditNomeCliente.Value <> "" Or Len(TxtEditNomeCliente) > 3 Then
Cliente.Nome = TxtEditNomeCliente.Value
Else
MsgBox "Verifique o campo nome", vbExclamation
Exit Sub
End If

If LblCelulaID = 0 Or LblCelulaID = "" And LblClienteID = 0 Or LblClienteID = "" Then
MsgBox "Acontenceu um erro em que o sistema não pôde entender", vbCritical
Exit Sub
End If
Cliente.ClienteID = CInt(LblClienteID)
If LblCelulaID <> CbbEditClienteID.Value Then
Cliente.CelulaID = CbbEditClienteID.Value
Else
Cliente.CelulaID = CInt(LblCelulaID)
End If

Cliente.Tipo = IIf(CbbEditTipo.Value = "Monitoramento", 0, 1)
Cliente.CelulaNome = CbbEditCelula.Value
QtdRegistroAtualizado = PatchClienteAPI(Cliente)
If QtdRegistroAtualizado > 0 Then
MlpClientes.Value = 1
End If

End Sub

Private Sub BtnCadastrar_Click()
If TxtCliente.Value <> Empty Then
Dim Cliente As CCliente
Set Cliente = New CCliente
Dim NovoCliente As Integer
Cliente.Nome = TxtCliente.Value
Cliente.Tipo = IIf(CbxTipo.Value = "MONITORAMENTO", 0, 1)
Cliente.CelulaNome = CbbCelula.Value
Cliente.CelulaID = CbbCelulaID.Value

NovoCliente = PostClienteAPI(Cliente)
Else
MsgBox "Campo cliente está vazio", vbExclamation
End If

If NovoCliente = 1 Then
Unload FrmCliente
End If

End Sub

Private Sub BtnCancelar_Click()
Unload FrmCliente
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub CbbCelula__Change()


End Sub

Private Sub CbbCelula_Change()
Dim Indice As Integer
Indice = FrmCliente.CbbCelula.ListIndex
CbbCelulaID.Value = CbbCelulaID.List(Indice)
End Sub

Private Sub CbbDadosCelula_Change()
If ClienteCarregado = True Then
Dim Dados, celula As Variant

With LsvClientes

Dim li As ListItem
Dim Dadoatual, QtdDados As Integer
Dados = GetClientesDadosAPI(CbbDadosCelula.Value)
If IsArrayEmpty(Dados) = False Then

QtdDados = UBound(Dados, 2)
.ListItems.Clear
For Dadoatual = 0 To QtdDados
Set li = .ListItems.Add(, , Dados(0, Dadoatual))
li.ListSubItems.Add , , Dados(1, Dadoatual)
li.ListSubItems.Add , , Dados(2, Dadoatual)
li.ListSubItems.Add , , IIf(Dados(3, Dadoatual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
Next Dadoatual
Else
LsvClientes.ListItems.Clear
End If
 .ColumnHeaders(1).Position = 1
End With
End If
End Sub

Private Sub CbbEditCelula_Change()
Dim I As Integer
If Alterado = False Then
CelulaAnterior = CbbCelula.Value
Alterado = True
End If
CbbCelula.Value = CbbEditCelula.Value
If QtdCelulas <> FrmCliente.CbbCelula.ListCount - 1 Then
For I = 0 To FrmCliente.CbbCelula.ListCount - 1
CbbEditCelula.AddItem CbbCelula.List(I)
CbbEditClienteID.AddItem CbbCelulaID.List(I)
Next I
End If
CbbEditClienteID.Value = CbbCelulaID.Value
End Sub

Private Sub LsvClientes_DblClick()
If LsvClientes.SelectedItem Is Nothing Then
Exit Sub
End If

Dim Celulas, Celula_ As Variant
Dim I As Integer

LblClienteID.Caption = CInt(LsvClientes.SelectedItem.Text)
LblCelulaID.Caption = LsvClientes.SelectedItem.SubItems(1)
TxtEditNomeCliente.Value = LsvClientes.SelectedItem.SubItems(2)

'CbbEditCelula.Value = "Nova York"
CbbEditTipo.Value = LsvClientes.SelectedItem.SubItems(3)
For I = 0 To FrmCliente.CbbCelula.ListCount - 1
CbbEditCelula.AddItem CbbCelula.List(I)
CbbEditClienteID.AddItem CbbCelulaID.List(I)
Next I
CbbEditClienteID.Value = LblCelulaID.Caption

CbbEditTipo.AddItem "Homologação"
CbbEditTipo.AddItem "Monitoramento"

CbbEditCelula.Value = CbbDadosCelula.Value

MlpClientes.Pages.Item(2).Visible = True
MlpClientes.Pages.Item(2).Enabled = True
MlpClientes.Value = 2
End Sub

Private Sub MlpClientes_Change()
Dim I As Integer
If MlpClientes.Pages.Item(1).Enabled = True And Alterado = True Then
CbbCelula.Value = CelulaAnterior
End If
End Sub


Private Sub MlpDocumentos_Change()

End Sub

Private Sub UserForm_Initialize()
'Dim celulas As New Collection
Dim celula As Variant
Dim Celulas As Variant
Celulas = GetCelulaAPI(0)
QtdCelulas = UBound(Celulas, 2)
If QtdCelulas > 0 Then
ClienteCarregado = False
For celula = 0 To QtdCelulas
CbbCelula.AddItem Celulas(0, celula)
CbbDadosCelula.AddItem Celulas(0, celula)
CbbCelulaID.AddItem Celulas(1, celula)
Next celula


CbxTipo.AddItem "Homologação"
CbxTipo.AddItem "Monitoramento"

CbxTipo.Value = CbxTipo.List(0)
CbbCelula.Value = CbbCelula.List(0)
CbbDadosCelula.Value = CbbDadosCelula.List(0)

With LsvClientes
.View = lvwReport
With .ColumnHeaders
.Clear
.Add , , "ID", 70
.Add , , "Célula ID", 100
.Add , , "Nome", 100
.Add , , "Tipo", 100
End With

Dim li As ListItem
Dim Dadoatual As Integer
Dim QtdDados As Integer
Dim Analistas As CAnalista
Set Analistas = New CAnalista
Dim AnalistaDados, Dados As Variant
Dim Tipo As String
Dados = GetClientesDadosAPI(CbbCelula.Value)
If IsArrayEmpty(Dados) = False Then

QtdDados = UBound(Dados, 2)
For Dadoatual = 0 To QtdDados
Set li = .ListItems.Add(, , Dados(0, Dadoatual))
li.ListSubItems.Add , , Dados(1, Dadoatual)
li.ListSubItems.Add , , Dados(2, Dadoatual)
li.ListSubItems.Add , , IIf(Dados(3, Dadoatual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
Next Dadoatual
Else
LsvClientes.ListItems.Clear
End If
 .ColumnHeaders(1).Position = 1
 End With
 ClienteCarregado = True
 End If
End Sub
