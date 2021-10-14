VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDocumento 
   Caption         =   "Documento"
   ClientHeight    =   10740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20955
   OleObjectBlob   =   "FrmDocumento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Cliente As CCliente
Private Documento As CDocumento
Private DocumentosCarregados, FiltroAtual As Variant
Private EstaCarregado, ClienteEstaFiltrado, Filtrar, DocumentoCarregado, SLACarregado As Boolean
Private ClientesDados As New Dictionary
Private CelulaDados  As New Dictionary
Private DocumentosDados As New Dictionary
Private ClienteAtual, DocumentoFiltrado As String
Private Sub TabStrip1_Change()

End Sub

Private Sub TbsDocumento_Change()

End Sub

Private Sub BtnAtualizar_Click()
Dim Documento As CDocumento
Set Documento = New CDocumento
Dim Erro As Boolean
Dim ct As Integer
If TxtEditDocumento.Value = "" Then
MsgBox "O campo documento está vazio!", vbCritical
Exit Sub
ElseIf TxtEditTempoMaximo.Value = "" Or Not IsDate(TxtEditTempoMaximo.Value) Then
MsgBox "O campo tempo médio está vazio!", vbCritical
Exit Sub
End If

If CbbEditCelulaID.Value = "" Or Not IsNumeric(CbbEditCelulaID.Value) Then
MsgBox "Houve um erro que o sistema não conseguiu compreender. CelulaID Vazio ou não numérico", vbCritical
Exit Sub
End If

If CbbEditClienteID.Value = "" Or Not IsNumeric(CbbEditClienteID.Value) Then
MsgBox "Houve um erro que o sistema não conseguiu compreender. ClienteID Vazio ou não numérico", vbCritical
Exit Sub
End If

If LblDocumentoID = "" Then
MsgBox "Houve um erro que o sistema não conseguiu compreender. ClienteID Vazio ou não numérico", vbCritical
Exit Sub
End If

If Erro = True Then Exit Sub
Documento.CelulaID = CInt(CbbEditCelulaID.Value)
Documento.celula = CbbEditCelula.Value
Documento.Cliente = CbbEditCliente.Value
Documento.Complexidade = CbbEditComplexidade.Value
Documento.Nome = TxtEditDocumento.Value
Documento.TempoMedioAnalise = TxtEditTempoMaximo.Value
Documento.PrazoMaximoAnalise = CbbEditPrazoMaximo.Value
Documento.Cliente = LblClienteID
Documento.Tipo = CbbEditTipo.Value

ct = PatchDocumentoAPI(Documento, CInt(LblDocumentoID.Caption))
Unload FrmCadastroDocumento
End Sub

Private Sub BtnCadastrar_Click()
Dim ct As Integer
Dim Documento As CDocumento
Set Documento = New CDocumento
Dim Erro As Boolean
If TxtDocumento.Value = "" Then
MsgBox "O campo documento está vazio!", vbCritical
Exit Sub
ElseIf TxtTempoMedio.Value = "" Or Not IsDate(TxtTempoMedio.Value) Then
MsgBox "O campo documento está vazio!", vbCritical
Exit Sub
End If


If Erro = True Then Exit Sub
Documento.CelulaID = CInt(CbbCelulaID.Value)
Documento.celula = CbbCelula.Value
Documento.Cliente = CbbCliente.Value
Documento.Complexidade = CbbComplexidade.Value
Documento.Nome = TxtDocumento.Value
Documento.SLAID = CInt(CbbSLAID.Value)
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

Private Sub CbbCelula_Change()
Dim celula As String
Dim Cliente As Variant
Dim Clientes, PrazoMaximo, PrazosMaximos As Variant
Dim QtdClientes, Indice As Integer
celula = CbbCelula.Value
Clientes = GetClienteAPI(CbbCelula.Value, 1)
If Not IsArrayEmpty(Clientes) Then
QtdClientes = UBound(Clientes, 2)
SLACarregado = False
CbbCliente.Clear
CbbSLAID.Clear
SLACarregado = True
For Cliente = 0 To QtdClientes
CbbCliente.AddItem Clientes(0, Cliente)
CbbDadosCliente.AddItem Clientes(0, Cliente)
CbbSLAID.AddItem Clientes(1, Cliente)
Next Cliente
CbbCliente.Value = CbbCliente.List(0)
CbbSLAID.Value = CbbSLAID.List(0)

PrazosMaximos = GetCelulaPrazoDocumentosAPI(CbbCelula.Value, 0)

If Not IsArrayEmpty(PrazosMaximos) Then
CbbPrazoMaximo.Clear
For Each PrazoMaximo In PrazosMaximos
CbbPrazoMaximo.AddItem PrazoMaximo
Next PrazoMaximo
CbbPrazoMaximo.Value = CbbPrazoMaximo.List(0)
End If
End If
Indice = CbbCelula.ListIndex
CbbCelulaID.Value = CbbCelulaID.List(Indice)

End Sub

Private Sub CbbCliente_Change()
If SLACarregado = True Then
Dim Indice As Integer
Indice = CbbCliente.ListIndex
CbbSLAID.Value = CbbSLAID.List(Indice)
End If
End Sub

Private Sub CbbDadosCelula_Change()

Dim Celulas As New Collection
Dim celula, Cliente, Clientes As Variant
Dim DocumentoClientes As New Dictionary

Dim li As ListItem
Dim Dado As Variant
Dim Dadoatual As Integer
Dim QtdDados, Indice, QtdClientes, I As Integer
Dim Analistas As CAnalista
Set Analistas = New CAnalista
Dim AnalistaDados As Variant
Indice = CbbDadosCelula.ListIndex
CbbDadosCelulaID.Value = CbbDadosCelulaID.List(Indice)

Dado = GetCelulaPrazoAPI(CbbDadosCelula.Value, 3)
DocumentosCarregados = Dado
FiltroAtual = Dado
If IsArrayEmpty(Dado) = False Then
'Dado = Dados.Item(1)

'CbbDadosCliente.Value = CbbCliente.List(0)
Clientes = GetClienteAPI(CbbDadosCelula.Value)
Filtrar = False
If IsArrayEmpty(Clientes) = False Then
QtdClientes = UBound(Clientes)
CbbDadosCliente.Clear
CbbDadosCliente.AddItem "Todos os clientes"
EstaCarregado = True
Filtrar = True
CbbDadosCliente.Value = CbbDadosCliente.List(0)

End If
For Each Cliente In Clientes
CbbDadosCliente.AddItem Cliente
Next Cliente
End If

For I = 0 To CbbDadosCelula.ListCount - 1
If CbbDadosCelula.List(I) = CbbDadosCelula.Value Then
Indice = FrmDocumento.CbbDadosCelula.ListIndex
CbbEditCelula.AddItem CbbDadosCelula.List(Indice)
CbbEditCelulaID.AddItem CbbDadosCelulaID.List(Indice)
CbbEditCelula.Value = CbbEditCelula.List(Indice)
CbbEditCelulaID.Value = CbbDadosCelulaID.List(Indice)
Else
CbbEditCelula.AddItem CbbDadosCelula.List(I)
CbbEditClienteID.AddItem CbbDadosCelulaID.List(I)
End If
Next I

EstaCarregado = True

End Sub

Private Sub CbbDadosCliente_Change()
Dim li As ListItem
Dim DocumentoAtual, TodosDocumentoCarregados, aux As Long
Dim DocumentosFiltradosFiltroCliente As New Collection

If Filtrar = True And CbbDadosCliente.ListCount > 0 Then
With LsvDadosDocumento


TodosDocumentoCarregados = UBound(DocumentosCarregados, 2)
.ListItems.Clear

EstaCarregado = False
CbbDadosDocumento.Clear
CbbDadosDocumento.AddItem "Todos os documentos"
If CbbDadosDocumento.List(0) <> "Todos os documentos" Then
CbbDadosDocumento.Value = CbbDadosDocumento.List(CbbDadosDocumento.ListCount - 1)
Else
CbbDadosDocumento.Value = CbbDadosDocumento.List(0)
End If

ClienteAtual = CbbDadosCliente.Value
DocumentoFiltrado = CbbDadosDocumento.Value
If ClienteAtual <> "Todos os clientes" Then
ClientesDados.RemoveAll
DocumentosDados.RemoveAll
For DocumentoAtual = 0 To TodosDocumentoCarregados
If DocumentosCarregados(1, DocumentoAtual) = CbbDadosCliente.Value Then
Set li = .ListItems.Add(, , DocumentosCarregados(0, DocumentoAtual))
If Not DocumentosDados.Exists(DocumentosCarregados(2, DocumentoAtual)) Then
DocumentosDados.Add DocumentosCarregados(2, DocumentoAtual), DocumentosCarregados(1, DocumentoAtual)
CbbDadosDocumento.AddItem DocumentosCarregados(2, DocumentoAtual)
End If
li.ListSubItems.Add , , DocumentosCarregados(7, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(6, DocumentoAtual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
li.ListSubItems.Add , , DocumentosCarregados(1, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(2, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(3, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(8, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(4, DocumentoAtual) = 0, "COMUM", "BLOQUEIO")
li.ListSubItems.Add , , DocumentosCarregados(5, DocumentoAtual)
End If
Next DocumentoAtual
Else
ClientesDados.RemoveAll
For DocumentoAtual = 0 To TodosDocumentoCarregados
If Not ClientesDados.Exists(DocumentosCarregados(1, DocumentoAtual)) Then
ClientesDados.Add DocumentosCarregados(1, DocumentoAtual), DocumentosCarregados(7, DocumentoAtual)
End If
Set li = .ListItems.Add(, , DocumentosCarregados(0, DocumentoAtual))
li.ListSubItems.Add , , DocumentosCarregados(7, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(6, DocumentoAtual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
li.ListSubItems.Add , , DocumentosCarregados(1, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(2, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(3, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(8, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(4, DocumentoAtual) = 0, "COMUM", "BLOQUEIO")
li.ListSubItems.Add , , DocumentosCarregados(5, DocumentoAtual)
CbbDadosDocumento.AddItem DocumentosCarregados(2, DocumentoAtual)
Next DocumentoAtual
End If


aux = 1


EstaCarregado = True

End With
DocumentoFiltrado = CbbDadosDocumento.Value
ClienteAtual = CbbDadosCliente.Value
End If
End Sub

Private Sub CbbDadosDocumento_Change()
Dim li As ListItem
Dim DocumentoAtual, TodosDocumentoCarregados, aux As Long


If ClienteAtual = "todos os clientes" And DocumentoFiltrado <> "Todos os documentos" Then
EstaCarregado = True
End If

If Filtrar = True And EstaCarregado = True Then

With LsvDadosDocumento


TodosDocumentoCarregados = UBound(DocumentosCarregados, 2)
.ListItems.Clear

DocumentoFiltrado = CbbDadosDocumento.Value
ClienteAtual = CbbDadosCliente.Value
If DocumentoFiltrado = "Todos os documentos" And ClienteAtual = "Todos os clientes" Then
EstaCarregado = True

For DocumentoAtual = 0 To TodosDocumentoCarregados
CbbDadosDocumento.AddItem DocumentosCarregados(2, DocumentoAtual)
Set li = .ListItems.Add(, , DocumentosCarregados(0, DocumentoAtual))
li.ListSubItems.Add , , DocumentosCarregados(7, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(6, DocumentoAtual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
li.ListSubItems.Add , , DocumentosCarregados(1, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(2, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(3, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(8, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(4, DocumentoAtual) = 0, "COMUM", "BLOQUEIO")
li.ListSubItems.Add , , DocumentosCarregados(5, DocumentoAtual)
Next DocumentoAtual
DocumentoFiltrado = CbbDadosDocumento.Value
aux = 1

ElseIf DocumentoFiltrado <> "Todos os documentos" And ClienteAtual = "Todos os clientes" Then
'EstaCarregado = False

For DocumentoAtual = 0 To TodosDocumentoCarregados
If DocumentosCarregados(2, DocumentoAtual) = CbbDadosDocumento.Value Then
CbbDadosDocumento.AddItem DocumentosCarregados(2, DocumentoAtual)
Set li = .ListItems.Add(, , DocumentosCarregados(0, DocumentoAtual))
li.ListSubItems.Add , , DocumentosCarregados(7, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(6, DocumentoAtual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
li.ListSubItems.Add , , DocumentosCarregados(1, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(2, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(3, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(8, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(4, DocumentoAtual) = 0, "COMUM", "BLOQUEIO")
li.ListSubItems.Add , , DocumentosCarregados(5, DocumentoAtual)
End If

Next DocumentoAtual
DocumentoFiltrado = CbbDadosDocumento.Value
aux = 1

'remover
ElseIf DocumentoFiltrado = "Todos os documentos" And ClienteAtual <> "Todos os clientes" Then
For DocumentoAtual = 0 To TodosDocumentoCarregados
If DocumentosCarregados(1, DocumentoAtual) = CbbDadosCliente.Value Then
CbbDadosDocumento.AddItem DocumentosCarregados(2, DocumentoAtual)
Set li = .ListItems.Add(, , DocumentosCarregados(0, DocumentoAtual))
li.ListSubItems.Add , , DocumentosCarregados(7, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(6, DocumentoAtual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
li.ListSubItems.Add , , DocumentosCarregados(1, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(2, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(3, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(8, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(4, DocumentoAtual) = 0, "COMUM", "BLOQUEIO")
li.ListSubItems.Add , , DocumentosCarregados(5, DocumentoAtual)
End If
'End If
Next DocumentoAtual
DocumentoFiltrado = CbbDadosDocumento.Value

ElseIf DocumentoFiltrado <> "Todos os documentos" And ClienteAtual <> "Todos os clientes" Then
For DocumentoAtual = 0 To TodosDocumentoCarregados
If DocumentosCarregados(1, DocumentoAtual) = CbbDadosCliente.Value And _
DocumentosCarregados(2, DocumentoAtual) = CbbDadosDocumento.Value Then
CbbDadosDocumento.AddItem DocumentosCarregados(2, DocumentoAtual)
Set li = .ListItems.Add(, , DocumentosCarregados(0, DocumentoAtual))
li.ListSubItems.Add , , DocumentosCarregados(7, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(6, DocumentoAtual) = 0, "MONITORAMENTO", "HOMOLOGAÇÃO")
li.ListSubItems.Add , , DocumentosCarregados(1, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(2, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(3, DocumentoAtual)
li.ListSubItems.Add , , DocumentosCarregados(8, DocumentoAtual)
li.ListSubItems.Add , , IIf(DocumentosCarregados(4, DocumentoAtual) = 0, "COMUM", "BLOQUEIO")
li.ListSubItems.Add , , DocumentosCarregados(5, DocumentoAtual)
End If
Next DocumentoAtual
End If
DocumentoFiltrado = CbbDadosDocumento.Value
End With
End If
End Sub

Private Sub CbbEditCelula_Change()
If DocumentoCarregado = True Then
Dim Cliente As Variant
Dim Clientes, PrazoMaximo, PrazosMaximos As Variant
Dim QtdClientes, Indice As Integer
Clientes = GetClienteAPI(CbbEditCelula.Value)
If Not IsArrayEmpty(Clientes) Then
QtdClientes = UBound(Clientes)
CbbCliente.Clear
For Each Cliente In Clientes
CbbEditCliente.AddItem Cliente
Next Cliente
CbbEditCliente.Value = CbbEditCliente.List(0)
Indice = CbbEditCelula.ListIndex
CbbEditCelulaID.Value = CbbEditCelulaID.List(Indice)
PrazosMaximos = GetCelulaPrazoDocumentosAPI(CbbEditCelula.Value, 0)

If Not IsArrayEmpty(PrazosMaximos) Then
CbbPrazoMaximo.Clear
For Each PrazoMaximo In PrazosMaximos
CbbEditPrazoMaximo.AddItem PrazoMaximo
Next PrazoMaximo
CbbEditPrazoMaximo.Value = CbbEditPrazoMaximo.List(0)
End If
End If
Indice = CbbEditCelula.ListIndex
CbbEditCelulaID.Value = CbbEditCelulaID.List(Indice)
End If

End Sub

Private Sub CbbEditCliente_Change()
If ClienteEstaFiltrado = False Then
Dim Indice, QtdDados, Dadoatual As Integer
Dim ClienteIDEncontrado As Boolean
QtdDados = ClientesDados.Count - 1

For Dadoatual = 0 To QtdDados
If CbbEditCliente.Value = ClientesDados.Keys(Dadoatual) Then
ClienteIDEncontrado = True
For Indice = 0 To CbbEditClienteID.ListCount - 1
If CInt(CbbEditClienteID.List(Indice)) = ClientesDados.Items(Dadoatual) Then
ClienteEstaFiltrado = True
CbbEditClienteID.Value = CbbEditClienteID.List(Indice)
Exit For
End If
Next Indice
If ClienteIDEncontrado = True Then
Exit For
End If
End If
Next Dadoatual
End If
ClienteEstaFiltrado = False
End Sub

Private Sub LsvDadosDocumento_DblClick()
If LsvDadosDocumento.SelectedItem Is Nothing Then
Exit Sub
End If

Dim Celulas, Celula_ As Variant
Dim Alterado As Boolean
Dim I, Indice, aux, QtdDados, QtdDadosCelula, QtdDadosCliente  As Integer
Dim ClienteID As New Dictionary
Dim PrazoMaximo As New Dictionary
Dim Complexidade As New Dictionary
CbbEditCelula.Clear
CbbEditCliente.Clear
CbbEditCelulaID.Clear
CbbEditClienteID.Clear
CbbEditTipo.Clear
CbbEditComplexidade.Clear
CbbEditPrazoMaximo.Clear

LblDocumentoID.Caption = CInt(LsvDadosDocumento.SelectedItem.Text)
LblClienteIDAtual.Caption = CInt(LsvDadosDocumento.SelectedItem.SubItems(1))
LblCelulaIDAtual.Caption = CInt(CbbDadosCelulaID.Value)
TxtEditTipoCliente.Value = IIf(LsvDadosDocumento.SelectedItem.SubItems(2) = "MONITORAMENTO", 0, 1)

QtdDados = UBound(DocumentosCarregados, 2)
QtdDadosCelula = CelulaDados.Count - 1
QtdDadosCliente = ClientesDados.Count - 1

For I = 0 To QtdDadosCelula
If CelulaDados.Keys(I) = CbbDadosCelula.Value Then
CbbEditCelula.AddItem CelulaDados.Keys(I)
CbbEditCelula.Value = CbbEditCelula.List(aux)
CbbEditCelulaID.AddItem CelulaDados.Items(I)
CbbEditCelulaID.Value = CbbEditCelulaID.List(aux)
Else
CbbEditCelula.AddItem CelulaDados.Keys(I)
CbbEditCelulaID.AddItem CelulaDados.Items(I)
aux = aux + 1
End If
Next I

aux = 0

For I = 0 To QtdDadosCliente
If ClientesDados.Keys(I) = LsvDadosDocumento.SelectedItem.SubItems(3) Then
CbbEditCliente.AddItem ClientesDados.Keys(I)
CbbEditCliente.Value = CbbEditCliente.List(aux)
CbbEditClienteID.AddItem ClientesDados.Items(I)
CbbEditClienteID.Value = CbbEditClienteID.List(aux)
Else
CbbEditCliente.AddItem ClientesDados.Keys(I)
CbbEditClienteID.AddItem ClientesDados.Items(I)
aux = aux + 1
End If
Next I


For I = 0 To QtdDados

If DocumentosCarregados(3, I) = LsvDadosDocumento.SelectedItem.SubItems(5) And _
DocumentosCarregados(1, I) = LsvDadosDocumento.SelectedItem.SubItems(3) Then
If Not PrazoMaximo.Exists(DocumentosCarregados(3, I)) Then
PrazoMaximo.Add DocumentosCarregados(3, I), 0
CbbEditPrazoMaximo.AddItem DocumentosCarregados(3, I)
CbbEditPrazoMaximo.Value = CbbEditPrazoMaximo.List(0)
End If
Else
If Not PrazoMaximo.Exists(DocumentosCarregados(3, I)) And DocumentosCarregados(1, I) = LsvDadosDocumento.SelectedItem.SubItems(3) Then
PrazoMaximo.Add DocumentosCarregados(3, I), 0
CbbEditPrazoMaximo.AddItem DocumentosCarregados(3, I)
End If
End If

If DocumentosCarregados(5, I) = LsvDadosDocumento.SelectedItem.SubItems(8) And _
DocumentosCarregados(1, I) = LsvDadosDocumento.SelectedItem.SubItems(3) Then
If Not Complexidade.Exists(DocumentosCarregados(5, I)) Then
Complexidade.Add DocumentosCarregados(5, I), 0
CbbEditComplexidade.AddItem DocumentosCarregados(5, I)
CbbEditComplexidade.Value = CbbEditComplexidade.List(0)
End If
Else
If Not Complexidade.Exists(DocumentosCarregados(5, I)) And DocumentosCarregados(1, I) = LsvDadosDocumento.SelectedItem.SubItems(3) Then
Complexidade.Add DocumentosCarregados(5, I), 0
CbbEditComplexidade.AddItem DocumentosCarregados(5, I)
End If
End If


Next I

TxtEditDocumento.Value = LsvDadosDocumento.SelectedItem.SubItems(4)
CbbEditTipo.AddItem "COMUM"
CbbEditTipo.AddItem "BLOQUEIO"
If LsvDadosDocumento.SelectedItem.SubItems(2) = "MONITORAMENTO" Then
CbbEditTipo.Value = CbbEditTipo.List(0)
Else
CbbEditTipo.Value = CbbEditTipo.List(0)
End If
TxtEditTempoMaximo.Value = LsvDadosDocumento.SelectedItem.SubItems(6)

MlpDocumento.Pages.Item(2).Visible = True
MlpDocumento.Pages.Item(2).Enabled = True
MlpDocumento.Value = 2
DocumentoCarregado = True

CbbEditCelula.Enabled = False
CbbEditCliente.Enabled = False




End Sub

Private Sub MultiPage1_Enter()


End Sub

Private Sub MlpDocumento_Change()
If DocumentoCarregado = True Then DocumentoCarregado = False
End Sub

Private Sub MlpDocumento_Enter()
Dim celula, Cliente, Clientes, Celulas, Dados, CelulaID As Variant
Dim Documento As CDocumento
Set Documento = New CDocumento
CbbTipo.AddItem "BLOQUEIO"
CbbTipo.AddItem "COMUM"
CbbTipo.Value = "BLOQUEIO"
With LsvDadosDocumento
.View = lvwReport
.Arrange = lvwAutoLeft
With .ColumnHeaders
.Clear
.Add , , "Documento ID", 70
.Add , , "Cliente ID", 70
.Add , , "Tipo de Cliente", 70
.Add , , "Cliente", 150
.Add , , "Documento", 500
.Add , , "Prazo", 50
.Add , , "Tempo Médio de Análise", 50
.Add , , "Tipo", 75
.Add , , "Complexidade", 75
End With
End With


Dim complexidades As Integer: complexidades = 5
Dim ComplexidadeAtual, PrazoAtual, QtdPrazos, QtdDados, Indice As Integer
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
QtdPrazos = UBound(PrazoMaximo, 2)

For PrazoAtual = 0 To QtdPrazos
CbbPrazoMaximo.AddItem PrazoMaximo(0, PrazoAtual)
Next PrazoAtual

CbbPrazoMaximo.Value = PrazoMaximo(0, 0)

End If
Else

Dados = GetCelulaAPI
If Not IsArrayEmpty(Dados) Then
QtdDados = UBound(Dados, 2)
For Celulas = 0 To QtdDados
CbbCelula.AddItem Dados(0, Celulas)
CbbCelulaID.AddItem Dados(1, Celulas)
CbbDadosCelula.AddItem Dados(0, Celulas)
CbbDadosCelulaID.AddItem Dados(1, Celulas)
CelulaDados.Add Dados(0, Celulas), Dados(1, Celulas)
Next Celulas
End If

'For Each celula In Dados
'CbbCelula.AddItem celula
'CbbDadosCelula.AddItem celula
'Next celula
CbbDadosCelulaID.Value = CbbDadosCelulaID.List(0)
CbbCelula.Value = CbbCelula.List(0)
CbbCelulaID.Value = CbbCelulaID.List(0)
CbbDadosCelula.Value = CbbDadosCelula.List(0)

Indice = FrmCliente.CbbCelula.ListIndex
CbbDadosCelulaID.Value = CbbDadosCelulaID.List(Indice)
Clientes = GetClienteAPI(CbbCelula.Value)

For Each Cliente In Clientes
CbbCliente.AddItem Cliente
Next Cliente
CbbCliente.Value = CbbCliente.List(0)
MlpDocumento.Pages.Item(2).Visible = False
MlpDocumento.Pages.Item(2).Enabled = False
MlpDocumento.Value = 1

End If
End Sub

Private Sub UserForm_Click()
Dim celula, Celulas As Variant
Set Celulas = GetCelulasAPI

For Each celula In Celulas
CbbCelula.AddItem celula
Next celula
End Sub
