VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAnalista 
   Caption         =   "Analista"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17085
   OleObjectBlob   =   "FrmAnalista.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private celula As String
Private Analistas As CAnalista
Private Dados As New Collection
Private CelulaAnterior As String
Private Alterado As Boolean
Private CelulaAlterada As Boolean
Private QtdCelulas As Integer
Private Sub BtnAtualizar_Click()
Set Analistas = New CAnalista
Dim QtdRegistroAtualizado As Integer
If TxtEditEmail.Value <> "" Or Len(TxtEditEmail.Value) > 10 Then
Analistas.Email = TxtEditEmail.Value
Else
MsgBox "Verifique o campo e-mail", vbExclamation
Exit Sub
End If
If TxtEditNome.Value <> "" Or Len(TxtEditNome.Value) > 2 Then
Analistas.Nome = TxtEditNome.Value
Else
MsgBox "Verifique o campo nome", vbExclamation
Exit Sub
End If
'LblCelulaIDAtual = LsvColaboradores.
If CbbEditCargo.Value <> "" Or Len(CbbEditCargo.Value) > 3 Then
Analistas.Cargo = GetCargoNumero(CbbEditCargo.Value)
Else
MsgBox "Verifique o campo cargo", vbExclamation
Exit Sub
End If

Select Case CbbEditCargo.Value
Case "Estagiário"
Analistas.CargoComplexidade = 1
Case "Auxiliar I"
Analistas.CargoComplexidade = 2
Case "Auxiliar II"
Analistas.CargoComplexidade = 3
Case "Auxiliar III"
Analistas.CargoComplexidade = 3
Case "Assistente I"
Analistas.CargoComplexidade = 4
Case "Assistente II"
Analistas.CargoComplexidade = 4
Case "Assistente III"
Analistas.CargoComplexidade = 4
Case "Analista I"
Case "Analista II"
Analistas.CargoComplexidade = 5
Case "Analista III"
Analistas.CargoComplexidade = 5
Case "Especialista I"
Analistas.CargoComplexidade = 5
Case "Coordenador"
Analistas.CargoComplexidade = 5
Case Else

End Select

Analistas.CelulaNome = CbbEditCelula.Value
Analistas.lider = ChlEditELider.Value
Analistas.AnalistaID = CInt(LblAnalistaID.Caption)
If LblCelulaIDAtual <> CbbEditCelulaID.Value Then
Analistas.CelulaID = CInt(CbbEditCelulaID.Value)
Else
Analistas.CelulaID = CInt(LblCelulaIDAtual.Caption)
End If
QtdRegistroAtualizado = PatchAnalistaAPI(Analistas)
If QtdRegistroAtualizado > 0 Then
MlpAnalistas.Value = 1
End If
End Sub

Private Sub BtnCadastrar_Click()
Set Analistas = New CAnalista
Dim NovoRegistro As Integer
Select Case CbbCargo.Value
Case "Estagiário"
Analistas.Cargo = 1
Analistas.CargoComplexidade = 2
Case "Auxiliar I"
Analistas.Cargo = 2
Analistas.CargoComplexidade = 3
Case "Auxiliar II"
Analistas.Cargo = 3
Analistas.CargoComplexidade = 3
Case "Auxiliar III"
Analistas.Cargo = 4
Analistas.CargoComplexidade = 3
Case "Assistente I"
Analistas.Cargo = 5
Analistas.CargoComplexidade = 4
Case "Assistente II"
Analistas.Cargo = 6
Analistas.CargoComplexidade = 4
Case "Assistente III"
Analistas.Cargo = 7
Analistas.CargoComplexidade = 4
Case "Analista I"
Analistas.Cargo = 8
Case "Analista II"
Analistas.CargoComplexidade = 5
Analistas.Cargo = 9
Case "Analista III"
Analistas.Cargo = 10
Analistas.CargoComplexidade = 5
Case "Especialista I"
Analistas.Cargo = 11
Analistas.CargoComplexidade = 5
Case "Coordenador"
Analistas.Cargo = 12
Analistas.CargoComplexidade = 5
Case Else
MsgBox "Houve um erro que o sistema não pôde entender", vbCritical
Exit Sub
End Select

Analistas.Email = TxtEmail.Value
Analistas.Nome = TxtNome.Value
Analistas.lider = ChkELider.Value
Analistas.CelulaID = CbbCelulaID.Value
Analistas.CelulaNome = CbbCelula.Value
NovoRegistro = PostAnalistaAPI(Analistas)

If NovoRegistro > 0 Then
TxtEmail.Value = ""
TxtNome.Value = ""
Analistas.lider = False
End If

End Sub

Private Sub BtnCancelar_Click()
Unload FrmAnalista
End Sub

Private Sub CbbCelula__Change()
celula = CbbCelula_.Value
Dim Dado As Variant
Dim QtdDados, Dadoatual, Indice  As Integer
Indice = CbbCelula_.ListIndex
CbbSelectClienteID.Value = CbbSelectClienteID.List(Indice)
'Set Dados = GetColaboradorDadosResumidosAPI(celula, 0)
Dado = GetAnalistasDadosCompletosAPI(celula, CInt(CbbSelectClienteID.Value))
If IsArrayEmpty(Dado) = False Then
QtdDados = UBound(Dado, 2)
Dim li As ListItem
With LsvColaboradores
.ListItems.Clear
For Dadoatual = 0 To QtdDados
   Set li = .ListItems.Add(, , Dado(0, Dadoatual))
    li.ListSubItems.Add , , Dado(1, Dadoatual)
    li.ListSubItems.Add , , Dado(2, Dadoatual)

Select Case Dado(3, Dadoatual)
Case 1
li.ListSubItems.Add , , "Estagiário"
Case 2
li.ListSubItems.Add , , "Auxiliar I"
Case 3
li.ListSubItems.Add , , "Auxiliar II"
Case 4
li.ListSubItems.Add , , "Auxiliar III"
Case 5
li.ListSubItems.Add , , "Assistente I"
Case 6
li.ListSubItems.Add , , "Assistente II"
Case 7
li.ListSubItems.Add , , "Assistente III"
Case 8
li.ListSubItems.Add , , "Analista I"
Case 9
li.ListSubItems.Add , , "Analista II"
Case 10
li.ListSubItems.Add , , "Analista III"
Case 11
li.ListSubItems.Add , , "Especialista I"
Case 12
li.ListSubItems.Add , , "Coordenador"
Case Else
End Select
li.ListSubItems.Add , , Dado(4, Dadoatual)
li.ListSubItems.Add , , IIf(Dado(6, Dadoatual) = True, "Sim", "Não")
li.ListSubItems.Add , , Dado(5, Dadoatual)
Next Dadoatual
 .ColumnHeaders(1).Position = 2
End With
Else
LsvColaboradores.ListItems.Clear
End If

End Sub

Private Sub CbbCelula_Change()
Dim Indice As Integer
Indice = FrmAnalista.CbbCelula.ListIndex
CbbCelulaID.Value = CbbCelulaID.List(Indice)
End Sub

Private Sub CbbEditCelula_Change()
Dim I As Integer
If Alterado = False Then
CelulaAnterior = CbbCelula.Value
Alterado = True
End If
CbbCelula.Value = CbbEditCelula.Value
If QtdCelulas <> FrmAnalista.CbbCelula.ListCount - 1 Then
For I = 0 To FrmAnalista.CbbCelula.ListCount - 1
CbbEditCelula.AddItem CbbCelula.List(I)
CbbEditCelulaID.AddItem CbbCelulaID.List(I)
Next I
End If
CbbEditCelulaID.Value = CbbCelulaID.Value
End Sub
Private Sub LsvColaboradores_DblClick()
Dim Celula_ As Variant
CbbEditCargo.AddItem "Estagiário"
CbbEditCargo.AddItem "Auxiliar I"
CbbEditCargo.AddItem "Auxiliar II"
CbbEditCargo.AddItem "Auxiliar III"
CbbEditCargo.AddItem "Assistente I"
CbbEditCargo.AddItem "Assistente II"
CbbEditCargo.AddItem "Assistente III"
CbbEditCargo.AddItem "Analista I"
CbbEditCargo.AddItem "Analista II"
CbbEditCargo.AddItem "Analista III"
CbbEditCargo.AddItem "Especialista I"
CbbEditCargo.AddItem "Especialista II"
CbbEditCargo.AddItem "Especialista III"
CbbEditCargo.AddItem "Coordenador I"
CbbEditCargo.AddItem "Coordenador II"
CbbEditCargo.AddItem "Coordenador III"
If LsvColaboradores.SelectedItem Is Nothing Then
Exit Sub
End If
LblAnalistaID.Caption = CInt(LsvColaboradores.SelectedItem.Text)
TxtEditNome.Value = LsvColaboradores.SelectedItem.SubItems(1)
TxtEditEmail.Value = LsvColaboradores.SelectedItem.SubItems(2)
CbbEditCargo.Value = LsvColaboradores.SelectedItem.SubItems(3)
LblCargoComplexidade.Caption = CInt(LsvColaboradores.SelectedItem.SubItems(4))
For I = 0 To FrmAnalista.CbbCelula.ListCount - 1
CbbEditCelula.AddItem CbbCelula.List(I)
Next I
ChlEditELider.Value = IIf(LsvColaboradores.SelectedItem.SubItems(5) = "Sim", True, False)
CbbEditCelula.Value = CbbCelula_.Value

MlpAnalistas.Pages.Item(2).Visible = True
MlpAnalistas.Pages.Item(2).Enabled = True
MlpAnalistas.Value = 2
CelulaAlterada = False
If CelulaAlterada = False Then
LblCelulaIDAtual = CbbEditCelulaID.Value
End If
'FrmAtualizarAnalista.Show
End Sub

Private Sub MlpAnalistas_Change()
Dim I As Integer
If MlpAnalistas.Pages.Item(1).Enabled = True And Alterado = True Then
CbbCelula.Value = CelulaAnterior
End If
End Sub

Private Sub MlpAnalistas_Enter()
'Dim Celulas As New Collection

Dim celula, Celulas As Variant
Celulas = GetCelulaAPI
QtdCelulas = UBound(Celulas, 2)
If IsArrayEmpty(Celulas) = False Then
For celula = 0 To QtdCelulas
CbbCelula.AddItem Celulas(0, celula)
CbbCelulaID.AddItem Celulas(1, celula)
CbbSelectClienteID.AddItem Celulas(1, celula)
Next celula
CbbSelectClienteID.Value = CbbSelectClienteID.List(0)
CbbCelula.Value = CbbCelula.List(0)

CbbCargo.AddItem "Estagiário"
CbbCargo.AddItem "Auxiliar I"
CbbCargo.AddItem "Auxiliar II"
CbbCargo.AddItem "Auxiliar III"
CbbCargo.AddItem "Assistente I"
CbbCargo.AddItem "Assistente II"
CbbCargo.AddItem "Assistente III"
CbbCargo.AddItem "Analista I"
CbbCargo.AddItem "Analista II"
CbbCargo.AddItem "Analista III"
CbbCargo.AddItem "Especialista I"
CbbCargo.AddItem "Especialista II"
CbbCargo.AddItem "Especialista III"

CbbCargo.Value = CbbCargo.List(1)

With LsvColaboradores
.View = lvwReport
With .ColumnHeaders
.Clear
.Add , , "ID", 70
.Add , , "Nome", 200
.Add , , "E-mail", 200
.Add , , "Cargo", 200
.Add , , "Complexidade do Cargo", 200
.Add , , "É Líder ? ", 200
.Add , , "Liderança ", 200
End With

Dim li As ListItem
Dim Dado As Variant
Dim Dadoatual As Integer
Dim QtdDados As Integer
Dim Analistas As CAnalista
Set Analistas = New CAnalista
Dim AnalistaDados As Variant


Dado = GetAnalistasDadosCompletosAPI(CbbCelula.Value, CInt(CbbSelectClienteID.Value))
If IsArrayEmpty(Dado) = False Then
QtdDados = UBound(Dado, 2)
For Dadoatual = 0 To QtdDados
Set li = .ListItems.Add(, , Dado(0, Dadoatual))
li.ListSubItems.Add , , Dado(1, Dadoatual)
li.ListSubItems.Add , , Dado(2, Dadoatual)
Select Case Dado(3, Dadoatual)
Case 1
li.ListSubItems.Add , , "Estagiário"
Case 2
li.ListSubItems.Add , , "Auxiliar I"
Case 3
li.ListSubItems.Add , , "Auxiliar II"
Case 4
li.ListSubItems.Add , , "Auxiliar III"
Case 5
li.ListSubItems.Add , , "Assistente I"
Case 6
li.ListSubItems.Add , , "Assistente II"
Case 7
li.ListSubItems.Add , , "Assistente III"
Case 8
li.ListSubItems.Add , , "Analista I"
Case 9
li.ListSubItems.Add , , "Analista II"
Case 10
li.ListSubItems.Add , , "Analista III"
Case 11
li.ListSubItems.Add , , "Especialista I"
Case 12
li.ListSubItems.Add , , "Coordenador"
Case Else
MsgBox "Houve um erro que o sistema não pôde entender", vbCritical
End Select
li.ListSubItems.Add , , Dado(4, Dadoatual)
li.ListSubItems.Add , , IIf(Dado(6, Dadoatual) = True, "Sim", "Não")
li.ListSubItems.Add , , Dado(5, Dadoatual)
Next Dadoatual

 .ColumnHeaders(1).Position = 1
 End If
End With

For celula = 0 To QtdCelulas
CbbCelula_.AddItem Celulas(0, celula)
Next celula

CbbCelula_.Value = CbbCelula_.List(0)
MlpAnalistas.Value = 1
End If
End Sub

Private Function GetCargo(Cargo As Integer) As String
Select Case Cargo
Case 1
GetCargo = "Estagiário"
Case 2
GetCargo = "Auxiliar I"
Case 3
GetCargo = "Auxiliar II"
Case 4
GetCargo = "Auxiliar III"
Case 5
GetCargo = "Assistente I"
Case 6
GetCargo = "Assistente II"
Case 7
GetCargo = "Assistente III"
Case 8
GetCargo = "Analista I"
Case 9
GetCargo = "Analista II"
Case 10
GetCargo = "Analista III"
Case 11
GetCargo = "Especialista I"
Case 12
GetCargo "Coordenador"
Case Else
End Select
End Function

Private Function GetCargoNumero(Cargo As String) As Integer

Select Case Cargo
Case "Estagiário"
GetCargoNumero = 1
Case "Auxiliar I"
GetCargoNumero = 2
Case "Auxiliar II"
GetCargoNumero = 3
Case "Auxiliar III"
GetCargoNumero = 4
Case "Assistente I"
GetCargoNumero = 5
Case "Assistente II"
GetCargoNumero = 6
Case "Assistente III"
GetCargoNumero = 7
Case "Analista I"
GetCargoNumero = 8
Case "Analista II"
GetCargoNumero = 9
Case "Analista III"
GetCargoNumero = 10
Case "Especialista I"
GetCargoNumero = 11
Case "Coordenador"
GetCargoNumero = 12
Case Else
MsgBox "Houve um erro que o sistema não pôde entender", vbCritical
End Select
End Function

