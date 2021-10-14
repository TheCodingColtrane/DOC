VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPgInicial 
   Caption         =   "DOC"
   ClientHeight    =   12555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21360
   OleObjectBlob   =   "FrmPgInicial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPgInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ScrollSaved As Integer

Private Sub ListBox1_Click()


End Sub

Private Sub UserForm_Load()

End Sub

Private Sub BtnDadosDocumento_Click()
FrmDocumento.Show
End Sub

Private Sub btnNovoAnalista_Click()
FrmAnalista.Show
End Sub

Private Sub btnNovoCliente_Click()
FrmCliente.Show
End Sub

Private Sub cmdFuncionalidade_Click()
Dim planilhacaminho As String: planilhacaminho = GetArquivo
If planilhacaminho <> "" Then
FiltrarCelulas CbbCelula, CbbCliente, planilhacaminho, CbbClienteID.Value
Else
Exit Sub
End If
End Sub

Private Sub CbbCelula_Change()
Dim celula As String
Dim IndiceCelula As Integer
Dim Celula_ As String
Dim Cliente As Variant
Dim Clientes As Variant
'Dim Clientes As New Collection
Dim Colaboradores As Variant
CbbCliente.Clear
cbbColaborador.Clear
celula = CbbCelula.Value
'Celula_ = cbbCelula.Value
Clientes = GetClienteAPI(celula)

IndiceCelula = frmPgInicial.CbbCelula.ListIndex
If IsArrayEmpty(Clientes) = True Then
MsgBox "Não existe clientes para esta célula.", vbInformation
Exit Sub
End If
'Colaboradores = GetColaborador(Celula_)
'Dim qtdclientes, qtdcolaboradores As Integer
'If VBA.VarType(Colaboradores) = 8204 Then
'qtdcolaboradores = UBound(Colaboradores, 2)
'End If

Dim cont, contColaboradores As Integer
cont = UBound(Clientes)

CbbCliente.AddItem "Todos os clientes"
CbbCliente.Value = CbbCliente.List(0)
CbbClienteID.Value = CInt(CbbClienteID.List(IndiceCelula))
For Each Cliente In Clientes
CbbCliente.AddItem Cliente
Next Cliente

'If VBA.VarType(Colaboradores) = 8204 Then
'For contColaboradores = 0 To qtdcolaboradores
'cbbColaborador.AddItem Colaboradores(0, contColaboradores)
'Next contColaboradores
'End If

If cont > 0 Then
CbbCliente.Value = CbbCliente.List(0)
'cbbColaborador.Value = cbbColaborador.List(0)
Else
CbbCliente.Value = CbbCliente.List(0)
'If contColaboradores > 0 Then
'cbbColaborador.Value = cbbColaborador.List(0)
'Else
'cbbColaborador.Value = ""
'End If
End If
End Sub


Private Sub ScrollBar1_Change()

End Sub

Private Sub ScrollBar1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'ScrollSaved = ScrollBar1.Value
End Sub

Private Sub UserForm_Initialize()
'C:\Users\weverson.rafael\AppData\Roaming\Microsoft\AddIns
lblbemvindo.Caption = lblbemvindo & ", " & Application.UserName

Dim celula As Variant
'Dim Celulas As New Collection

Dim Celulas As Variant
Dim QtdCelulas As Integer
Celulas = GetCelulaAPI()
QtdCelulas = UBound(Celulas, 2)
If QtdCelulas > 0 Then
For celula = 0 To QtdCelulas
CbbCelula.AddItem Celulas(0, celula)
CbbClienteID.AddItem Celulas(1, celula)
Next celula

CbbClienteID.Value = CbbClienteID.List(0)
CbbCelula.Value = CbbCelula.List(0)
End If


End Sub

Private Sub Form_Load()
End Sub

