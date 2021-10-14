VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRAVDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pDepositoID As Long
Public pCliente As String
Public pFornecedor As String
Public pUnidade As String
Public pDocumento As String
Public pEmpregado As String
Public pAnalista As String
Public pAnalistaEmail As String
Public pDataDeposito As Date
Public pDIA As Integer
Public pPrazoMaximoAnalise As Integer
Public pStatus As String
Public pTipo As String

Public Property Get DepositoID() As Long
DepositoID = pDepositoID
End Property
Public Property Let DepositoID(Value As Long)
pDepositoID = Value
End Property
Public Property Get Cliente() As String
Cliente = pCliente
End Property
Public Property Let Cliente(Value As String)
pCliente = Value
End Property
Public Property Get Fornecedor() As String
Fornecedor = pFornecedor
End Property
Public Property Let Fornecedor(Value As String)
pFornecedor = Value
End Property
Public Property Let Documento(Value As String)
pDocumento = Value
End Property
Public Property Get Documento() As String
Documento = pDocumento
End Property
Public Property Let Unidade(Value As String)
pUnidade = Value
End Property
Public Property Get Unidade() As String
Unidade = pUnidade
End Property
Public Property Let Empregado(Value As String)
pEmpregado = Value
End Property
Public Property Get Empregado() As String
Empregado = pEmpregado
End Property
Public Property Let Analista(Value As String)
pAnalista = Value
End Property
Public Property Get Analista() As String
Analista = pAnalista
End Property
Public Property Let AnalistaEmail(Value As String)
pAnalistaEmail = Value
End Property
Public Property Get AnalistaEmail() As String
AnalistaEmail = pAnalistaEmail
End Property
Public Property Let DataDeposito(Value As String)
pDataDeposito = Value
End Property
Public Property Get DataDeposito() As String
DataDeposito = pDataDeposito
End Property
Public Property Get Dia() As Long
Dia = pDIA
End Property
Public Property Let Dia(Value As Long)
pDIA = Value
End Property
Public Property Get PrazoMaximoAnalise() As Long
PrazoMaximoAnalise = pPrazoMaximoAnalise
End Property
Public Property Let PrazoMaximoAnalise(Value As Long)
pPrazoMaximoAnalise = Value
End Property
Public Property Get Status() As String
Status = pStatus
End Property
Public Property Let Status(Value As String)
pStatus = Value
End Property
Public Property Get Tipo() As String
Tipo = pTipo
End Property
Public Property Let Tipo(Value As String)
pTipo = Value
End Property
    
Private Sub class_initialize()
DepositoID = 0
Cliente = ""
Fornecedor = ""
Unidade = ""
Documento = ""
Empregado = ""
Analista = ""
AnalistaEmail = ""
DataDeposito = Now()
Dia = 0
PrazoMaximoAnalise = 0
Status = ""
Tipo = ""
End Sub
