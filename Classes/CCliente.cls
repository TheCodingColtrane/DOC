VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pClienteID As Integer
Public pCelulaID As Integer
Public pCelulaNome As String
Public pNome As String
Public pTipo As Integer
Public pSLAID As Integer
Public Property Get SLAID() As Integer
SLAID = pSLAID
End Property
Public Property Let SLAID(Value As Integer)
pSLAID = Value
End Property
Public Property Get ClienteID() As Integer
ClienteID = pClienteID
End Property
Public Property Let ClienteID(Value As Integer)
pClienteID = Value
End Property
Public Property Get CelulaID() As Integer
CelulaID = pCelulaID
End Property
Public Property Let CelulaID(Value As Integer)
pCelulaID = Value
End Property
Public Property Get Nome() As String
Nome = pNome
End Property
Public Property Let Nome(Value As String)
pNome = Value
End Property
Public Property Get CelulaNome() As String
CelulaNome = pCelulaNome
End Property
Public Property Let CelulaNome(Value As String)
pCelulaNome = Value
End Property
Public Property Get Tipo() As Integer
Tipo = pTipo
End Property
Public Property Let Tipo(Value As Integer)
pTipo = Value
End Property

