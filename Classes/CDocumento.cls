VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pNome As String
Public pPrazoMaximoAnalise As Variant
Public pComplexidade As Integer
Public pTipo As String
Public pTempoMedioAnalise As Date
Public pCelula As Variant
Public pCliente As Variant
Public pCelulaID As Integer
Public pSLAID As Integer
Public Property Get celula() As String
celula = pCelula
End Property
Public Property Let celula(Value As String)
pCelula = Value
End Property
Public Property Get Cliente() As Variant
Cliente = pCliente
End Property
Public Property Let Cliente(Value As Variant)
pCliente = Value
End Property
Public Property Get Nome() As String
Nome = pNome
End Property
Public Property Let Nome(Value As String)
pNome = Value
End Property
Public Property Get PrazoMaximoAnalise() As Variant
PrazoMaximoAnalise = pPrazoMaximoAnalise
End Property
Property Let PrazoMaximoAnalise(Value As Variant)
pPrazoMaximoAnalise = Value
End Property
Public Property Get Complexidade() As Integer
Complexidade = pComplexidade
End Property
Property Let Complexidade(Value As Integer)
pComplexidade = Value
End Property
Public Property Get Tipo() As String
Tipo = pTipo
End Property
Property Let Tipo(Value As String)
pTipo = Value
End Property
Public Property Get TempoMedioAnalise() As String
TempoMedioAnalise = pTempoMedioAnalise
End Property
Property Let TempoMedioAnalise(Value As String)
pTempoMedioAnalise = Value
End Property
Public Property Get SLAID() As Integer
SLAID = pSLAID
End Property
Property Let SLAID(Value As Integer)
pSLAID = Value
End Property
Public Property Get CelulaID() As Integer
CelulaID = pCelulaID
End Property
Property Let CelulaID(Value As Integer)
pCelulaID = Value
End Property


