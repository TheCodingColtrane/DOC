VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRAVPreferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pCelula As String
Public pCliente As String
Public pAnalista As String
Public pAmbiente As String
Public pAnalistas As Variant
Public pCelulaID As Integer
Public Property Get celula() As String
celula = pCelula
End Property
Public Property Let celula(Value As String)
pCelula = Value
End Property
Public Property Get Cliente() As String
Cliente = pCliente
End Property
Public Property Let Cliente(Value As String)
pCliente = Value
End Property
Public Property Let Analista(Value As String)
pAnalista = Value
End Property
Public Property Get Analista() As String
Analista = pAnalista
End Property
Public Property Let ambiente(Value As String)
pAmbiente = Value
End Property
Public Property Get ambiente() As String
ambiente = pAmbiente
End Property
Public Property Let Analistas(Value As Variant)
pAnalistas = Value
End Property
Public Property Get Analistas() As Variant
Analistas = pAnalistas
End Property
Public Property Let CelulaID(Value As Variant)
pCelulaID = Value
End Property
Public Property Get CelulaID() As Variant
CelulaID = pCelulaID
End Property

