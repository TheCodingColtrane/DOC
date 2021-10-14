VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCelula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pCelulaID As Integer
Public pNome As String
Public pTipo As Integer
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
Public Property Get Tipo() As Integer
Tipo = pTipo
End Property
Public Property Let Tipo(Value As Integer)
pTipo = Value
End Property

