VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pAnalistaID As Integer
Public pNome As String
Public pEmail As String
Public pCelulaNome As String
Public pCelulaID As Integer
Public pCargo As Integer
Public pCargoComplexidade As Integer
Public pLider As Boolean
Public pLiderenca As String

Public Property Get CelulaID() As Integer
CelulaID = pCelulaID
End Property
Public Property Let CelulaID(Value As Integer)
pCelulaID = Value
End Property
Public Property Get AnalistaID() As Variant
AnalistaID = pAnalistaID
End Property
Public Property Let AnalistaID(Value As Variant)
pAnalistaID = Value
End Property
Public Property Get Nome() As String
Nome = pNome
End Property
Public Property Let Nome(Value As String)
pNome = Value
End Property
Public Property Get Email() As String
Email = pEmail
End Property
Public Property Let Email(Value As String)
pEmail = Value
End Property
Public Property Get Cargo() As Integer
Cargo = pCargo
End Property
Public Property Let Cargo(Value As Integer)
pCargo = Value
End Property
Public Property Get CargoComplexidade() As Integer
CargoComplexidade = pCargoComplexidade
End Property
Public Property Let CargoComplexidade(Value As Integer)
pCargoComplexidade = Value
End Property
Public Property Get lider() As Boolean
lider = pLider
End Property
Public Property Let lider(Value As Boolean)
pLider = Value
End Property
Public Property Get Liderenca() As String
Liderenca = pLiderenca
End Property
Public Property Let Liderenca(Value As String)
pLiderenca = Value
End Property
Public Property Get CelulaNome() As String
CelulaNome = pCelulaNome
End Property
Public Property Let CelulaNome(Value As String)
pCelulaNome = Value
End Property
