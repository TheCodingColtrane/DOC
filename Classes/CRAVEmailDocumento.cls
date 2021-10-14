VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRAVEmailDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pPara As Variant
Public pCopia As Variant
Public pCopiaOculta As Variant
Public pAssunto As Variant
Public pMensagem As Variant
Public pAnexo As Variant
Public pLote As Boolean
'Public pAnalistaNome As Boolean

Public Property Get Para() As Variant
Para = pPara
End Property
Public Property Let Para(Value As Variant)
pPara = Value
End Property
Public Property Get Copia() As Variant
Copia = pCopia
End Property
Public Property Let Copia(Value As Variant)
pCopia = Value
End Property
Public Property Get CopiaOculta() As Variant
CopiaOculta = pCopiaOculta
End Property
Public Property Let CopiaOculta(Value As Variant)
pCopiaOculta = Value
End Property
Public Property Get Mensagem() As Variant
Mensagem = pMensagem
End Property
Public Property Let Mensagem(Value As Variant)
pMensagem = Value
End Property
Public Property Get Lote() As String
Lote = pLote
End Property
Public Property Let Lote(Value As String)
pLote = Value
End Property
Public Property Get Assunto() As String
Assunto = pAssunto
End Property
Public Property Let Assunto(Value As String)
pAssunto = Value
End Property
Public Property Get Anexo() As String
Anexo = pAnexo
End Property
Public Property Let Anexo(Value As String)
pAnexo = Value
End Property
'Public Property Get AnalistaNome() As String
'AnalistaNome = pAnalistaNome
'End Property
'Public Property Let AnalistaNome(Value As String)
'pAnalistaNome = Value
'End Property

Private Sub class_initializer()
Para = ""
Copia = ""
CopiaOculta = ""
Assunto = ""
Mensagem = ""
Anexo = ""
Lote = True
End Sub
