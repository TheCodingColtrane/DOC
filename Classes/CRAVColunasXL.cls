VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRAVColunasXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public pProtocolo As String
Public pCliente As String
Public pFornecedor As String
Public pDataInclusao As String
Public pDocumento As String
Public pEmpregado As String
Public pAnalista As String
Public pUnidade As String
Public pMesDeposito As String
Public pDataInicio As String
Public pTempoEmAnalise As String
Public pJustificativaEmAnalise As String
Public pQLP As String
Public pDivida As String
Public pInadimpliencia As String
Public pFimInadimplencia As String
Public pDiasNoSistema As String
Public pTipo As String
Public pStatus As String
Public pDIA As String
Public pFeriados As String
Public pPrazoMaximoAnalise As String
Public pDiasAguardandoAnalise As String
Public pDocumentoComplexidade As String
Public pAnalistaEmail As String
Public pLinha As String

Public Property Get DataInclusao() As String
DataInclusao = pDataInclusao
End Property
Public Property Let DataInclusao(Value As String)
pDataInclusao = Value
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
Public Property Let Divida(Value As String)
pDivida = Value
End Property
Public Property Get Divida() As String
Divida = pDivida
End Property
Public Property Let Inadimpliencia(Value As String)
pInadimpliencia = Value
End Property
Public Property Get Inadimpliencia() As String
Inadimpliencia = pInadimpliencia
End Property
Public Property Let FimInadimplencia(Value As String)
pFimInadimplencia = Value
End Property
Public Property Get FimInadimplencia() As String
FimInadimplencia = pFimInadimplencia
End Property
Public Property Let DiasNoSistema(Value As String)
pDiasNoSistema = Value
End Property
Public Property Get DiasNoSistema() As String
DiasNoSistema = pDiasNoSistema
End Property
Public Property Let DiasAguardandoAnalise(Value As String)
pDiasAguardandoAnalise = Value
End Property
Public Property Get DiasAguardandoAnalise() As String
DiasAguardandoAnalise = pDiasAguardandoAnalise
End Property
Public Property Let Tipo(Value As String)
pTipo = Value
End Property
Public Property Get Tipo() As String
Tipo = pTipo
End Property
Public Property Get Status() As String
Status = pStatus
End Property
Public Property Let Status(Value As String)
pStatus = Value
End Property
Public Property Get PrazoMaximoAnalise() As String
PrazoMaximoAnalise = pPrazoMaximoAnalise
End Property
Public Property Let PrazoMaximoAnalise(Value As String)
pPrazoMaximoAnalise = Value
End Property
Public Property Get Dia() As String
Dia = pDIA
End Property
Public Property Let Dia(Value As String)
pDIA = Value
End Property
Public Property Get Feriados() As String
Feriados = pFeriados
End Property
Public Property Let Feriados(Value As String)
pFeriados = Value
End Property
Public Property Get Protocolo() As String
Protocolo = pProtocolo
End Property
Public Property Let Protocolo(Value As String)
pProtocolo = Value
End Property
Public Property Get Unidade() As String
Unidade = pUnidade
End Property
Public Property Let Unidade(Value As String)
pUnidade = Value
End Property
Public Property Get MesDeposito() As String
MesDeposito = pMesDeposito
End Property
Public Property Let MesDeposito(Value As String)
pMesDeposito = Value
End Property
Public Property Get DataInicio() As String
DataInicio = pDataInicio
End Property
Public Property Let DataInicio(Value As String)
pDataInicio = Value
End Property
Public Property Get TempoEmAnalise() As String
TempoEmAnalise = pTempoEmAnalise
End Property
Public Property Let TempoEmAnalise(Value As String)
pTempoEmAnalise = Value
End Property
Public Property Get JustificativaEmAnalise() As String
JustificativaEmAnalise = pJustificativaEmAnalise
End Property
Public Property Let JustificativaEmAnalise(Value As String)
pJustificativaEmAnalise = Value
End Property
Public Property Get QLP() As String
QLP = pQLP
End Property
Public Property Let QLP(Value As String)
pQLP = Value
End Property
Public Property Get Empregado() As String
Empregado = pEmpregado
End Property
Public Property Let Empregado(Value As String)
pEmpregado = Value
End Property
Public Property Get Analista() As String
Analista = pAnalista
End Property
Public Property Let Analista(Value As String)
pAnalista = Value
End Property
Public Property Get AnalistaEmail() As String
AnalistaEmail = pAnalistaEmail
End Property
Public Property Let AnalistaEmail(Value As String)
pAnalistaEmail = Value
End Property
Public Property Get DocumentoComplexidade() As String
DocumentoComplexidade = pDocumentoComplexidade
End Property
Public Property Let DocumentoComplexidade(Value As String)
pDocumentoComplexidade = Value
End Property
Public Property Get Linha() As String
Linha = pLinha
End Property
Public Property Let Linha(Value As String)
pLinha = Value
End Property

