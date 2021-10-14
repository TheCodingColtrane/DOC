Attribute VB_Name = "Data"
Option Explicit
Public NovoRegistro As Integer
Public DocumentoNovo As New Collection
Private Const CelulaAPI As String = "https://apmsdocapiprd.azurewebsites.net/celula"
Private Const URIDesenvolvimentoAPI As String = "https://localhost:44377/celula"
Public Function GetCelulaAPI(Optional Tipo As Integer) As Variant
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI As String
Dim JSON As String
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
If Tipo = 0 Then
URI = CelulaAPI
Else
URI = CelulaAPI & "?tipo=1"
End If
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.Send

Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Nenhuma célula encontrada", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object
If Tipo = 0 Then
ReDim resposta(1, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nome")
resposta(1, aux) = obj("celulaId")
aux = aux + 1
Next obj
ElseIf Tipo = 1 Then
ReDim resposta(2, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("celulaId")
resposta(1, aux) = obj("nome")
resposta(2, aux) = obj("tipo")
aux = aux + 1
Next obj
Else
ReDim resposta(2, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("err")
resposta(1, aux) = obj("msg")
resposta(2, aux) = obj("tipo")
aux = aux + 1
Next obj
End If
GetCelulaAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err.Description, vbCritical + vbOKOnly
End Function
Public Function GetClienteAPI(celula As String, Optional Tipo As Integer) As Variant
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI As String
Dim JSON As String
celula = URLEncode(celula)
If Tipo = 0 Then
URI = CelulaAPI & "/" & celula & "/clientes/dados?tipo=0"
Else
URI = CelulaAPI & "/" & celula & "/clientes/dados?tipo=1"
End If
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json;charset=UTF-8"
oReq.Send
'oReq.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'oReq.SetRequestHeader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'oReq.SetRequestHeader "Accept-Language", "en-us,en;q=0.5"
'oReq.SetRequestHeader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"


Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)

QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Não há clientes para esta célula", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object
If Tipo = 0 Then
ReDim resposta(0 To QtdResultados)
For Each obj In objJSON
resposta(aux) = obj("nome")
aux = aux + 1
Next obj
Else
ReDim resposta(1, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nome")
resposta(1, aux) = obj("slaid")
aux = aux + 1
Next obj
End If
GetClienteAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
End Function
Public Function GetClientesDadosAPI(celula As String) As Variant
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI As String
celula = URLEncode(celula)
URI = "https://apmsdocapiprd.azurewebsites.net/celula/" & celula & "/clientes"
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
'oReq.setrequestheader "Connection", "Keep-Alive"
'oReq.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'oReq.setrequestheader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
'oReq.setrequestheader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"
'oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send

Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
'MsgBox oReq.ResponseText
QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Não há clientes para esta célula", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object
ReDim resposta(3, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("clienteId")
resposta(1, aux) = obj("celulaId")
resposta(2, aux) = obj("nome")
resposta(3, aux) = obj("tipo")
aux = aux + 1
Next obj
GetClientesDadosAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
End Function
Public Function GetCelulaPrazoDocumentosAPI(celula As String, Optional CelulaID As Integer) As Variant
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux, QtdNulos As Integer
Dim PossuiNulos As Boolean
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI As String
Dim JSON As String
celula = URLEncode(celula)
If CelulaID > 0 Then
URI = CelulaAPI & "/" & RAVPreferencias(1).CelulaID & "/" & celula & "/documentos/prazo"
Else
URI = CelulaAPI & "/0/" & celula & "/documentos/prazo"
End If
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json;charset=UTF-8"
oReq.Send
'oReq.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'oReq.SetRequestHeader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'oReq.SetRequestHeader "Accept-Language", "en-us,en;q=0.5"
'oReq.SetRequestHeader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"


Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)

QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Não há clientes para esta célula", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object

ReDim resposta(0 To QtdResultados)
For Each obj In objJSON
If obj("prazoMaximoAnalise") > 0 Then
resposta(aux) = obj("prazoMaximoAnalise")
aux = aux + 1
Else
PossuiNulos = True
QtdNulos = QtdNulos + 1
End If
Next obj
If PossuiNulos = True Then
Dim auxResposta As Variant
aux = UBound(resposta) - QtdNulos
auxResposta = resposta
ReDim resposta(0 To aux)
For aux = 0 To UBound(auxResposta)
If auxResposta(aux) > 0 Then
resposta(aux) = auxResposta(aux)
End If
Next aux
End If

GetCelulaPrazoDocumentosAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err.Description, vbCritical + vbOKOnly
End Function


Public Function GetCelulaPrazoAPI(celula As String, Tipo As Integer, Optional termo As String) As Variant
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI, Tempo As String
Dim horas, minutos As Double
celula = URLEncode(celula)

'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!

Select Case Tipo
Case 1
URI = CelulaAPI & "/" & celula & "/documentos/dados?tipo=1"
Case 2
termo = URLEncode(termo)
URI = CelulaAPI & "/" & celula & "/documentos/dados?tipo=2?consulta=" & termo & ""
Case 3
URI = CelulaAPI & "/" & celula & "/documentos/dados?tipo=3"
Case 4
termo = URLEncode(termo)
URI = CelulaAPI & "/" & celula & "/documentos/dados?tipo=4?consulta=" & termo & ""
End Select
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json;charset=UTF-8"
oReq.Send
'oReq.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'oReq.SetRequestHeader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'oReq.SetRequestHeader "Accept-Language", "en-us,en;q=0.5"
'oReq.SetRequestHeader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"


Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)

QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Não há documentos encontrados para esta célula", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object
If Tipo = 1 Then
ReDim resposta(4, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("clienteNome")
resposta(1, aux) = obj("documentoNome")
resposta(2, aux) = obj("prazoMaximoAnalise")
resposta(3, aux) = obj("tipo")
resposta(4, aux) = obj("complexidade")
aux = aux + 1
Next obj
ElseIf Tipo = 2 Then
ReDim resposta(4, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("clienteNome")
resposta(1, aux) = obj("documentoNome")
resposta(2, aux) = obj("prazoMaximoAnalise")
resposta(3, aux) = obj("tipo")
resposta(4, aux) = obj("complexidade")
aux = aux + 1
Next obj
ElseIf Tipo = 3 Then
ReDim resposta(8, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("documentoId")
resposta(1, aux) = obj("clienteNome")
resposta(2, aux) = obj("documentoNome")
resposta(3, aux) = obj("prazoMaximoAnalise")
resposta(4, aux) = obj("tipo")
resposta(5, aux) = obj("complexidade")
resposta(6, aux) = obj("clienteTipo")
resposta(7, aux) = obj("clienteId")
Tempo = Application.WorksheetFunction.Text(obj("tempoMedioAnalise")("minutes") & ":" & obj("tempoMedioAnalise")("seconds"), "hh:mm")
resposta(8, aux) = Tempo
aux = aux + 1
Next obj
Else
ReDim resposta(5, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("documentoId")
resposta(1, aux) = obj("clienteNome")
resposta(2, aux) = obj("documentoNome")
resposta(3, aux) = obj("prazoMaximoAnalise")
resposta(4, aux) = obj("tipo")
resposta(5, aux) = obj("complexidade")
aux = aux + 1
Next obj
End If
GetCelulaPrazoAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
End Function


Public Function GetColaboradorDadosResumidosAPI(celula As String, Tipo As Integer, Optional termo As String)
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI As String
Dim JSON As String
celula = URLEncode(celula)
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!

Select Case Tipo
'getcolaboradoresinfo_tipo1
Case 1
URI = CelulaAPI & "/" & celula & "/analistas/dados-resumidos?tipo=1"
'sp_getcolaboradoresinfo_tipo_Nome
Case 2
termo = URLEncode(termo)
URI = CelulaAPI & "/" & celula & "/analistas/dados-resumidos?tipo=2?termo=" & termo
'getcolaboradoresinfo_tipo_Email
Case 3
termo = URLEncode(termo)
URI = CelulaAPI & "/" & celula & "/analistas/dados-resumidos?tipo=3?termo=" & termo
'_colaboradorcargocomplexidade
Case 4
URI = CelulaAPI & "/" & celula & "/analistas/dados-resumidos?tipo=4"
'GetLiderEmail
Case 5
URI = CelulaAPI & "/" & celula & "/analistas/dados-resumidos?tipo=5"
End Select
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json;charset=UTF-8"
'oReq.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'oReq.SetRequestHeader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'oReq.SetRequestHeader "Accept-Language", "en-us,en;q=0.5"
'oReq.SetRequestHeader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"
oReq.Send

Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)

QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Não há colaboradores para esta célula", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object
If Tipo = 1 Then
ReDim resposta(1, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nome")
resposta(1, aux) = obj("email")
aux = aux + 1
Next obj
ElseIf Tipo = 2 Then
ReDim resposta(4, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nomeAnalista")
resposta(1, aux) = obj("email")
resposta(2, aux) = obj("cargo")
resposta(3, aux) = obj("lideranca")
resposta(4, aux) = obj("nomeCelula")
aux = aux + 1
Next obj
ElseIf Tipo = 3 Then
ReDim resposta(5, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nomeAnalista")
resposta(1, aux) = obj("email")
resposta(2, aux) = obj("cargo")
resposta(3, aux) = obj("lideranca")
resposta(4, aux) = obj("nomeCelula")
aux = aux + 1
Next obj
ElseIf Tipo = 4 Then
ReDim resposta(1, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nome")
resposta(1, aux) = obj("cargoComplexidade")
aux = aux + 1
Next obj
Else
ReDim resposta(1, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("nome")
resposta(1, aux) = obj("email")
aux = aux + 1
Next obj
End If
GetColaboradorDadosResumidosAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function
Public Function GetAnalistasDadosCompletosAPI(celula As String, idcelula As Integer) As Variant
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim URI As String
celula = URLEncode(celula)
URI = CelulaAPI & "/" & idcelula & "/" & celula & "/analistas/dados"
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "GET", URI, False
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
'oReq.setrequestheader "Connection", "Keep-Alive"
'oReq.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'oReq.setrequestheader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
'oReq.setrequestheader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"
'oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send

Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
'MsgBox oReq.ResponseText
QtdResultados = objJSON.Count
If objJSON.Count = 0 Then
MsgBox "Não há clientes para esta célula", vbInformation
Exit Function
End If
QtdResultados = QtdResultados - 1
Dim obj As Object
ReDim resposta(6, 0 To QtdResultados)
For Each obj In objJSON
resposta(0, aux) = obj("analistaId")
resposta(1, aux) = obj("nome")
resposta(2, aux) = obj("email")
resposta(3, aux) = obj("cargo")
resposta(4, aux) = obj("cargoComplexidade")
resposta(5, aux) = obj("lideranca")
resposta(6, aux) = obj("eliderenca")
aux = aux + 1
Next obj
GetAnalistasDadosCompletosAPI = resposta
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
End Function

Public Function PostClienteAPI(Cliente As CCliente) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim url As String
Dim obj As Object
Dim objJSON As Object
Dim JSON As Object
Dim NovoCliente As String
Cliente.Nome = URLEncode(Cliente.Nome)
Cliente.Nome = Chr$(34) & Cliente.Nome & Chr$(34)
NovoCliente = "{""clienteId"":0,""celulaId"":" & Cliente.CelulaID & " ,""nome"":" & Cliente.Nome & ",""tipo"":" & Cliente.Tipo & "}"
celula = URLEncode(Cliente.CelulaNome)
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "POST", CelulaAPI & "/" & celula & "/cliente/novo", False
NovoCliente = JsonConverter.ConvertToJson(NovoCliente)
NovoCliente = Mid(NovoCliente, 2, Len(NovoCliente) - 2)
NovoCliente = Replace(NovoCliente, "\", "")
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (NovoCliente)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("id") > 0 Then
MsgBox "Cliente inserido com sucesso! ", vbInformation
Exit Function
Else
MsgBox "Não foi possível inserir o cliente", vbCritical
Exit Function
End If
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function
Public Function PatchClienteAPI(Cliente As CCliente) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim url As String
Dim obj As Object
Dim objJSON As Object
Dim JSON As Object
Dim ClienteAlterado As String
Cliente.Nome = URLEncode(Cliente.Nome)
Cliente.Nome = Chr$(34) & Cliente.Nome & Chr$(34)
ClienteAlterado = "{""clienteId"":" & Cliente.ClienteID & ",""celulaId"":" & Cliente.CelulaID & " ,""nome"":" & Cliente.Nome & ",""tipo"":" & Cliente.Tipo & "}"
celula = URLEncode(Cliente.CelulaNome)
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "PATCH", "https://apmsdocapiprd.azurewebsites.net/celula/" & Cliente.CelulaID & "/" & celula & "/cliente/alterar/" & Cliente.ClienteID, False
ClienteAlterado = JsonConverter.ConvertToJson(ClienteAlterado)
ClienteAlterado = Mid(ClienteAlterado, 2, Len(ClienteAlterado) - 2)
ClienteAlterado = Replace(ClienteAlterado, "\", "")
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (ClienteAlterado)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("registroAlterado") > 0 Then
MsgBox "Cliente alterado com sucesso! ", vbInformation
Exit Function
End If
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function
Public Function PostAnalistaAPI(Analista As CAnalista) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim url As String
Dim obj As Object
Dim objJSON As Object
Dim JSON As Object
Dim NovoAnalista As String
Analista.Nome = URLEncode(Analista.Nome)
Analista.Nome = Chr$(34) & Analista.Nome & Chr$(34)
Analista.Email = URLEncode(Analista.Email)
Analista.Email = Chr$(34) & Analista.Email & Chr$(34)
Analista.Liderenca = URLEncode("asdsa")
Analista.Liderenca = Chr$(34) & Analista.Liderenca & Chr$(34)
Dim lider As String
lider = IIf(Analista.lider = False, " false", " true")
NovoAnalista = "{""analistaId"":0,""celulaId"":" & Analista.CelulaID & " ,""Nome"":" & Analista.Nome & "," _
& """Cargo"":" & Analista.Cargo & ",""eLiderenca"":" & lider & ",""lideranca"":" & Analista.Liderenca & "," _
& """Email"":" & Analista.Email & ",""CargoComplexidade"":" & Analista.CargoComplexidade & ",""eLocal"": true}"
celula = URLEncode(Analista.CelulaNome)
'ATENÇÃO. QUANDO A API FOR PUBLICADA, ALTERE AS URIs!

oReq.Open "POST", "https://apmsdocapiprd.azurewebsites.net/celula/" & Analista.CelulaID & "/" & celula & "/analista/novo", False
NovoAnalista = JsonConverter.ConvertToJson(NovoAnalista)

NovoAnalista = Mid(NovoAnalista, 2, Len(NovoAnalista) - 2)
NovoAnalista = Replace(NovoAnalista, "\", "")
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (NovoAnalista)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("analistaId") > 0 Then
MsgBox "Analista inserido com sucesso! ", vbInformation
Exit Function
End If
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function
Public Function PatchAnalistaAPI(Analista As CAnalista) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim url As String
Dim obj As Object
Dim objJSON As Object
Dim JSON As Object
Dim AnalistaAlterado As String
Analista.Nome = URLEncode(Analista.Nome)
Analista.Nome = Chr$(34) & Analista.Nome & Chr$(34)
Analista.Email = URLEncode(Analista.Email)
Analista.Email = Chr$(34) & Analista.Email & Chr$(34)
Analista.Liderenca = URLEncode("")
Analista.Liderenca = Chr$(34) & Analista.Liderenca & Chr$(34)
Dim lider As String
lider = IIf(Analista.lider = False, " false", " true")
AnalistaAlterado = "{""analistaId"":" & Analista.AnalistaID & ",""celulaId"":" & Analista.CelulaID & " ,""Nome"":" & Analista.Nome & "," _
& """Cargo"":" & Analista.Cargo & ",""eLiderenca"":" & lider & ",""lideranca"":" & Analista.Liderenca & "," _
& """Email"":" & Analista.Email & ",""CargoComplexidade"":" & Analista.CargoComplexidade & ",""eLocal"": true}"
celula = URLEncode(Analista.CelulaNome)
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "PATCH", "https://apmsdocapiprd.azurewebsites.net/celula/" & Analista.CelulaID & "/" & celula & "/analista/alterar/" & Analista.AnalistaID, False
AnalistaAlterado = JsonConverter.ConvertToJson(AnalistaAlterado)
AnalistaAlterado = Mid(AnalistaAlterado, 2, Len(AnalistaAlterado) - 2)
AnalistaAlterado = Replace(AnalistaAlterado, "\", "")
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (AnalistaAlterado)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("registroAlterado") > 0 Then
MsgBox "Analista alterado com sucesso! ", vbInformation
Exit Function
Else
MsgBox "Não foi possível inserir o cliente", vbCritical
Exit Function
End If
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function
Public Function EditDocumentoAPI(ByVal DocumentosACorrigir As Dictionary, celula As String, RAVColunas As cRAVColunasXL) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux, QtdDocumentos, DocumentoAtual As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim DocumentoASolicitar As New Dictionary
Dim DocumentoAdicionado As New Dictionary
QtdDocumentos = DocumentosACorrigir.Count - 1
Dim DocumentoAusenteAtual, RespostaConsulta As Variant
Dim Documento As CDocumento
Dim LinhasAfetadas, LinhasRespostaConsulta As Integer
Dim URI, Cliente, Documentos As String
Dim objJSON As Object
celula = URLEncode(celula)
'ATENÇÃO. QUANDO A API FOR PUBLICADA, ALTERE AS URIs!
For DocumentoAtual = 0 To QtdDocumentos
Cliente = CStr(DocumentosACorrigir.Items(DocumentoAtual))
Documentos = CStr(DocumentosACorrigir.Keys(DocumentoAtual))
URI = "https://apmsdocapiprd.azurewebsites.net/celula/" & RAVPreferencias(1).CelulaID & "/" & celula & "/cliente/" & Cliente & "/documento/alterar/" & Documentos
oReq.Open "POST", CelulaAPI & "/" & RAVPreferencias(1).CelulaID & "/" & celula & "/documento/alterar/", False
Documentos = "{""nome"":" & Chr$(34) & Documentos & Chr$(34) & ",""cliente"":" & Chr$(34) & Cliente & Chr$(34) & "}"

oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (Documentos)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("tipo") = 0 Or objJSON("tipo") = 1 Then
LinhasAfetadas = LinhasAfetadas + 1
ElseIf objJSON("tipo") = 2 Then
DocumentoASolicitar.Add DocumentosACorrigir.Keys(DocumentoAtual), DocumentosACorrigir.Items(DocumentoAtual)
Else
    
End If
Next DocumentoAtual


If DocumentoASolicitar.Count = DocumentosACorrigir.Count Then

resposta = MsgBox("Nenhum documento foi encontrado em nossa base. Deseja inserir manualmente à base ? " _
& "São " & DocumentoASolicitar.Count & " documento(s)", vbQuestion + vbYesNo)
If resposta = vbYes Then

Set Documento = New CDocumento
RespostaConsulta = GetCelulaPrazoDocumentosAPI(RAVPreferencias(1).celula, RAVPreferencias(1).CelulaID)
LinhasRespostaConsulta = UBound(RespostaConsulta)

If LinhasRespostaConsulta > 0 Then

For Each DocumentoAusenteAtual In DocumentoASolicitar.Keys
Documento.Nome = DocumentoAusenteAtual
Documento.Complexidade = 0
Documento.PrazoMaximoAnalise = RespostaConsulta
Documento.Tipo = ""
Documento.TempoMedioAnalise = TimeValue("00:02:00")
Documento.celula = RAVPreferencias(1).celula
'RAVPreferencias(1).Cliente = DocumentoASolicitar.Item(DocumentoAusenteAtual)
Documento.Cliente = DocumentoASolicitar.Item(DocumentoAusenteAtual)
DocumentoNovo.Add Documento
Set Documento = New CDocumento
FrmCadastroDocumento.Show
Next DocumentoAusenteAtual
If NovoRegistro > 0 Then
AtualizaRepositorio PastadeTrabalhoRAV, RAVColunas
End If
Else
MsgBox "Como a consulta não gerou resultados, Será enviado o e-mail com a solicitações pertinentes.", vbExclamation
EmailSolicitacaoInfoDocumentos DocumentoASolicitar, celula
End If

Else
EmailSolicitacaoInfoDocumentos DocumentoASolicitar, celula
End If

ElseIf DocumentoASolicitar.Count > 0 Then
resposta = MsgBox("Conseguimos encontrar ou inserir novo(s) documento(s) em nossa base, mas não obtemos sucesso para outro(s)." _
& "Deseja inserir manualmente à base ? São " & DocumentoASolicitar.Count & " documento(s)", vbQuestion + vbYesNo)
If resposta = vbYes Then
Set Documento = New CDocumento

RespostaConsulta = GetCelulaPrazoDocumentosAPI(RAVPreferencias(1).celula)
LinhasRespostaConsulta = UBound(RespostaConsulta)

If LinhasRespostaConsulta > 0 Then

For Each DocumentoAusenteAtual In DocumentoASolicitar.Keys
Documento.Nome = DocumentoAusenteAtual
Documento.Complexidade = 0
Documento.PrazoMaximoAnalise = RespostaConsulta
Documento.Tipo = ""
Documento.TempoMedioAnalise = TimeValue("00:02:00")
'RAVPreferencias(1).Cliente = DocumentoASolicitar.Item(DocumentoAusenteAtual)
Documento.celula = RAVPreferencias(1).celula
Documento.Cliente = DocumentoASolicitar.Item(DocumentoAusenteAtual)
DocumentoNovo.Add Documento
FrmCadastroDocumento.Show
Next DocumentoAusenteAtual
Set DocumentoNovo = Nothing
If NovoRegistro > 0 Then
AtualizaRepositorio PastadeTrabalhoRAV, RAVColunas
End If
'AtualizaRepositorio PastadeTrabalhoRAV, RAVColunas
Else
MsgBox "Como a consulta não gerou resultados, Será enviado o e-mail com a solicitações pertinentes.", vbExclamation
EmailSolicitacaoInfoDocumentos DocumentoASolicitar, celula
End If

Else
EmailSolicitacaoInfoDocumentos DocumentoASolicitar, celula
End If
Else
MsgBox "Documento ausente inserido automaticamente com sucesso!", vbInformation
AtualizaRepositorio PastadeTrabalhoRAV, RAVColunas
End If
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function
Public Function PostDocumentoAPI(Documento As CDocumento) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim Data As String
Dim obj As Object
Dim objJSON As Object
Dim JSON As Object
Dim NovoDocumento As String
'documento.Nome = URLEncode(documento.Nome)
Documento.Nome = Chr$(34) & Documento.Nome & Chr$(34)
Cliente = Chr$(34) & Documento.Cliente & Chr$(34)
'cliente = Chr$(34) & cliente & Chr$(34)
celula = URLEncode(Documento.celula)
Data = Now + 1
Data = Format(Data, "yyyy-mm-dd") + "T" + Format(Documento.TempoMedioAnalise, "hh:mm:ss") + ".000Z"
'data = URLEncode(data)
Data = Chr$(34) & Data & Chr$(34)
Dim Tipo As Integer
Tipo = IIf(Documento.Tipo = "0", 0, 1)
If Cliente = """" Then
NovoDocumento = "{""documentoId"":0,""slaId"":" & Documento.SLAID & ",""nome"":" & Documento.Nome & "," _
& """prazoMaximoAnalise"":" & Documento.PrazoMaximoAnalise & ",""tipo"":" & Tipo & ",""tempoMedioAnaliseBruto"":" & Data & "," _
& """complexidade"":" & Documento.Complexidade & "}"
Else
NovoDocumento = "{""documentoId"":0,""slaId"":" & Documento.SLAID & ",""nome"":" & Documento.Nome & "," _
& """prazoMaximoAnalise"":" & Documento.PrazoMaximoAnalise & ",""tipo"":" & Tipo & ",""tempoMedioAnaliseBruto"":" & Data & "," _
& """complexidade"":" & Documento.Complexidade & ",""cliente"":" & Cliente & " }"
End If
'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "POST", CelulaAPI & "/" & Documento.CelulaID & "/" & celula & "/documento/novo", False
'NovoDocumento = JsonConverter.ConvertToJson(NovoDocumento)
'NovoDocumento = Mid(NovoDocumento, 2, Len(NovoDocumento) - 2)
'NovoDocumento = Replace(NovoDocumento, "\", "")
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (NovoDocumento)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("documentoNovo") > 0 Then
MsgBox "Documento inserido com sucesso! ", vbInformation
NovoRegistro = 1
PostDocumentoAPI = 1
Exit Function
Else
MsgBox "Não foi possível inserir o cliente", vbCritical
Exit Function
End If
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function

Public Function PatchDocumentoAPI(Documento As CDocumento, DocumentoID As Integer) As Integer
Dim oReq As Object
Dim resposta As Variant
Dim QtdResultados, aux As Integer
DoEvents
On Error GoTo FecharAplicacao
Set oReq = CreateObject("MSXML2.ServerXMLHTTP")
Dim Data As String
Dim obj As Object
Dim objJSON As Object
Dim JSON As Object
Dim DocumentoAlterado As String
'documento.Nome = URLEncode(documento.Nome)
Documento.Nome = Chr$(34) & Documento.Nome & Chr$(34)
'cliente = URLEncode(cliente)
'cliente = Chr$(34) & cliente & Chr$(34)
celula = URLEncode(Documento.celula)
Data = Now + 1
Data = Format(Data, "yyyy-mm-dd") + "T" + Format(Documento.TempoMedioAnalise, "hh:mm:ss") + ".000Z"
Data = Chr$(34) & Data & Chr$(34)
Dim Tipo As Integer
Tipo = IIf(Documento.Tipo = "COMUM", 0, 1)
DocumentoAlterado = "{""documentoId"":" & DocumentoID & ",""slaId"":0,""nome"":" & Documento.Nome & "," _
& """prazoMaximoAnalise"":" & Documento.PrazoMaximoAnalise & ",""tipo"":" & Tipo & ",""tempoMedioAnaliseBruto"":" & Data & "," _
& """complexidade"":" & Documento.Complexidade & "}"

'ATENÇÃO. QUANDO A API FOR PUBLICADA ALTERE AS URLs!
oReq.Open "PATCH", CelulaAPI & "/" & Documento.CelulaID & "/" & celula & "/documento/alterar/" & DocumentoID, False
'NovoDocumento = JsonConverter.ConvertToJson(NovoDocumento)
'NovoDocumento = Mid(NovoDocumento, 2, Len(NovoDocumento) - 2)
'NovoDocumento = Replace(NovoDocumento, "\", "")
oReq.setrequestheader "Content-type", "application/json; charset=utf-8"
oReq.setrequestheader "Connection", "Keep-Alive"
oReq.setrequestheader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
oReq.setrequestheader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
oReq.setrequestheader "Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3"
oReq.setrequestheader "Accept-Charset", "utf-8;q=0.7,*;q=0.7"
oReq.setrequestheader "Access-Control-Max-Age", "0"
oReq.Send (DocumentoAlterado)
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
If objJSON("documentosAlterados") > 0 Then
MsgBox "Documento alterado com sucesso! ", vbInformation
NovoRegistro = 1
PatchDocumentoAPI = 1
Exit Function
Else
MsgBox "Não foi possível alterar o documento", vbCritical
Exit Function
End If
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br. " & Err.Description, vbCritical + vbOKOnly
End Function

