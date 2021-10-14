Attribute VB_Name = "Interop"
Option Explicit
'Gera a Interoperabilidade entre sistemas. Integra APIs e Web Services à aplicação.
Public FeriadosHomonimos As New Dictionary
Public Function APIFeriados(DepositoMaisAntigo As Date) As Dictionary
Dim Feriados As New Dictionary
Dim Ano, AnoAtual As String
Dim DataFeriado As Date
Dim AnoDeposito As Long
Ano = Year(AnoDeposito)
AnoDeposito = Year(DepositoMaisAntigo)
Ano = Year(Now())
Dim oReq As Object
DoEvents
On Error GoTo FecharAplicacao
For AnoDeposito = AnoDeposito To Ano
Set oReq = CreateObject("Microsoft.XMLHTTP")
oReq.Open "GET", "https://api.calendario.com.br/?json=true&ano=" & AnoDeposito & "&ibge=3106200&token=b2JpbmFzYXJtQGdtYWlsLmNvbSZoYXNoPTEzOTY2NTYwMw", False
oReq.setrequestheader "Content-type", "application/json"
oReq.Send

Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
Dim o As Object
For Each o In objJSON
DataFeriado = CDate(o("date"))
If DataFeriado >= DepositoMaisAntigo Then
If Not Feriados.Exists(DataFeriado) Then
Feriados.Add DataFeriado, o("name")
Else
FeriadosHomonimos.Add DataFeriado, o("name")
End If
End If
Next o
Next AnoDeposito
Set APIFeriados = Feriados
Exit Function
FecharAplicacao:
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
End Function


