Attribute VB_Name = "Util"
Option Explicit
Public Cliente As String
Public celula As String
Public Planilha As Excel.Workbook

Public Function Init(ClienteSelecionado As String, CelulaSelecionada As String, Planilha_Aberta_Editar As Excel.Workbook)
Cliente = ClienteSelecionado
celula = CelulaSelecionada
Set Planilha = Planilha_Aberta_Editar
Dim Sessao(2) As Variant
Sessao(0) = ClienteSelecionado
Sessao(1) = CelulaSelecionada
Sessao(2) = Planilha
End Function

Public Function IsInit(ClienteSelecionado As String, CelulaSelecionada As String, Planilha_Aberta_Editar As Excel.Workbook) As Boolean
If ClienteSelecionado <> "" And CelulaSelecionada <> "" And Planilha_Aberta_Editar.FullName <> "" Then
IsInit = True
Else
IsInit = False
End If
End Function

Public Function GetFeriados() As Dictionary
Dim Feriados As New Dictionary
Dim Ano As String
Ano = Year(Now())
Dim oReq As Object
DoEvents
On Error Resume Next
Set oReq = CreateObject("Microsoft.XMLHTTP")

If Err <> 0 Then
MsgBox "Não foi possível buscar os feriados do Serviço Web. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
Exit Function
End If
oReq.Open "GET", "https://api.calendario.com.br/?json=true&ano=" & Ano & "&ibge=3106200&token=b2JpbmFzYXJtQGdtYWlsLmNvbSZoYXNoPTEzOTY2NTYwMw", False

If Err <> 0 Then
MsgBox "Não foi possível buscar os feriados do Serviço Web. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
Exit Function
End If

oReq.setrequestheader "Content-type", "application/json"
oReq.Send

If Err <> 0 Then
MsgBox "Não foi possível buscar os feriados da API solicitada. Favor tentar mais tarde ou contacte o desenvolvedor." _
& "E-mail: weverson.rafael@demarco.com.br." & Err, vbCritical + vbOKOnly
Exit Function
End If

Dim objJSON As Object
Set objJSON = JsonConverter.ParseJson(oReq.ResponseText)
Dim o As Object
For Each o In objJSON
Feriados.Add o("date"), o("name")
Next o
Set GetFeriados = Feriados
End Function

Public Function FeriadoContextual(ByVal Feriado As Dictionary, PastadeTrabalhoRAV As Excel.Workbook) As Dictionary
 
Dim DocumentoDataDepositado, Hoje As Date
Dim DataDepositoMaisAntigo As Date
Dim FeriadoAtual As Integer
Dim CicloEstaCompleto As Boolean
Dim FeriadosTrabalhados As New Dictionary
Hoje = Format(Now(), "dd/mm/yyyy")
Dim IntervaloAbaColunas As Excel.Worksheet
Set IntervaloAbaColunas = PastadeTrabalhoRAV.Worksheets(1)
Dim IntervaloAba As Excel.Range
Dim Dividas, DataDeposito As Excel.Range
Dim LinhaInicial, LinhaFinal As Long
Dim ColunaDivida, ColunaDataDeposito As Excel.Range

Set DataDeposito = IntervaloAbaColunas.Range("J:J").Rows.SpecialCells(xlCellTypeVisible)
Set Dividas = IntervaloAbaColunas.Range("L:L").Rows.SpecialCells(xlCellTypeVisible)
LinhaInicial = IntervaloAbaColunas.Range("J:J").SpecialCells(xlCellTypeVisible).Rows(1).Row
LinhaFinal = IntervaloAbaColunas.Range("J:J").SpecialCells(xlCellTypeVisible).End(xlDown).Row
Set ColunaDataDeposito = IntervaloAbaColunas.Range("J" & LinhaInicial & ":" & "J" & LinhaFinal)
Set ColunaDivida = Range("L" & LinhaInicial & ":" & "L" & LinhaFinal)

DataDepositoMaisAntigo = Format(IntervaloAbaColunas.Application.WorksheetFunction.MinIfs _
(IntervaloAbaColunas.Range("J" & LinhaInicial & ":" & "J" & LinhaFinal), _
IntervaloAbaColunas.Range("M" & LinhaInicial & ":" & "M" & LinhaFinal), "Não"), "dd/mm/yyyy")

For DocumentoDataDepositado = DataDepositoMaisAntigo To Hoje

If Feriado.Keys(FeriadoAtual) >= DataDepositoMaisAntigo And Feriado.Keys(FeriadoAtual) <= Hoje Then
If Not FeriadosTrabalhados.Exists(FeriadoAtual) And CicloEstaCompleto = False Then
FeriadosTrabalhados.Add Feriado.Keys(FeriadoAtual), Feriado.Items(FeriadoAtual)
Else
'É possível que esta linha venha a dar erro, em razão
GoTo Saida
End If
End If


FeriadoAtual = FeriadoAtual + 1

If FeriadoAtual >= Feriado.Count Then
CicloEstaCompleto = True
FeriadoAtual = 0
End If

Next DocumentoDataDepositado
Saida:
Set FeriadoContextual = FeriadosTrabalhados
End Function

Public Function EmailSolicitacaoInfoDocumentos(ByVal Documentos As Dictionary, celula As String)

Dim Email As Outlook.Application
Set Email = New Outlook.Application
Dim Novo_Email As Outlook.MailItem
Dim Hora As Date: Hora = Format(Now(), "hh:mm")
Dim MomentodoDia As String
Dim DocumentosSolicitados As String
DocumentosSolicitados = "<ol>" & vbCr
DocumentosSolicitados = "<style>table,th,td" _
& "{padding: 10px;border: 1px solid black;border-collapse: collapse;}</style><table><tr><th>Documento</th><th>Cliente</th></tr>"
Dim QtdElementos, ElementoAtual As Integer
QtdElementos = Documentos.Count - 1
Dim InfoLideres As Variant
InfoLideres = GetColaboradorDadosResumidosAPI(RAVPreferencias(1).celula, 5)
Dim QtdInfoLideres As Integer: QtdInfoLideres = UBound(InfoLideres, 2)
Dim InfoLideresAtual As Integer
Dim EmailLideres, NomeLideres As String

Dim chave As Variant
Dim Lideres As String
Dim QtdElementosAnalista As Integer
Dim EmailAtual As Integer
Dim Hoje As String: Hoje = Format(Date, "dd/mm/yyyy")

Set Novo_Email = Email.CreateItem(olMailItem)

For InfoLideresAtual = 0 To QtdInfoLideres
If InfoLideresAtual <> QtdInfoLideres Then
NomeLideres = NomeLideres + InfoLideres(0, InfoLideresAtual) & ", "
EmailLideres = EmailLideres + InfoLideres(1, InfoLideresAtual) & ";"
Else
NomeLideres = NomeLideres + InfoLideres(0, InfoLideresAtual)
EmailLideres = EmailLideres + InfoLideres(1, InfoLideresAtual)
End If
Next InfoLideresAtual

Novo_Email.To = EmailLideres

Novo_Email.CC = "weverson.rafael@demarco.com.br"

Novo_Email.Subject = "Informações sobre documentos - EMAIL AUTOMÁTICO - DOC, Organizador de Células"

If Hora >= "06:00" And Hora < "12:00" Then
MomentodoDia = "bom dia!"
ElseIf Hora >= "12:00" And Hora < "18:00" Then
MomentodoDia = "boa tarde!"
Else
MomentodoDia = "boa noite!"
End If

For Each chave In Documentos.Keys
If ElementoAtual <> QtdElementos Then
DocumentosSolicitados = "<tr>" & DocumentosSolicitados + "<td>" & chave & "</td>" & "<td>" & Documentos.Item(chave) & "</td></tr>"
ElementoAtual = ElementoAtual + 1
Else
DocumentosSolicitados = "<tr>" & DocumentosSolicitados + "<td>" & chave & "</td>" & "<td>" & Documentos.Item(chave) & "</td></tr></table>"
End If
Next chave
Novo_Email.HTMLBody = "<h3 style=text-align:center;>E-MAIL AUTOMÁTICO - SOLICITAÇÃO DE INFORMAÇÕES DE DOCUMENTO</h3> </br></br>" _
& "Olá, " & MomentodoDia & "<br><br>" _
& "Este E-mail, é um E-mail automatizado, em que buscamos aumentar a qualidade e eficiência do sistema DOC, Organizador de Células.<br><br> " _
& "Para isso, precisamos de saber algumas informações sobre o(s) seguinte(s) documento(s) <br><br><br>" _
& DocumentosSolicitados & " <br>" _
& "Sobre os documentos referidos, precisamos saber:" _
& "<ol><li>O documento é de admissão ou de empresa?</li>" _
& "<li>Complexidade</li>" _
& "<li>Se possível, o tempo de análise médio</li></ol>" _
& "Agradecemos desde já sua cooperação, juntos, facilitaremos a vida de todos analistas no bancodoc." & "<br><br>" _
& "Está com alguma dúvida do porquê estar recebendo este e-mail ? Basta responder este e-mail, e em um dia útil responderemos a sua dúvida!"
Novo_Email.Display

End Function

Public Function FiltrarUmaColuna(PastadeTrabalhoRAV As Excel.Workbook, Coluna As Excel.Range, _
Criterio As String, ColunaNumero As Integer) As Boolean
DoEvents
Dim PlanilhaAtiva As Excel.Worksheet
Set PlanilhaAtiva = PastadeTrabalhoRAV.Worksheets(1)
Dim LinhasFiltradas As Excel.Range
Dim NumColuna As Integer
NumColuna = Coluna.Column
PlanilhaAtiva.Range(Coluna.Address).SpecialCells(xlCellTypeVisible).AutoFilter Field:=ColunaNumero, Criteria1:=Criterio
If PlanilhaAtiva.UsedRange.SpecialCells(xlVisible).Areas.Count > 1 Then
FiltrarUmaColuna = True
Else
FiltrarUmaColuna = False
End If
End Function

Public Function FiltrarVariasColunas(PastadeTrabalhoRAV As Excel.Workbook, Coluna1 As Excel.Range, Criterio1 As Variant, _
Optional Coluna2 As Excel.Range, Optional Criterio2 As Variant, Optional Coluna3 As Excel.Range, Optional Criterio3 As Variant, _
Optional Coluna4 As Excel.Range, Optional Criterio4 As Variant) As Integer
'Filtra várias colunas

Dim PlanilhaAtiva As Excel.Worksheet
Set PlanilhaAtiva = PastadeTrabalhoRAV.Worksheets(1)
Dim LinhasFiltradas As Excel.Range
Dim NumColuna1, NumColuna2, NumColuna3, NumColuna4 As Integer

NumColuna1 = Coluna1.Column
'NumColuna2 = Coluna2.Column
If Criterio1 = 1 Then
PlanilhaAtiva.Range(Coluna1.Address).SpecialCells(xlVisible).AutoFilter Field:=NumColuna1, _
Criteria1:=Array("Em dia", "Atrasado", "Prazo Fatal"), Operator:=xlFilterValues
'Criteria1:=Array("Atrasado", "Prazo Fatal"), Operator:=xlFilterValues
End If
If PlanilhaAtiva.UsedRange.SpecialCells(xlVisible).Areas.Count > 1 Then
'PlanilhaAtiva.Range(Coluna1.Address).SpecialCells(xlVisible).AutoFilter Field:=NumColuna2, Criteria1:=Criterio2
'If PlanilhaAtiva.UsedRange.SpecialCells(xlVisible).Areas.Count > 1 Then
FiltrarVariasColunas = 1
Else
FiltrarVariasColunas = 0
End If
End Function

Public Function EnvioEmail(Para As Variant, Copia As Variant, Assunto As Variant, _
Mensagem As Variant, Optional CopiaOculta As Variant, Optional Tabela As Variant, _
Optional Lote As Boolean, Optional Anexo As Variant)
'Envia E-mails
Dim Email As Outlook.Application
Set Email = New Outlook.Application
Dim Novo_Email As Outlook.MailItem
Set Novo_Email = Email.CreateItem(olMailItem)
With Novo_Email
.To = Para
.CC = Copia & ";elaine.leocadio@demarco.com.br;joseane.ribeiro@demarco.com.br"
.Subject = Assunto
.HTMLBody = Mensagem
.Save
If Not IsMissing(Anexo) And Anexo <> "" Then
.Attachments.Add (Anexo)
End If
If Lote = True Then
.Display
Else
.Display
End If


End With
End Function


Public Function AtualizaRepositorio(PastadeTrabalhoRAV As Excel.Workbook, RAVColunas As cRAVColunasXL)
'Atualiza o repositorio de documentos

PastadeTrabalhoRAV.Application.ScreenUpdating = False
PastadeTrabalhoRAV.Application.Calculation = xlCalculationManual
PastadeTrabalhoRAV.Application.EnableEvents = False
PastadeTrabalhoRAV.Application.DisplayAlerts = False

Dim PlanilhaRepositorio, PlanilhaAguardandovalidacao As Excel.Worksheet
Set PlanilhaRepositorio = PastadeTrabalhoRAV.Worksheets(2)
Dim ElementoAtual As Integer
Dim NumLinha As Integer
Dim QtdDocumentosConsultados As Long
Dim ValorCelulasRepositorio, Feriados As Variant
Dim EnderecoCelula As Excel.Range
NumLinha = PlanilhaRepositorio.UsedRange.Rows.Count
ValorCelulasRepositorio = PlanilhaRepositorio.Range("A2:D" & Cells(NumLinha, 4).End(xlEnd).Row).Value2
Set PlanilhaAguardandovalidacao = PastadeTrabalhoRAV.Worksheets(2)
Dim QtdLinhasRepositorio As Long
QtdLinhasRepositorio = UBound(ValorCelulasRepositorio, 1) - 1
Dim LinhaAtualRepositorio As Integer
Dim PrimeiraLinha, UltimaLinha, UltimaLinhaRepositorio, aux As Long
Dim DocumentosConsultados As Variant

If RAVPreferencias(1).Cliente = "Todos os clientes" Then
DocumentosConsultados = GetCelulaPrazoAPI(RAVPreferencias(1).celula, 1, "")
Else
DocumentosConsultados = GetCelulaPrazoAPI(RAVPreferencias(1).celula, 1, RAVPreferencias(1).Cliente)
End If

QtdDocumentosConsultados = UBound(DocumentosConsultados, 2)
If QtdLinhasRepositorio < QtdDocumentosConsultados Then
LinhaAtualRepositorio = QtdLinhasRepositorio + 2
For ElementoAtual = QtdLinhasRepositorio To QtdDocumentosConsultados
PlanilhaRepositorio.Cells(LinhaAtualRepositorio, 1).Value2 = DocumentosConsultados(0, QtdLinhasRepositorio)
PlanilhaRepositorio.Cells(LinhaAtualRepositorio, 2).Value2 = DocumentosConsultados(1, QtdLinhasRepositorio)
PlanilhaRepositorio.Cells(LinhaAtualRepositorio, 3).Value2 = DocumentosConsultados(3, QtdLinhasRepositorio)
If DocumentosConsultados(2, QtdLinhasRepositorio) = 1 Then
PlanilhaRepositorio.Cells(LinhaAtualRepositorio, 4).Value2 = "BLOQUEIO"
Else
PlanilhaRepositorio.Cells(LinhaAtualRepositorio, 4).Value2 = "COMUM"
End If
PlanilhaRepositorio.Cells(LinhaAtualRepositorio, 5).Value2 = DocumentosConsultados(4, ElementoAtual)
QtdLinhasRepositorio = QtdLinhasRepositorio + 1
LinhaAtualRepositorio = LinhaAtualRepositorio + 1
Next ElementoAtual
End If

PlanilhaRepositorio.Columns("A:A").AutoFit
PlanilhaRepositorio.Columns("B:B").AutoFit
PlanilhaRepositorio.Columns("C:C").AutoFit
PlanilhaRepositorio.Columns("D:D").AutoFit

Set PlanilhaAguardandovalidacao = PastadeTrabalhoRAV.Worksheets(1)
Dim linhasdq As Excel.Range
Set linhasdq = FiltroPadrao

For Each Feriados In FiltroPadrao.Rows

NumLinha = PlanilhaAguardandovalidacao.Range(Feriados.Address).Row
If PlanilhaAguardandovalidacao.Cells(NumLinha, 22).Value2 <> "" Then
aux = NumLinha + 1
Else
UltimaLinha = FiltroPadrao.SpecialCells(xlCellTypeVisible).End(xlDown).Row
Exit For
End If
Next Feriados



'UltimaLinhaRepositorio = PlanilhaRepositorio.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row

'PlanilhaAguardandovalidacao.Range("AS" & PrimeiraLinha).FormulaArray = _
'"=IFERROR(INDEX(Repositório!$C$2:$C$" & UltimaLinhaRepositorio & ",MATCH(Plan1!B" & PrimeiraLinha & ":B" & UltimaLinha & "" _
'& "&Plan1!E" & PrimeiraLinha & ":E" & UltimaLinha & ",Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" _
'& UltimaLinhaRepositorio & ",0)),""NE"")"
'
'PlanilhaAguardandovalidacao.Range("AP" & PrimeiraLinha).Formula = "=IF(AS" & PrimeiraLinha & "<>""NE"",DAYS(TODAY(),J" & PrimeiraLinha & "),""NE"")"
'
'PlanilhaAguardandovalidacao.Range("AQ" & PrimeiraLinha).Formula = _
'"=IF(AS" & PrimeiraLinha & "=""NE"",""NE"",IF(AND(M" & PrimeiraLinha & "<>""-"",N" & PrimeiraLinha & "=""-""),""-""," _
'& "IF(N" & PrimeiraLinha & "=""-"",CPDOD(J" & PrimeiraLinha & ",$AU$" & PrimeiraLinha & ":$AU$" & Aux & ")," _
'& "IF(N" & PrimeiraLinha & ">J" & PrimeiraLinha & ",CPDOD(N" & PrimeiraLinha & ",$AU$" & PrimeiraLinha & ":$AU$" & Aux & ")," _
'& "CPDOD(J" & PrimeiraLinha & ",$AU$" & PrimeiraLinha & ":$AU$" & Aux & ")))))"
'
'PlanilhaAguardandovalidacao.Range("R" & PrimeiraLinha).FormulaArray = _
'"=IFERROR(INDEX(Repositório!$D$2:$D$" & UltimaLinhaRepositorio & ",MATCH(Plan1!B" & PrimeiraLinha & ":B" & UltimaLinha & "" _
'& "&Plan1!E" & PrimeiraLinha & ":E" & UltimaLinha & ",Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" _
'& UltimaLinhaRepositorio & ",0)),""NE"")"
'
'PlanilhaAguardandovalidacao.Range("T" & PrimeiraLinha).Formula = _
'"=IF(AND(ISNUMBER(AQ" & PrimeiraLinha & "),ISNUMBER(AS" & PrimeiraLinha & "))," _
'& "SWITCH(AR" & PrimeiraLinha & ",""COMUM"",IF(OR(E" & PrimeiraLinha & "=""Acordo Coletivo de Trabalho""," _
'& "E" & PrimeiraLinha & "=""Convenção Coletiva de Trabalho""),IF(AQ" & PrimeiraLinha & ">AS" & PrimeiraLinha & ",""Atrasado""," _
'& "IF(AQ" & PrimeiraLinha & "=AS" & PrimeiraLinha & ",""Prazo Fatal"",""Em Dia"")),IF(AQ" & PrimeiraLinha & ">AS" & PrimeiraLinha & ",""" _
'& "Atrasado"",IF(AQ" & PrimeiraLinha & "=AS" & PrimeiraLinha & ",""Prazo Fatal"",""Em Dia""))),""BLOQUEIO""," _
'& "IF(AQ" & PrimeiraLinha & ">AS" & PrimeiraLinha & ",""Atrasado"",IF(AQ" & PrimeiraLinha & "=AS" & PrimeiraLinha & ",""" _
'& "Prazo Fatal"",""Em Dia""))),IF(AQ" & PrimeiraLinha & "=""-"",""N/A"",""NE""))"
'
'' --------

PrimeiraLinha = FiltroPadrao.Rows.Row
UltimaLinha = PlanilhaAguardandovalidacao.Range("J:J").SpecialCells(xlCellTypeVisible).End(xlDown).Row
UltimaLinhaRepositorio = PlanilhaRepositorio.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row
aux = PrimeiraLinha

PlanilhaAguardandovalidacao.Range(RAVColunas.Feriados & ":" & RAVColunas.Feriados).NumberFormat = "m/dd/yyyy"

PlanilhaAguardandovalidacao.Range(RAVColunas.PrazoMaximoAnalise & PrimeiraLinha).FormulaArray = _
"=IFERROR(INDEX(Repositório!$C$2:$C$" & UltimaLinhaRepositorio & ",MATCH(Plan1!" & RAVColunas.Cliente & PrimeiraLinha & "" _
& ":" & RAVColunas.Cliente & UltimaLinha & "&Plan1!" & RAVColunas.Documento & PrimeiraLinha & ":" & RAVColunas.Documento & UltimaLinha & "," _
& "Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" & UltimaLinhaRepositorio & ",0)),""NE"")"

PlanilhaAguardandovalidacao.Range(RAVColunas.DiasNoSistema & PrimeiraLinha).Formula = "=IF(" & RAVColunas.Tipo & PrimeiraLinha & "<>""NE""," _
& "DAYS(TODAY()," & RAVColunas.DataInclusao & PrimeiraLinha & "),""NE"")"

PlanilhaAguardandovalidacao.Range(RAVColunas.DiasAguardandoAnalise & PrimeiraLinha).Formula = _
"=IF(" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & "=""NE"",""NE"",IF(AND(" & RAVColunas.Inadimpliencia & PrimeiraLinha & "<>""-""," _
& RAVColunas.FimInadimplencia & PrimeiraLinha & "=""-""),""-"",IF(" & RAVColunas.FimInadimplencia & PrimeiraLinha & "=""-""," _
& "CPDOD(" & RAVColunas.DataInclusao & PrimeiraLinha & ",$" & RAVColunas.Feriados & "$" & PrimeiraLinha & ":$" _
& RAVColunas.Feriados & "$" & aux & "),IF(" & RAVColunas.FimInadimplencia & PrimeiraLinha & ">" _
& RAVColunas.DataInclusao & PrimeiraLinha & ",CPDOD(" & RAVColunas.FimInadimplencia & PrimeiraLinha & "," _
& "$" & RAVColunas.Feriados & "$" & PrimeiraLinha & ":$" & RAVColunas.Feriados & "$" & aux & ")," _
& "CPDOD(" & RAVColunas.DataInclusao & PrimeiraLinha & ",$" & RAVColunas.Feriados & "$" & PrimeiraLinha & "" _
& ":$" & RAVColunas.Feriados & "$" & aux & ")))))"

PlanilhaAguardandovalidacao.Range(RAVColunas.Tipo & PrimeiraLinha).FormulaArray = _
"=IFERROR(INDEX(Repositório!$D$2:$D$" & UltimaLinhaRepositorio & ",MATCH(Plan1!" & RAVColunas.Cliente & PrimeiraLinha & "" _
& ":B" & UltimaLinha & "&Plan1!" & RAVColunas.Documento & PrimeiraLinha & ":" & RAVColunas.Feriados & UltimaLinha & "," _
& "Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" & UltimaLinhaRepositorio & ",0)),""NE"")"


PlanilhaAguardandovalidacao.Range(RAVColunas.Status & PrimeiraLinha).Formula = _
"=IF(AND(ISNUMBER(" & RAVColunas.DiasNoSistema & PrimeiraLinha & "),ISNUMBER(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & "))," _
& "SWITCH(" & RAVColunas.Tipo & PrimeiraLinha & ",""COMUM"",IF(OR(" & RAVColunas.Documento & PrimeiraLinha & "=""Acordo Coletivo de Trabalho""," _
& RAVColunas.Documento & PrimeiraLinha & "=""Convenção Coletiva de Trabalho"")," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & ">" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & ",""Atrasado""," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & "=" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & ",""Prazo Fatal"",""Em Dia""))," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & ">" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & ",""Atrasado""," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & "=" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & ",""Prazo Fatal"",""Em Dia""))),""BLOQUEIO""," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & ">" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & ",""Atrasado""," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & "=" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & ",""Prazo Fatal"",""Em Dia"")))," _
& "IF(" & RAVColunas.DiasAguardandoAnalise & PrimeiraLinha & "=""-"",""N/A"",""NE""))"

PlanilhaAguardandovalidacao.Range(RAVColunas.DiasNoSistema & PrimeiraLinha & ":" & RAVColunas.Status & UltimaLinha).FillDown

PlanilhaAguardandovalidacao.Columns(RAVColunas.DiasNoSistema & ":" & RAVColunas.DiasNoSistema).AutoFit
PlanilhaAguardandovalidacao.Columns(RAVColunas.PrazoMaximoAnalise & ":" & RAVColunas.PrazoMaximoAnalise).AutoFit
PlanilhaAguardandovalidacao.Columns(RAVColunas.Tipo & ":" & RAVColunas.Tipo).AutoFit
PlanilhaAguardandovalidacao.Columns(RAVColunas.PrazoMaximoAnalise & ":" & RAVColunas.PrazoMaximoAnalise).AutoFit
PlanilhaAguardandovalidacao.Columns(RAVColunas.Status & ":" & RAVColunas.Status).AutoFit
PlanilhaAguardandovalidacao.Columns(RAVColunas.Feriados & ":" & RAVColunas.Feriados).AutoFit


PastadeTrabalhoRAV.Application.Calculate

Do While Application.CalculationState <> xlDone
     DoEvents
Loop

MsgBox "Documentos e prazos atualizados.", vbInformation

End Function

Public Function GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV As Excel.Workbook, PlanilhaNumero As Integer, _
GerarCopia As Boolean, NumLinha As Variant, RAVColunas As cRAVColunasXL) As Excel.Worksheet
'Retorna o uma tabela temporária com os argumentos solciitados. Útil quando as linhas de uma tabela não são contíguas (sequenciais)

Dim PlanilhaEspelho As Excel.Worksheet
Set PlanilhaEspelho = PastadeTrabalhoRAV.Worksheets(PlanilhaNumero)
Dim PlanilhaTemporia As Excel.Worksheet

If GerarCopia = True Then
PastadeTrabalhoRAV.Sheets.Add after:=PlanilhaEspelho
Set PlanilhaTemporia = PastadeTrabalhoRAV.Worksheets(PlanilhaNumero + 1)
PlanilhaTemporia.Name = "Temp"
PlanilhaEspelho.Range("A1:" & RAVColunas.Linha & NumLinha).SpecialCells(xlCellTypeVisible).Copy
PlanilhaTemporia.Range("A1:" & RAVColunas.Linha & NumLinha).PasteSpecial xlPasteValues
PlanilhaTemporia.Range(RAVColunas.DataInclusao & ":" & RAVColunas.DataInclusao).NumberFormat = "0.00"
    Set GerarCopiaPlanilhaTemporia = PlanilhaTemporia
Else
PastadeTrabalhoRAV.Application.DisplayAlerts = False
Set PlanilhaEspelho = PastadeTrabalhoRAV.Worksheets(PlanilhaNumero + 1)
PlanilhaEspelho.Delete
PastadeTrabalhoRAV.Application.DisplayAlerts = True
Set GerarCopiaPlanilhaTemporia = PlanilhaEspelho
End If
End Function

Public Function CabecalhoTabelaHTML(Tipo As Integer) As String
'Retorna o cabecalho de uma Tabela HTML de documentos de determinado analista
Select Case Tipo
Case 1
CabecalhoTabelaHTML = "<html lang=pt-br><head><style>table, th, td {border: 1px solid gray;}</style></head><body>" & _
        "<table><tr>" & _
        "<th bgcolor=""#bdf0ff"">Protocolo</th>" & _
        "<th bgcolor=""#bdf0ff"">Cliente</th>" & _
        "<th bgcolor=""#bdf0ff"">Fornecedor</th>" & _
        "<th bgcolor=""#bdf0ff"">Unidade</th>" & _
        "<th bgcolor=""#bdf0ff"">Documento</th>" & _
        "<th bgcolor=""#bdf0ff"">Empregado</th>" & _
        "<th bgcolor=""#bdf0ff"">Analista</th>" & _
        "<th bgcolor=""#bdf0ff"">Data de Inclusão</th>" & _
        "<th bgcolor=""#bdf0ff"">Dias em análise</th>" & _
        "<th bgcolor=""#bdf0ff"">Prazo Máximo de Análise</th>" & _
        "<th bgcolor=""#bdf0ff"">Status</th>"
Case 2
CabecalhoTabelaHTML = "<html lang=pt-br><head><style>table, th, td {border: 1px solid gray;}</style></head><body>" & _
        "<table><tr>" & _
        "<th bgcolor=""#bdf0ff"">Protocolo</th>" & _
        "<th bgcolor=""#bdf0ff"">Cliente</th>" & _
        "<th bgcolor=""#bdf0ff"">Fornecedor</th>" & _
        "<th bgcolor=""#bdf0ff"">Unidade</th>" & _
        "<th bgcolor=""#bdf0ff"">Documento</th>" & _
        "<th bgcolor=""#bdf0ff"">Empregado</th>" & _
        "<th bgcolor=""#bdf0ff"">Analista</th>" & _
        "<th bgcolor=""#bdf0ff"">Data de Inclusão</th>" & _
        "<th bgcolor=""#bdf0ff"">Dias em análise</th>" & _
        "<th bgcolor=""#bdf0ff"">Prazo Máximo de Análise</th>" & _
        "<th bgcolor=""#bdf0ff"">Status</th>"
End Select
End Function

Public Function CorpoTabelaHTML(Tipo As Integer, ByVal Tabela As Collection)
'Retorna o corpo de uma Tabela HTML de documentos de determinado analista.
Dim CorpoTabela As String
Dim Item As Variant
Dim Prazos As New Dictionary
Dim CSS As String
Dim ItemAtual As Integer
Dim ItensTabela As Long
Dim ExisteCorpoTabelaHTML As Boolean
ItensTabela = Tabela.Count

Select Case Tipo
Case 1
For Each Item In Tabela
If Item.Status = "Prazo Fatal" Then
CSS = "<tr style=""background-color:#FF0000; color:white"">"
ElseIf Item.Status = "Atrasado" Then
CSS = "<tr style=""background-color:#404040; color:white"" > "
End If
If Tipo = 1 Then
If Item.Status = "Prazo Fatal" Or Item.Status = "Atrasado" Then
ExisteCorpoTabelaHTML = True
 CorpoTabela = CorpoTabela & CSS & ""
        CorpoTabela = CorpoTabela & "<td>" & Item.Protocolo & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Cliente & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Fornecedor & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Unidade & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Documento & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Empregado & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Analista & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.DataInclusao & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.DiasAguardandoAnalise & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.PrazoMaximoAnalise & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Status & "</td>"
 CorpoTabela = CorpoTabela & "</tr>"
 End If
 End If
 Next Item
Case 2
For Each Item In Tabela
If Item.Status = "Prazo Fatal" Then
CSS = "<tr style=""background-color:#FF0000; color:white"">"
ElseIf Item.Status = "Atrasado" Then
CSS = "<tr style=""background-color:#404040; color:white""> "
End If
If Item.Status = "Prazo Fatal" Or Item.Status = "Atrasado" Then
ExisteCorpoTabelaHTML = True
 CorpoTabela = CorpoTabela & CSS & ""
        CorpoTabela = CorpoTabela & "<td>" & Item.Protocolo & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Cliente & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Fornecedor & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Unidade & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Documento & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Empregado & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Analista & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.DataInclusao & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.DiasAguardandoAnalise & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.PrazoMaximoAnalise & "</td>"
        CorpoTabela = CorpoTabela & "<td>" & Item.Status & "</td>"
 CorpoTabela = CorpoTabela & "</tr>"
 End If
  Next Item
 End Select
 If ExisteCorpoTabelaHTML = True Then
 CorpoTabelaHTML = CorpoTabela
 Else
 CorpoTabelaHTML = ""
 End If
End Function

Public Function GerarTabelaDocumentosHTML(Tipo As Integer, ByVal Tabela As Collection)
'Retorna uma tabela em HTML com documentos de determinado analista.
Dim Cabecalho As String: Cabecalho = CabecalhoTabelaHTML(Tipo)
Dim Corpo As String: Corpo = CorpoTabelaHTML(Tipo, Tabela)
If Corpo <> "" Then
GerarTabelaDocumentosHTML = Cabecalho & Corpo & "</table></body></html>"
Else
GerarTabelaDocumentosHTML = "VOCÊ NÃO POSSUI DOCUMENTOS EM PRAZO FATAL E EM PRAZO PERDIDO"
End If
End Function

Public Function SistemaArquivos(Optional Planaliha As Excel.Workbook)

End Function
Public Function CabecalhoTabelaLegendaHTML()
'Retorna o cabecalho de uma Tabela HTML com legenda de cores.
Dim Tabela As String
Tabela = "<style>table, th, td {border: 1px solid gray;}</style></head><body>" & _
        "<table><tr>" & _
        "<th bgcolor=""#bdf0ff"">Cor</th>" & _
        "<th bgcolor=""#bdf0ff"">Dias Em Análise</th>"
End Function
Public Function CorpoTabelaLegendaHTML() As String
'Retorna o corpo de uma Tabela HTML com legenda de cores.
 Dim CSS As String
 Dim CorpoTabela As String
' CorpoTabela = CorpoTabela & CSS & ""
'        CorpoTabela = CorpoTabela & "<td style=""background-color:#FF0000"">" & Tabela(ItemAtual).DepositoID & "</td>"
'        CorpoTabela = CorpoTabela & "<td>" & Tabela(ItemAtual).Cliente & "</td>"
'        CorpoTabela = CorpoTabela & "<td>" & Tabela(ItemAtual).Fornecedor & "</td>"
End Function
Public Function GetValoresColunaExcel(PastadeTrabalhoRAV As Excel.Workbook, NumPlanilha As Integer, _
Coluna1 As String, Optional Coluna2 As String) As Variant
'Retorna coluna(s) em formato de array para busca rápida de valores na tabela solicitada.

Dim PlanilhaSolicitada As Excel.Worksheet
Set PlanilhaSolicitada = PastadeTrabalhoRAV.Worksheets(NumPlanilha)
Dim LinhaInicial, LinhaFinal As Long
LinhaInicial = PlanilhaSolicitada.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
LinhaFinal = PlanilhaSolicitada.UsedRange.SpecialCells(xlCellTypeVisible).End(xlDown).Row
If Coluna2 <> "" Then
GetValoresColunaExcel = PlanilhaSolicitada.Range(Coluna1 & LinhaInicial & ":" & Coluna2 & _
Cells(LinhaFinal, 1).SpecialCells(xlCellTypeVisible).End(xlEnd).Row).Value2
Else
GetValoresColunaExcel = PlanilhaSolicitada.Range(Coluna1 & LinhaInicial & ":" & Coluna1 & _
Cells(LinhaFinal, 1).SpecialCells(xlCellTypeVisible).End(xlEnd).Row).Value2
End If

End Function
Public Function GetColunasExcel(PastadeTrabalhoRAV As Excel.Workbook, NumPlanilha As Integer, _
Coluna1 As String, Coluna2 As String) As Variant
'Retorna coluna(s) em formato de array para busca rápida de valores na tabela solicitada.
Dim PlanilhaSolicitada As Excel.Worksheet
Set PlanilhaSolicitada = PastadeTrabalhoRAV.Worksheets(NumPlanilha)
GetColunasExcel = PlanilhaSolicitada.Range(Coluna1 & 1 & ":" & Coluna2 & _
Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlToRight).Row).Value2
End Function

Public Function GetPrimeiraColunaCriada(PastadeTrabalhoRAV As Excel.Workbook, NumPlanilha As Integer) As String
Dim IndiceColunaCriada As Integer
Dim PlanilhaSolicitada As Excel.Worksheet
Set PlanilhaSolicitada = PastadeTrabalhoRAV.Worksheets(NumPlanilha)
IndiceColunaCriada = 6 - PlanilhaSolicitada.Cells(1, 1).End(xlToRight).Column
GetPrimeiraColunaCriada = Split(Cells(1, IndiceColunaCriada).Address, "$")(1)
End Function
Public Function GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV As Excel.Workbook, NumPlanilha As Integer, _
Optional Coluna As Integer) As String
Dim IndiceColunaDisponivel As Integer
Dim PlanilhaSolicitada As Excel.Worksheet
Set PlanilhaSolicitada = PastadeTrabalhoRAV.Worksheets(NumPlanilha)
If Coluna = 0 Then
IndiceColunaDisponivel = 1 + PlanilhaSolicitada.Cells(1, 1).End(xlToRight).Column
GetPrimeiraColunaDisponvel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)
Else
GetPrimeiraColunaDisponvel = Split(Cells(1, Coluna).Address, "$")(1)
End If
End Function
Public Function GetUltimaColunaPreenchida(PastadeTrabalhoRAV As Excel.Workbook, NumPlanilha As Integer) As String
Dim IndiceColunaPreenchida As Integer
Dim PlanilhaSolicitada As Excel.Worksheet
Set PlanilhaSolicitada = PastadeTrabalhoRAV.Worksheets(NumPlanilha)
IndiceColunaPreenchida = PlanilhaSolicitada.Cells(1, 1).End(xlToRight).Column
GetUltimaColunaPreenchida = Split(Cells(1, IndiceColunaPreenchida).Address, "$")(1)
End Function
Public Function GetColunasIndiceAlfabetico(PastadeTrabalhoRAV As Excel.Workbook) As cRAVColunasXL
Dim RavColuna As cRAVColunasXL
Set RavColuna = New cRAVColunasXL
Dim UltimaColuna, Colunas, Coluna As Variant
Dim ColunaAtual As Integer

UltimaColuna = GetUltimaColunaPreenchida(PastadeTrabalhoRAV, 1)
Colunas = GetColunasExcel(PastadeTrabalhoRAV, 1, "A", CStr(UltimaColuna))
ColunaAtual = 1

For Each Coluna In Colunas
If RavColuna.Protocolo <> "" And RavColuna.Cliente <> "" And RavColuna.Fornecedor <> "" And RavColuna.Documento <> "" _
And RavColuna.DataInclusao <> "" And RavColuna.Empregado <> "" And RavColuna.Divida <> "" And RavColuna.Inadimpliencia <> "" _
And RavColuna.Unidade <> "" And RavColuna.MesDeposito <> "" And RavColuna.FimInadimplencia <> "" And RavColuna.DiasNoSistema <> "" _
And RavColuna.DataInicio <> "" And RavColuna.DiasAguardandoAnalise <> "" And RavColuna.Tipo <> "" And RavColuna.PrazoMaximoAnalise <> "" _
And RavColuna.Status <> "" And RavColuna.Feriados <> "" And RavColuna.JustificativaEmAnalise <> "" And RavColuna.QLP <> "" _
And RavColuna.TempoEmAnalise <> "" And RavColuna.Analista <> "" And RavColuna.DocumentoComplexidade <> "" And RavColuna.Linha <> "" Then
Exit For
Else

Select Case Coluna
    Case "Protocolo"
    RavColuna.Protocolo = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Cliente"
    RavColuna.Cliente = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Fornecedor"
    RavColuna.Fornecedor = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Unidade"
    RavColuna.Unidade = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Documento"
    RavColuna.Documento = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Data de Inclusão"
    RavColuna.DataInclusao = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Empregado"
    RavColuna.Empregado = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Mês do Depósito"
    RavColuna.MesDeposito = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Possui Dívidas"
    RavColuna.Divida = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Início Inadimplência"
    RavColuna.Inadimpliencia = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Fim Inadimplência"
    RavColuna.FimInadimplencia = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Dias no Sistema"
    RavColuna.DiasNoSistema = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Data Início"
    RavColuna.DataInicio = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Dias Aguardando Análise"
    RavColuna.DiasAguardandoAnalise = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Tipo de Documento"
    RavColuna.Tipo = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Prazo Máximo Para Análise"
    RavColuna.PrazoMaximoAnalise = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Status Documento"
    RavColuna.Status = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Feriados Selecionados"
    RavColuna.Feriados = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Justificativa em análise"
    RavColuna.JustificativaEmAnalise = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "QLP"
    RavColuna.QLP = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Tempo em análise (dias)"
    RavColuna.TempoEmAnalise = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Atribuído a:"
    RavColuna.Analista = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Complexidade do Documento"
    RavColuna.DocumentoComplexidade = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    Case "Linha"
    RavColuna.Linha = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1, ColunaAtual)
    End Select
     ColunaAtual = ColunaAtual + 1
End If
Next Coluna
Set GetColunasIndiceAlfabetico = RavColuna
End Function
Public Function GetColunaIndiceNumerico(Coluna As String) As Integer
GetColunaIndiceNumerico = Range(Coluna & 1).Column
End Function
Public Function DistribuirDocumentos(PastadeTrabalhoRAV As Excel.Workbook, Filtro As Excel.Range, _
ByVal DocumentoLinhasPlanRav As Dictionary, RAVColunas As cRAVColunasXL) As String
On Error GoTo Catch:
Dim AnalistaCargoComplexidade, Nivel, NivelAtual, Analistas, Status, DiasEmAnalise, LinhaPlanRav, _
LinhaAtualPlanRav, QtdAnalistaPorNivel, QtdDocumentoPorNivel As Variant
Dim NivelAnalista As New Collection
Dim DocumentosEstagiario As New Collection
Dim DocumentosAuxiliar As New Collection
Dim DocumentosAssitente As New Collection
Dim DocumentosAnalista As New Collection
Dim Preferencias As CRAVPreferencias
Dim ExisteEstagiario As Boolean
Dim ExisteAuxiliar As Boolean
Dim ExisteAssistente As Boolean
Dim ExisteAnalista As Boolean


Dim DocumentoCompEstagiarioAux1 As Boolean
Dim DocumentoCompEstagiarioAux2 As Boolean
Dim DocumentoCompAuxiliarAux1 As Boolean
Dim DocumentoCompAuxiliarAux2 As Boolean
Dim DocumentoCompAuxiliarAux3 As Boolean
Dim DocumentoCompAssitenteAux1 As Boolean
Dim DocumentoCompAssitenteAux2 As Boolean
Dim DocumentoCompAssitenteAux3 As Boolean
Dim DocumentoCompAssitenteAux4 As Boolean
Dim DocumentoCompAnalistaAux1 As Boolean
Dim DocumentoCompAnalistaAux2 As Boolean
Dim DocumentoCompAnalistaAux3 As Boolean
Dim DocumentoCompAnalistaAux4 As Boolean
Dim DocumentoCompAnalistaAux5 As Boolean

DocumentoCompEstagiarioAux1 = False
DocumentoCompEstagiarioAux2 = False
DocumentoCompAuxiliarAux1 = False
DocumentoCompAuxiliarAux2 = False
DocumentoCompAuxiliarAux3 = False
DocumentoCompAssitenteAux1 = False
DocumentoCompAssitenteAux2 = False
DocumentoCompAssitenteAux3 = False
DocumentoCompAssitenteAux4 = False
DocumentoCompAnalistaAux1 = False
DocumentoCompAnalistaAux2 = False
DocumentoCompAnalistaAux3 = False
DocumentoCompAnalistaAux4 = False
DocumentoCompAnalistaAux5 = False


Dim PlanilhaRav As Excel.Worksheet
Dim PlanilhaEspelho As Excel.Worksheet
Dim Intervalo, LinhaSelecionada As Excel.Range
Dim DocumentoAtual, ColunaDisponivel  As String

Dim PrimeiraLinha, UltimaLinha, QtdDocumentosSemDivida, Linha, aux, SegundosInicioDistribuicao, SegundosFimDistribuicao As Long
Dim QtdAnalista, resposta, AnalistaAtual, AnalistasNivel2, AnalistasNivel3, AnalistasNivel4, AnalistasNivel5, _
QtdDocumentoNivel1, QtdDocumentoNivel2, QtdDocumentoNivel3, QtdDocumentoNivel4, QtdDocumentoNivel5, _
QtdDocumentosCompensados1, QtdDocumentosCompensados2, QtdDocumentosCompensados3, QtdDocumentosCompensados4, _
QtdDocumentosCompensados5, QtdDocumentosReservadosNivel1, QtdDocumentosReservadosNivel2, QtdDocumentosReservadosNivel3, _
QtdDocumentosReservadosNivel4, QtdDocumentosReservadosNivel5, _
DocumentosNivel2TotalCompensados, DocumentosNivel3TotalCompensados, Tipo, TipoTingimento, _
DocRestoEstagiarioAux1, DocRestoEstagiarioAux2, DocRestoAuxiliarAux1, DocRestoAuxiliarAux2, _
DocRestoAuxiliarAux3, DocRestoAssitenteAux1, DocRestoAssitenteAux2, DocRestoAssitenteAux3, _
DocRestoAssitenteAux4, DocRestoAnalistaAux1, DocRestoAnalistaAux2, DocRestoAnalistaAux3, _
DocRestoAnalistaAux4, DocRestoAnalistaAux5, DocumentosDistribuidos, QtdLinhasPlanRav, AreasPlanRav, Modo As Integer
Dim QuotaIdeal1, QuotaIdeal2, QuotaIdeal3, QuotaIdeal4, QuotaIdeal5, _
Percentual2, Percentual3, Percentual4, aux1, aux2, aux3, aux4 As Double

Dim DocumentoNivel1 As New Dictionary
Dim DocumentoNivel2 As New Dictionary
Dim DocumentoNivel3 As New Dictionary
Dim DocumentoNivel4 As New Dictionary
Dim DocumentoNivel5 As New Dictionary
'Dim DocumentoLinhasPlanRav As New Dictionary
Dim LinhasPlanRav As New Dictionary
Dim AnalistaNivel1 As New Dictionary
Dim AnalistaNivel2 As New Dictionary
Dim AnalistaNivel3 As New Dictionary
Dim AnalistaNivel4 As New Dictionary
Dim AnalistaNivel5 As New Dictionary

Set Planilha = PastadeTrabalhoRAV


PastadeTrabalhoRAV.Application.ScreenUpdating = False
PastadeTrabalhoRAV.Application.Calculation = xlCalculationManual
PastadeTrabalhoRAV.Application.EnableEvents = False
PastadeTrabalhoRAV.Application.DisplayAlerts = False
PastadeTrabalhoRAV.Application.EnableMacroAnimations = False




Set PlanilhaRav = PastadeTrabalhoRAV.Worksheets(1)
'Coloque aqui a consulta caso dê merda


PrimeiraLinha = PlanilhaRav.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
UltimaLinha = PlanilhaRav.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row

'Set Intervalo = PlanilhaRAV.Range(RAVColunas.DocumentoComplexidade & ":" & RAVColunas.DocumentoComplexidade).SpecialCells(xlCellTypeVisible).CurrentRegion
'ColunaDisponivel = GetPrimeiraColunaDisponvel(PastadeTrabalhoRAV, 1)
'AreasPlanRav = PlanilhaRAV.Range(RAVColunas.Linha & PrimeiraLinha & ":" & RAVColunas.Linha & UltimaLinha).SpecialCells(xlCellTypeVisible).Areas.Count
'If AreasPlanRav > 1 Then
'For aux1 = 1 To AreasPlanRav
'
'LinhaPlanRav = PlanilhaRAV.Range(RAVColunas.Linha & PrimeiraLinha & ":" & RAVColunas.Linha & UltimaLinha).SpecialCells(xlCellTypeVisible).Areas(aux1).Value2
'If IsArray(LinhaPlanRav) = True Then
'For Each LinhaAtualPlanRav In LinhaPlanRav
'If Not DocumentoLinhasPlanRav.Exists(LinhaAtualPlanRav) Then
'DocumentoLinhasPlanRav.Add LinhaAtualPlanRav, 0
'End If
''QtdLinhasPlanRav = QtdLinhasPlanRav + UBound(LinhaPlanRav)
'Next LinhaAtualPlanRav
'Else
'If Not DocumentoLinhasPlanRav.Exists(LinhaPlanRav) Then
'DocumentoLinhasPlanRav.Add LinhaPlanRav, 0
'End If
'End If
'Next aux1
'
'
'Else
'
'LinhaPlanRav = PlanilhaRAV.Range(RAVColunas.Linha & PrimeiraLinha & ":" & RAVColunas.Linha & UltimaLinha).SpecialCells(xlCellTypeVisible).Value2
'For Each LinhaAtualPlanRav In LinhaPlanRav
'If Not DocumentoLinhasPlanRav.Exists(LinhaAtualPlanRav) Then
'DocumentoLinhasPlanRav.Add LinhaAtualPlanRav, 0
'End If
'Next LinhaAtualPlanRav
'End If


Set PlanilhaEspelho = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, True, UltimaLinha, RAVColunas)
resposta = MsgBox("Como deseja distribuir? Deseja distribuir manual ou automático? A distribuição manual é ideal quando se deseja" _
& "distribuir documentos de forma isolada", vbYesNo)
If resposta = vbYes Then
FrmDistribuiçãoInteligente.Show
End If
AnalistaCargoComplexidade = GetColaboradorDadosResumidosAPI(RAVPreferencias(1).celula, 4)
Set Filtro = PlanilhaEspelho.Range("A1").CurrentRegion

Set Preferencias = New CRAVPreferencias
Preferencias.celula = CStr(RAVPreferencias(1).celula)
Preferencias.Analista = RAVPreferencias(1).Analista
Preferencias.Cliente = RAVPreferencias(1).Cliente
Preferencias.ambiente = RAVPreferencias(1).ambiente
Preferencias.Analistas = AnalistaCargoComplexidade


RAVPreferencias.Remove (1)
RAVPreferencias.Add Preferencias
FrmAnalistasDisponiveis.Show
If UBound(RAVPreferencias(1).Analistas, 2) > 0 Then
AnalistaCargoComplexidade = RAVPreferencias(1).Analistas
End If
QtdAnalista = UBound(AnalistaCargoComplexidade, 2)
PrimeiraLinha = 2
Linha = 0
aux = 1
UltimaLinha = Filtro.Rows.Count
Nivel = PlanilhaEspelho.Range(RAVColunas.DocumentoComplexidade & 2 & ":" & RAVColunas.DocumentoComplexidade & UltimaLinha).Value2
Analistas = PlanilhaEspelho.Range(RAVColunas.Analista & 2 & ":" & RAVColunas.Analista & UltimaLinha).Value2
Status = PlanilhaEspelho.Range(RAVColunas.Status & 2 & ":" & RAVColunas.Status & UltimaLinha).Value2



resposta = MsgBox("Deseja somente distribuir os documentos que fazem parte da meta diária? Isto é, documentos em prazo fatal" _
& " e documentos atrasados.", vbYesNo + vbInformation)

If resposta = vbYes Then
Modo = 0
For NivelAtual = 1 To UBound(Nivel)
Select Case Nivel(NivelAtual, 1)

Case 1
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
If Status(NivelAtual, 1) = "Atrasado" Or Status(NivelAtual, 1) = "Prazo Fatal" Then
DocumentoNivel1.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
End If
Case 2
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
If Status(NivelAtual, 1) = "Atrasado" Or Status(NivelAtual, 1) = "Prazo Fatal" Then
DocumentoNivel2.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
End If
Case 3
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
If Status(NivelAtual, 1) = "Atrasado" Or Status(NivelAtual, 1) = "Prazo Fatal" Then
DocumentoNivel3.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
End If
Case 4
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
If Status(NivelAtual, 1) = "Atrasado" Or Status(NivelAtual, 1) = "Prazo Fatal" Then
DocumentoNivel4.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
End If
Case 5
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
If Status(NivelAtual, 1) = "Atrasado" Or Status(NivelAtual, 1) = "Prazo Fatal" Then
DocumentoNivel5.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
End If
End Select
If Status(NivelAtual, 1) = "NE" Then
Linha = Linha + 1
GoTo proxnivel_
End If
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) = "N/A" _
Or Analistas(NivelAtual, 1) <> "-" Then
If Status(NivelAtual, 1) <> "Atrasado" Or Status(NivelAtual, 1) <> "Prazo Fatal" Then
Linha = Linha + 1
End If
End If
'aux = aux + 1
proxnivel_:
Next NivelAtual




Else
Modo = 1
For NivelAtual = 1 To UBound(Nivel)
Select Case Nivel(NivelAtual, 1)
Case 1
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
'DocumentoNivel1.Add NivelAtual, Nivel(NivelAtual, 1)
DocumentoNivel1.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
Case 2
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
DocumentoNivel2.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
Case 3
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
DocumentoNivel3.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
Case 4
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
DocumentoNivel4.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
Case 5
If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) <> "N/A" And Status(NivelAtual, 1) <> "NE" Then
DocumentoNivel5.Add DocumentoLinhasPlanRav.Keys(Linha), Nivel(NivelAtual, 1)
Linha = Linha + 1
End If
End Select

If Status(NivelAtual, 1) = "NE" Then
Linha = Linha + 1
GoTo proxnivel
End If

If Analistas(NivelAtual, 1) = "-" And Status(NivelAtual, 1) = "N/A" _
Or Analistas(NivelAtual, 1) <> "-" Then
Linha = Linha + 1
End If
'aux = aux + 1
proxnivel:
Next NivelAtual
End If


QtdDocumentosSemDivida = PlanilhaEspelho.Application.WorksheetFunction.CountIf _
(PlanilhaEspelho.Range(RAVColunas.Divida & PrimeiraLinha & ":" & RAVColunas.Divida & UltimaLinha), "Não")


For AnalistaAtual = 0 To QtdAnalista

If AnalistaCargoComplexidade(1, AnalistaAtual) = 2 Then
AnalistasNivel2 = AnalistasNivel2 + 1
AnalistaNivel2.Add AnalistaCargoComplexidade(0, AnalistaAtual), AnalistaCargoComplexidade(1, AnalistaAtual)
ElseIf AnalistaCargoComplexidade(1, AnalistaAtual) = 3 Then
AnalistasNivel3 = AnalistasNivel3 + 1
AnalistaNivel3.Add AnalistaCargoComplexidade(0, AnalistaAtual), AnalistaCargoComplexidade(1, AnalistaAtual)
ElseIf AnalistaCargoComplexidade(1, AnalistaAtual) = 4 Then
AnalistasNivel4 = AnalistasNivel4 + 1
AnalistaNivel4.Add AnalistaCargoComplexidade(0, AnalistaAtual), AnalistaCargoComplexidade(1, AnalistaAtual)
Else
AnalistasNivel5 = AnalistasNivel5 + 1
AnalistaNivel5.Add AnalistaCargoComplexidade(0, AnalistaAtual), AnalistaCargoComplexidade(1, AnalistaAtual)
End If

Next AnalistaAtual

ReDim QtdAnalistaPorNivel(3)
ReDim QtdDocumentoPorNivel(4)
QtdAnalistaPorNivel(0) = AnalistaNivel2.Count
QtdAnalistaPorNivel(1) = AnalistaNivel3.Count
QtdAnalistaPorNivel(2) = AnalistaNivel4.Count
QtdAnalistaPorNivel(3) = AnalistaNivel5.Count
QtdDocumentoPorNivel(0) = DocumentoNivel1.Count
QtdDocumentoPorNivel(1) = DocumentoNivel2.Count
QtdDocumentoPorNivel(2) = DocumentoNivel3.Count
QtdDocumentoPorNivel(3) = DocumentoNivel4.Count
QtdDocumentoPorNivel(4) = DocumentoNivel5.Count



If QtdDocumentosSemDivida > 0 Then
FiltroDistribuirDocumentos PastadeTrabalhoRAV, 2, Filtro, RAVColunas.Analista
Set PlanilhaEspelho = PastadeTrabalhoRAV.Worksheets(2)
Else
MsgBox "Não há documentos para distribuir.", vbInformation
Exit Function
End If

QtdDocumentoNivel1 = DocumentoNivel1.Count
QtdDocumentoNivel2 = DocumentoNivel2.Count
QtdDocumentoNivel3 = DocumentoNivel3.Count
QtdDocumentoNivel4 = DocumentoNivel4.Count
QtdDocumentoNivel5 = DocumentoNivel5.Count
QtdAnalista = QtdAnalista + 1


resposta = MsgBox("Como deseja que seja feita a distribuição ? " _
& "Deseja faze-la com base na complexidade de cada cargo ?", vbYesNo + vbInformation)
If resposta = vbYes Then
Modo = 0
Else
Modo = 1
End If
AutoDistrib AnalistaCargoComplexidade, QtdAnalistaPorNivel, QtdAnalistaPorNivel, DocumentoNivel1, _
DocumentoNivel2, DocumentoNivel3, DocumentoNivel4, DocumentoNivel5, Modo, PlanilhaRav, RAVColunas
GerarCopiaPlanilhaTemporia PastadeTrabalhoRAV, 1, False, UltimaLinha, RAVColunas
Exit Function
Catch:
MsgBox "Algo deu errado. " & Err.Description & " " & Err.Source, vbCritical
Set PastadeTrabalhoRAV = Nothing
Set RAVColunas = Nothing
Set RAVPreferencias = Nothing
Set SelecaoFeriados = Nothing
Set FiltroPadrao = Nothing
Set EmailsCriados = Nothing
Set NomesEmailsCriados = Nothing
Set EmailsParaVisualizacao = Nothing
Set DocumentoNovo = Nothing
Unload frmPgInicial
End Function

Public Function FiltroDistribuirDocumentos(PastadeTrabalhoRAV As Excel.Workbook, PlanilhaNumero As Integer, _
Filtro As Excel.Range, Coluna As String)
Dim PlanilhaRav As Excel.Worksheet
Set PlanilhaRav = PastadeTrabalhoRAV.Worksheets(PlanilhaNumero)
PlanilhaRav.Range(Filtro.Address).AutoFilter Field:=GetColunaIndiceNumerico(Coluna), Criteria1:="-"
End Function



Function func(ParamArray args() As Variant) As Double
    Dim I As Long
    Dim cell As Range

    For I = LBound(args) To UBound(args)
        If TypeName(args(I)) = "Range" Then
            For Each cell In args(I)
                func = func + cell.Value
            Next cell
        Else
            func = func + args(I)
        End If
    Next I
End Function
Public Function GetCargo(Cargo As Integer) As String
Select Case Cargo
Case 1
GetCargo = "Estagiário"
Case 2
GetCargo = "Auxiliar I"
Case 3
GetCargo = "Auxiliar II"
Case 4
GetCargo = "Auxiliar III"
Case 5
GetCargo = "Assistente I"
Case 6
GetCargo = "Assistente II"
Case 7
GetCargo = "Assistente III"
Case 8
GetCargo = "Analista I"
Case 9
GetCargo = "Analista II"
Case 10
GetCargo = "Analista III"
Case 11
GetCargo = "Especialista I"
Case 12
GetCargo "Coordenador"
Case Else
End Select
End Function
Public Function URLEncode(termo As String)
URLEncode = WorksheetFunction.EncodeURL(termo)
End Function


    
