Attribute VB_Name = "SA"
Option Explicit
' Sistema de Arquivos

Public Function GerarPlanilhaDiasTrabalhados(ByVal FeriadoTrabalhados As Dictionary) As Excel.Workbook
'Gera e abre a planilha de Feriados
Dim PastaDiasTrabalhados As Excel.Workbook
Dim PlanilhaDiasTrabalhados As Excel.Worksheet
Set PastaDiasTrabalhados = Workbooks.Add
Dim DiretorioFoiCriado As Boolean
Dim MDTExiste As Boolean
DiretorioFoiCriado = GerarDiretorio(Environ("LOCALAPPDATA") & "\DOC")
MDTExiste = ExistePlanilhaDiasTrabalhados
If DiretorioFoiCriado = False Or DiretorioFoiCriado = True And MDTExiste = False Then
PastaDiasTrabalhados.SaveAs Environ("LOCALAPPDATA") & "\DOC\MDT.xlsx"
Set PlanilhaDiasTrabalhados = PastaDiasTrabalhados.Worksheets(1)
PlanilhaDiasTrabalhados.Cells(1, 1).Value2 = "Descrição Feriado"
PlanilhaDiasTrabalhados.Cells(1, 2).Value2 = "Data Feriado"
PlanilhaDiasTrabalhados.Cells(1, 3).Value2 = "Dia Trabalhado"
Dim Feriados As Variant
Dim Linha As Long
Dim Status As Boolean
Linha = 2
Dim Hoje As Date

Hoje = Format(Now(), "dd/mm/YYYY")

For Each Feriados In FeriadoTrabalhados.Keys
If FeriadosHomonimos.Exists(Feriados) = Feriados Then
PlanilhaDiasTrabalhados.Cells(Linha, 1).Value2 = FeriadosHomonimos.Item(Feriados)
PlanilhaDiasTrabalhados.Cells(Linha, 2).Value2 = Feriados
Else
PlanilhaDiasTrabalhados.Cells(Linha, 1).Value2 = FeriadoTrabalhados.Item(Feriados)
PlanilhaDiasTrabalhados.Cells(Linha, 2).Value2 = Feriados
End If

PlanilhaDiasTrabalhados.Cells(Linha, 3).Value2 = "-"

Linha = Linha + 1
Next Feriados

PlanilhaDiasTrabalhados.Columns("A:A").AutoFit
PlanilhaDiasTrabalhados.Columns("B:B").AutoFit
PlanilhaDiasTrabalhados.Range("B:B").NumberFormat = "m/d/yyyy"
PlanilhaDiasTrabalhados.Columns("C:C").AutoFit
PastaDiasTrabalhados.Save
Set GerarPlanilhaDiasTrabalhados = PastaDiasTrabalhados
Else
Set GerarPlanilhaDiasTrabalhados = AbrirPlanilhaDiasTrabalhados
End If
End Function

Public Function ExistePlanilhaDiasTrabalhados() As Boolean
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
If FSO.FileExists(Environ("LOCALAPPDATA") & "\DOC\MDT.xlsx") = True Then
Set FSO = Nothing
ExistePlanilhaDiasTrabalhados = True
Else
Set FSO = Nothing
ExistePlanilhaDiasTrabalhados = False
End If
End Function
Public Function AbrirPlanilhaDiasTrabalhados() As Excel.Workbook
Dim PastaMDT As Excel.Workbook
Application.EnableEvents = False
Set PastaMDT = Workbooks.Open(Environ("LOCALAPPDATA") & "\DOC\MDT.xlsx")
Application.EnableEvents = True
Set AbrirPlanilhaDiasTrabalhados = PastaMDT
End Function


Public Function AtualizarPlanilhaDiasTrabalhados(ByVal SelecaoFeriado As Dictionary, PastaMDT As Excel.Workbook) As Excel.Workbook
Dim PlanilhaMDT As Excel.Worksheet
Set PlanilhaMDT = PastaMDT.Worksheets(1)
Dim FeriadoTrabalhado As Variant
Dim EnderecoFeriado As Excel.Range
For Each FeriadoTrabalhado In SelecaoFeriado.Keys
Set EnderecoFeriado = PlanilhaMDT.Range("B:B").Find(what:=FeriadoTrabalhado, after:=Range("B1"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows)
If Not EnderecoFeriado Is Nothing Then
PlanilhaMDT.Cells(EnderecoFeriado.Row, 3).Value2 = "Não"
Else
PlanilhaMDT.Cells(EnderecoFeriado.Row, 3).Value2 = "-"
End If
Next FeriadoTrabalhado
PastaMDT.Save
Set AtualizarPlanilhaDiasTrabalhados = PastaMDT
End Function

Public Function GerarDiretorio(Caminho As String) As Boolean
'Verifica a existência de um diretório. Retorna sempre uma string independente de estar criado ou não.
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
If Not FSO.FolderExists(Caminho) Then
FSO.CreateFolder Caminho
Set FSO = Nothing
GerarDiretorio = True
Else
Set FSO = Nothing
GerarDiretorio = False
End If
End Function
Public Function HorarioCriacaoArquivo(Caminho As String) As Date
On Error GoTo Catch
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim DataCriacaoArquivo As Date
DataCriacaoArquivo = FSO.GetFile(Caminho).DateCreated
Set FSO = Nothing
HorarioCriacaoArquivo = DataCriacaoArquivo
Exit Function
Catch:
MsgBox "Não foi possível recuperar o horário de criação da planilha", vbCritical
End Function
Public Function PastaAberta(Pasta As String) As Boolean
 Dim AQ As Long, ErrNo As Long
    On Error Resume Next
    AQ = FreeFile()
    Open Pasta For Input Lock Read As #AQ
    Close AQ
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    PastaAberta = False
    Case 70:   PastaAberta = True
    Case Else: Error ErrNo
    End Select
End Function


