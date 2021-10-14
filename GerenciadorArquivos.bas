Attribute VB_Name = "GerenciadorArquivos"
 Option Explicit
Option Compare Text
Public PastadeTrabalhoRAV As Excel.Workbook
Public RAVColunas As cRAVColunasXL
Public RAVPreferencias As New Collection
Public SelecaoFeriados As New Dictionary
Public FiltroPadrao As Excel.Range
Public EmailsCriados As New Dictionary
Public NomesEmailsCriados As Variant
Public EmailsParaVisualizacao As New Collection

Public Function GetArquivo() As String
Dim Dialogo As FileDialog
Set Dialogo = Application.FileDialog(msoFileDialogFilePicker)
Dim planilhacaminho As String
If Dialogo.Filters.Count > 1 Then
Dialogo.Filters.Clear
End If
Dialogo.InitialFileName = "C:\users\" & Environ("username") & "\Downloads"
Dialogo.Filters.Add "Arquivos Excel", "*.xlsx, *.xlsm, *.xlsb, *.xls"
Dialogo.FilterIndex = 1
Dialogo.Title = "Escolha um arquivo"
Dialogo.AllowMultiSelect = False
If Dialogo.Show Then
planilhacaminho = Dialogo.SelectedItems(1)
Else
MsgBox "Favor selecionar um arquivo excel.", vbExclamation: Exit Function
End If
GetArquivo = planilhacaminho
Set Dialogo = Nothing
End Function

Public Function FiltrarCelulas(celula As String, Clientes As Object, caminhoplanilha As String, CelulaID As Integer) As String
On Error GoTo Catch
Dim appXL As Excel.Application
Set appXL = CreateObject("Excel.Application")
Dim wbXl As Excel.Workbook
Dim shXL As Excel.Worksheet
Dim raXL As Excel.Range
Set wbXl = appXL.Workbooks.Open(Filename:=caminhoplanilha, ReadOnly:=False)
Dim SuplementoAtual As Excel.AddIn
appXL.Visible = True
Dim IndiceColunaDisponivel As Integer
Dim Linhas As Variant
Dim ColunaDisponivel As String
Dim Preferencias As CRAVPreferencias
For Each SuplementoAtual In appXL.AddIns
If SuplementoAtual.Installed Then
SuplementoAtual.Installed = False
SuplementoAtual.Installed = True
End If
Next SuplementoAtual
wbXl.Application.ScreenUpdating = False
wbXl.Application.Calculation = xlCalculationManual
wbXl.Application.EnableEvents = False
wbXl.Application.DisplayAlerts = False
wbXl.Application.AskToUpdateLinks = False
Set PastadeTrabalhoRAV = wbXl
Set shXL = wbXl.Worksheets(1)

'If shXL.AutoFilter.FilterMode = True Then
'shXL.ShowAllData
'End If

IndiceColunaDisponivel = 1 + shXL.Cells(1, 1).End(xlToRight).Column
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Dias no Sistema"
shXL.Range(ColunaDisponivel & ":" & ColunaDisponivel).NumberFormat = "General"

IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Dias Aguardando Análise"
shXL.Range(ColunaDisponivel & ":" & ColunaDisponivel).NumberFormat = "General"


IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Tipo de Documento"
shXL.Range(ColunaDisponivel & ":" & ColunaDisponivel).NumberFormat = "General"

IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Prazo Máximo Para Análise"
shXL.Range(ColunaDisponivel & ":" & ColunaDisponivel).NumberFormat = "General"

IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Status Documento"
shXL.Range(ColunaDisponivel & ":" & ColunaDisponivel).NumberFormat = "General"

IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Feriados Selecionados"

IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Complexidade do Documento"

IndiceColunaDisponivel = IndiceColunaDisponivel + 1
ColunaDisponivel = Split(Cells(1, IndiceColunaDisponivel).Address, "$")(1)

shXL.Columns(ColunaDisponivel).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
shXL.Range(ColunaDisponivel & 1).Value = "Linha"
shXL.Range(ColunaDisponivel & ":" & ColunaDisponivel).NumberFormat = "General"


Dim RAVColunas As cRAVColunasXL
Set RAVColunas = New cRAVColunasXL
Set RAVColunas = GetColunasIndiceAlfabetico(PastadeTrabalhoRAV)

Set raXL = shXL.Range("A1").CurrentRegion

If Clientes.Value = "Todos os clientes" Then
Dim QtdClientes As Integer: QtdClientes = Clientes.ListCount - 1
Dim ListaClientes As New Dictionary
Dim Cliente As Integer
For Cliente = 0 To QtdClientes
ListaClientes.Add Clientes.List(Cliente), celula
Next Cliente
shXL.Range(raXL.Address).AutoFilter Field:=2, Criteria1:=ListaClientes.Keys, Operator:=xlFilterValues
Else
shXL.Range(raXL.Address).AutoFilter Field:=2, Criteria1:=Clientes.List(0)
End If
Set PastadeTrabalhoRAV = wbXl

Set Preferencias = New CRAVPreferencias
Preferencias.celula = celula
Preferencias.Cliente = Clientes
Preferencias.CelulaID = CelulaID
RAVPreferencias.Add Preferencias
If Preferencias.Cliente = "Todos os clientes" Then
Linhas = GetCelulaPrazoAPI(Preferencias.celula, 1, "")
Else
Linhas = GetCelulaPrazoAPI(Preferencias.celula, 1, Preferencias.Cliente)
End If
CalculaPrazo Linhas, PastadeTrabalhoRAV, RAVColunas
Set PastadeTrabalhoRAV = Nothing
Set RAVColunas = Nothing
Set RAVPreferencias = Nothing
Set SelecaoFeriados = Nothing
Set FiltroPadrao = Nothing
Set EmailsCriados = Nothing
Set NomesEmailsCriados = Nothing
Set EmailsParaVisualizacao = Nothing
Set DocumentoNovo = Nothing
Exit Function
'frmControlesPrazo.Show
Catch:
MsgBox "Aconteceu um erro " & Err.Description & " " & Err.Source, vbCritical
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


Public Function TingePrazo(PlanilhaAtiva As Excel.Worksheet, Linha As Excel.Range, _
ByVal Tipo As Integer, ByVal DiasEmAnalise As Variant, ByVal Documento As String, Optional TipoTingemento As Integer)

If TipoTingemento = Empty Then

If Tipo = 0 And Documento <> "Acordo Coletivo de Trabalho" And Documento <> "Convenção Coletiva de Trabalho" Then
Select Case DiasEmAnalise
Case 0
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(76, 175, 80)
Case 1
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(217, 255, 0)
Case 2
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 255, 0)
Case 3
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 229, 0)
Case 4
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 120, 0)
Case 5
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 0, 0)
Linha.Font.Color = vbWhite
Case Else
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(51, 51, 51)
Linha.Font.Color = vbWhite
End Select

ElseIf Documento = "Acordo Coletivo de Trabalho" Or Documento = "Convenção Coletiva de Trabalho" Then
Select Case DiasEmAnalise
Case 0
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(76, 175, 80)
Case 1
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 220, 153)
Case 2
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 158, 0)
Case 3
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 0, 0)
Linha.Font.Color = vbWhite
Case Else
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(51, 51, 51)
Linha.Font.Color = vbWhite
End Select

Else
Select Case DiasEmAnalise
Case 0
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(76, 175, 80)
Case 1
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(255, 0, 0)
Linha.Font.Color = vbWhite
Case Else
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Interior.Color = RGB(51, 51, 51)
Linha.Font.Color = vbWhite
End Select
End If
Else
If Tipo = 0 And Documento <> "Acordo Coletivo de Trabalho" And Documento <> "Convenção Coletiva de Trabalho" Then
If DiasEmAnalise < 5 Then
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(50, 205, 50)
Linha.Font.Underline = True
ElseIf DiasEmAnalise = 5 Then
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(255, 0, 0)
Linha.Font.Underline = True
Else
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(105, 105, 105)
Linha.Font.Underline = True
End If

ElseIf Documento = "Acordo Coletivo de Trabalho" Or Documento = "Convenção Coletiva de Trabalho" Then
If DiasEmAnalise < 3 Then
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(76, 175, 80)
Linha.Font.Underline = True
Linha.Font.Bold = True
ElseIf DiasEmAnalise = 3 Then
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(255, 0, 0)
Linha.Font.Underline = True
Linha.Font.Bold = True
Else
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(105, 105, 105)
Linha.Font.Underline = True
Linha.Font.Bold = True
End If

Else
If DiasEmAnalise < 1 Then
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(76, 175, 80)
Linha.Font.Bold = True
ElseIf DiasEmAnalise = 1 Then
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(255, 0, 0)
Linha.Font.Bold = True
Else
Set Linha = PlanilhaAtiva.Range(Linha.Address)
Linha.Font.Color = RGB(97, 97, 97)
Linha.Font.Bold = True
End If
End If
End If
End Function
Public Function PlanilhaDinamica(PastadeTrabalhoRAV As Excel.Workbook)
Dim FonteDeDadosPlanilha As Excel.Worksheet
Set FonteDeDadosPlanilha = PastadeTrabalhoRAV.Worksheets(1)
Dim Repositorio As Excel.Worksheet
Set Repositorio = PastadeTrabalhoRAV.Worksheets(2)
Dim PlanilhaDinamicaCriada As Excel.Worksheet
Dim IntervaloDinamico As Excel.Range
Dim LinhaFinal, ColunaFinal As Long
Dim Pivot As PivotTable
Dim PivotCache As PivotCache
Dim celula As CCelula
Set celula = New CCelula
Dim Clientes, CL As Variant
Dim ClienteAtual, QtdClientes As Integer
PastadeTrabalhoRAV.Sheets.Add after:=Repositorio
Set PlanilhaDinamicaCriada = PastadeTrabalhoRAV.Worksheets(3)
PlanilhaDinamicaCriada.Name = "Planilha Dinâmica"

PlanilhaDinamicaCriada.Application.DisplayAlerts = False
PlanilhaDinamicaCriada.Application.ScreenUpdating = False
PlanilhaDinamicaCriada.Application.Calculation = xlCalculationManual
PastadeTrabalhoRAV.Application.Calculation = xlCalculationManual
PastadeTrabalhoRAV.Application.ScreenUpdating = False

LinhaFinal = FonteDeDadosPlanilha.Range("J:J").SpecialCells(xlCellTypeVisible).End(xlDown).Row
ColunaFinal = FonteDeDadosPlanilha.Cells(1, Columns.Count).End(xlToLeft).Column
Set IntervaloDinamico = FonteDeDadosPlanilha.Cells(1, 1).Resize(LinhaFinal, ColunaFinal)
Set PivotCache = PastadeTrabalhoRAV.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=IntervaloDinamico)
Set Pivot = PivotCache.CreatePivotTable(TableDestination:=PlanilhaDinamicaCriada.Cells(2, 2), TableName:="MacroCelula")

With PlanilhaDinamicaCriada.PivotTables("MacroCelula").PivotFields("Atribuído a:")
.Orientation = xlRowField
.Position = 1
End With

With PlanilhaDinamicaCriada.PivotTables("MacroCelula").PivotFields("Documento")
.Orientation = xlDataField
.Position = 1
End With

With PlanilhaDinamicaCriada.PivotTables("MacroCelula").PivotFields("Dias Aguardando Análise")
.Orientation = xlColumnField
.Position = 1
End With

With PlanilhaDinamicaCriada.PivotTables("MacroCelula").PivotFields("Documento")
.Orientation = xlPageField
.Position = 1
End With


celula.Nome = RAVPreferencias(1).celula
Clientes = GetClienteAPI(celula.Nome)

With PlanilhaDinamicaCriada.PivotTables("MacroCelula").PivotFields("Cliente")
.Orientation = xlPageField
.Position = 1
.EnableMultiplePageItems = True

Dim PDItm As PivotItem

For Each PDItm In PlanilhaDinamicaCriada.PivotTables("MacroCelula").PivotFields("Cliente").PivotItems
   If Not IsError(Application.Match(PDItm.Caption, Clientes, 0)) Then
       PDItm.Visible = True
   Else
        PDItm.Visible = False
    End If
    Next PDItm


End With

PastadeTrabalhoRAV.SlicerCaches.Add2(PlanilhaDinamicaCriada.PivotTables("MacroCelula"), _
        "Status Documento").Slicers.Add PlanilhaDinamicaCriada, , "Status Documento", _
        "Status Documento", 150, 380.75, 150, 450.75

PastadeTrabalhoRAV.SlicerCaches.Add2(PlanilhaDinamicaCriada.PivotTables("MacroCelula"), _
        "Tipo de Documento").Slicers.Add PlanilhaDinamicaCriada, , "Tipo de Documento", _
        "Tipo de Documento", 155, 385.75, 155, 455.75
        
        
PastadeTrabalhoRAV.SlicerCaches.Add2(PlanilhaDinamicaCriada.PivotTables("MacroCelula"), _
        "Atribuído a:").Slicers.Add PlanilhaDinamicaCriada, , "Atribuído a:", _
        "Atribuído a:", 170, 405.75, 175, 475.75

End Function


Public Function CalculaPrazo(ByVal Documentos As Variant, PastadeTrabalhoRAV As Excel.Workbook, RAVColunas As cRAVColunasXL)
Dim LinhaInicialFeriado, LinhaFinalFeriado As Long
On Error GoTo Catch:
'FrmCarregandoGrande.Show vbModeless
PastadeTrabalhoRAV.Application.ScreenUpdating = False
PastadeTrabalhoRAV.Application.Calculation = xlCalculationManual
PastadeTrabalhoRAV.Application.EnableEvents = False
PastadeTrabalhoRAV.Application.DisplayAlerts = False
PastadeTrabalhoRAV.Application.EnableMacroAnimations = False
Dim PastaMDT As Excel.Workbook
Dim PlanilhaRav, PlanilhaMDT, PlanilhaTemp As Excel.Worksheet
Dim PlanilhaIntervalo As Excel.Range
Set PlanilhaRav = PastadeTrabalhoRAV.Worksheets(1)
Dim LinhaFinal As Excel.Range
Set LinhaFinal = PlanilhaRav.UsedRange.SpecialCells(xlCellTypeVisible).CurrentRegion
Set PlanilhaIntervalo = PlanilhaRav.UsedRange.Offset(1, 0).Resize(LinhaFinal.Rows.Count - 1, LinhaFinal.Columns.Count).SpecialCells(xlCellTypeVisible).Rows
Set FiltroPadrao = PlanilhaIntervalo
Dim Linha As Range
Dim UltimaColuna, ColunaCliente, ColunaFornecedor, ColunaDocumento, ColunaDataInclusao, _
ColunaDivida, ColunaInadimpliencia, ColunaFimInadimplencia, ColunaDiasNoSistema, ColunaDiasAguardandoAnalise, _
ColunaTipo, ColunaPrazoMaximo, ColunaStatus, ColunaFeriados As String
Dim FiltroExistente As Boolean
Dim SemFeriados As Boolean: SemFeriados = False
Dim Feriados As New Dictionary
Dim FeriadoAux As New Dictionary
Dim DocumentosForadaBase As New Dictionary
Dim DocumentosLinhas As New Dictionary
Dim LinhaMDT, LinhaMDTAtual, PrimeiraLinha, UltimaLinha, UltimaLinhaRepositorio As Long
Dim MDT, DocumentosDisponiveis, Status, Linhas, DocumentoDisponivel As Variant
Dim DataDepositoMaisAntigo As Date
DataDepositoMaisAntigo = Now()


UltimaLinha = PlanilhaRav.Range("J:J").SpecialCells(xlCellTypeVisible).End(xlDown).Row

Set PlanilhaTemp = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, True, UltimaLinha, RAVColunas)

PrimeiraLinha = PlanilhaTemp.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
UltimaLinha = PlanilhaTemp.Range("J:J").SpecialCells(xlCellTypeVisible).End(xlDown).Row

DataDepositoMaisAntigo = PlanilhaTemp.Application.WorksheetFunction.MinIfs _
(PlanilhaTemp.Range(RAVColunas.DataInclusao & 2 & ":" & RAVColunas.DataInclusao & UltimaLinha), _
PlanilhaTemp.Range(RAVColunas.Divida & 2 & ":" & RAVColunas.Divida & UltimaLinha), "Não")

Set PlanilhaTemp = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, False, UltimaLinha, RAVColunas)

If ExistePlanilhaDiasTrabalhados = True Then

Dim Dia As Date: Dia = Format(Now(), "dd/mm/YYYY")
Dim CDataDepositoMaisAntigo As Date
CDataDepositoMaisAntigo = Format(DataDepositoMaisAntigo, "dd/mm/YYYY")
Set PastaMDT = AbrirPlanilhaDiasTrabalhados
Set PlanilhaMDT = PastaMDT.Worksheets(1)

LinhaMDT = PlanilhaMDT.Range("A:A").End(xlDown).Row
MDT = PlanilhaMDT.Range("A2:C" & Cells(LinhaMDT, 1).End(xlEnd).Row).Value2

LinhaMDT = LinhaMDT - 1

For LinhaMDTAtual = 1 To LinhaMDT
If DateValue(Dia) >= MDT(LinhaMDTAtual, 2) And MDT(LinhaMDTAtual, 3) = "-" Then
If Not SelecaoFeriados.Exists(MDT(LinhaMDTAtual, 2)) Then
SelecaoFeriados.Add MDT(LinhaMDTAtual, 2), MDT(LinhaMDTAtual, 1)
End If
Else
If CDate(MDT(LinhaMDTAtual, 2)) >= DateValue(CDataDepositoMaisAntigo) And MDT(LinhaMDTAtual, 3) = "Não" Then
FeriadoAux.Add MDT(LinhaMDTAtual, 2), MDT(LinhaMDTAtual, 1)
End If
End If
Next LinhaMDTAtual

'SelecaoFeriados.RemoveAll
'FeriadoAux.RemoveAll


If SelecaoFeriados.Count > 0 Then
frmEscolhaFeriados.Show vbModal
Unload frmEscolhaFeriados
If SelecaoFeriados.Count > 0 Then
AtualizarPlanilhaDiasTrabalhados SelecaoFeriados, PastaMDT
PastaMDT.Close (True)
End If
ElseIf FeriadoAux.Count > 0 Then
SemFeriados = False
Set SelecaoFeriados = FeriadoAux
PastaMDT.Close (True)
Else
SemFeriados = True
PastaMDT.Close (True)
End If

Else

Dim DividasMDT, DataDepositoMDT As Excel.Range
Dim LinhaInicialMDT, LinhaFinalMDT As Long

Set DataDepositoMDT = _
PlanilhaRav.Range(RAVColunas.DataInclusao & ":" & RAVColunas.DataInclusao).Rows.SpecialCells(xlCellTypeVisible)
Set DividasMDT = _
PlanilhaRav.Range(RAVColunas.Divida & ":" & RAVColunas.Divida).Rows.SpecialCells(xlCellTypeVisible)
LinhaInicialMDT = PlanilhaRav.Range("A:A").SpecialCells(xlCellTypeVisible).Rows(1).Row
LinhaFinalMDT = PlanilhaRav.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row


Set SelecaoFeriados = APIFeriados(DataDepositoMaisAntigo)
Set PastaMDT = GerarPlanilhaDiasTrabalhados(SelecaoFeriados)
Set PlanilhaMDT = PastaMDT.Worksheets(1)
SelecaoFeriados.RemoveAll

LinhaMDT = PlanilhaMDT.Range("A:A").End(xlDown).Row
MDT = PlanilhaMDT.Range("A2:C" & Cells(LinhaMDT, 1).End(xlEnd).Row).Value2
PastaMDT.Close (True)
LinhaMDT = LinhaMDT - 1
For LinhaMDTAtual = 1 To LinhaMDT
Dia = Format(Now(), "dd/mm/yyyy")

If DateValue(Dia) >= MDT(LinhaMDTAtual, 2) And MDT(LinhaMDTAtual, 3) = "-" Then
If Not SelecaoFeriados.Exists(MDT(LinhaMDTAtual, 2)) Then
SelecaoFeriados.Add MDT(LinhaMDTAtual, 2), MDT(LinhaMDTAtual, 1)
End If
Else
FeriadoAux.Add MDT(LinhaMDTAtual, 2), MDT(LinhaMDTAtual, 1)
End If
Next LinhaMDTAtual

If SelecaoFeriados.Count > 0 Then
frmEscolhaFeriados.Show vbModal
Unload frmEscolhaFeriados
If SelecaoFeriados.Count > 0 Then
SemFeriados = False
AtualizarPlanilhaDiasTrabalhados SelecaoFeriados, PastaMDT
End If
ElseIf FeriadoAux.Count > 0 Then
Set SelecaoFeriados = FeriadoAux
Else
SemFeriados = True
End If

End If

'frmEscolhaFeriados.Show vbModal
'Set Feriados = SelecaoFeriados

Dim NumLinha, ElementoAtual, QtdElementos, QtdDiasEmAnalise, aux, VerificadorExistenciaDocumento
Dim QtdDiaUtilPosFeriado, Qtdferiado, QtdFiltros, ColunaAtual, LinhaAtual As Integer
Dim FeriadoAtual, Colunas, Coluna, Empresa, Documento, PrazoMaximo, Complexidade, Tipo As Variant
Dim FeriadosSelecionados As Excel.Range




If SemFeriados = True Then
Qtdferiado = Feriados.Count - 1
aux = PlanilhaRav.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
For FeriadoAtual = 0 To Qtdferiado
PlanilhaRav.Range(RAVColunas.Feriados & aux).Value2 = CDate(Feriados.Keys(FeriadoAtual))
aux = aux + 1
Next FeriadoAtual
End If

PastadeTrabalhoRAV.Sheets.Add after:=PlanilhaRav
Dim PlanilhaRepositorio As Excel.Worksheet
Set PlanilhaRepositorio = PastadeTrabalhoRAV.Worksheets(2)

PlanilhaRepositorio.Name = "Repositório"
PlanilhaRepositorio.Range("A1").Value2 = "Empresa"
PlanilhaRepositorio.Range("B1").Value2 = "Documento"
PlanilhaRepositorio.Range("C1").Value2 = "Prazo"
PlanilhaRepositorio.Range("D1").Value2 = "Tipo"
PlanilhaRepositorio.Range("E1").Value2 = "Complexidade"
QtdElementos = UBound(Documentos, 2)
ReDim Empresa(0 To QtdElementos)
ReDim Documento(0 To QtdElementos)
ReDim PrazoMaximo(0 To QtdElementos)
ReDim Complexidade(0 To QtdElementos)
ReDim Tipo(0 To QtdElementos)

aux = 2


For ElementoAtual = 0 To QtdElementos
Empresa(ElementoAtual) = Documentos(0, ElementoAtual)
Documento(ElementoAtual) = Documentos(1, ElementoAtual)
PrazoMaximo(ElementoAtual) = Documentos(2, ElementoAtual)
If Documentos(3, ElementoAtual) = 0 Then
Tipo(ElementoAtual) = "COMUM"
Else
Tipo(ElementoAtual) = "BLOQUEIO"
End If
Complexidade(ElementoAtual) = Documentos(4, ElementoAtual)
Next ElementoAtual

PlanilhaRepositorio.Range("A2").Resize(QtdElementos + 1) = Application.Transpose(Empresa)
PlanilhaRepositorio.Range("B2").Resize(QtdElementos + 1) = Application.Transpose(Documento)
PlanilhaRepositorio.Range("C2").Resize(QtdElementos + 1) = Application.Transpose(PrazoMaximo)
PlanilhaRepositorio.Range("D2").Resize(QtdElementos + 1) = Application.Transpose(Tipo)
PlanilhaRepositorio.Range("E2").Resize(QtdElementos + 1) = Application.Transpose(Complexidade)

PlanilhaRepositorio.Columns("A:E").AutoFit


UltimaLinhaRepositorio = PlanilhaRepositorio.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row
PrimeiraLinha = PlanilhaRav.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
UltimaLinha = PlanilhaRav.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row

If SemFeriados = False Then
PlanilhaRav.Range(RAVColunas.Feriados & ":" & RAVColunas.Feriados).NumberFormat = "m/d/yyyy"
aux = PrimeiraLinha + Qtdferiado
End If

PlanilhaRav.Range(RAVColunas.PrazoMaximoAnalise & PrimeiraLinha).FormulaArray = _
"=IFERROR(INDEX(Repositório!$C$2:$C$" & UltimaLinhaRepositorio & ",MATCH(Plan1!" & RAVColunas.Cliente & PrimeiraLinha & "" _
& ":" & RAVColunas.Cliente & UltimaLinha & "&Plan1!" & RAVColunas.Documento & PrimeiraLinha & ":" & RAVColunas.Documento & UltimaLinha & "," _
& "Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" & UltimaLinhaRepositorio & ",0)),""NE"")"

PlanilhaRav.Range(RAVColunas.DiasNoSistema & PrimeiraLinha).Formula = "=IF(" & RAVColunas.Tipo & PrimeiraLinha & "<>""NE""," _
& "DAYS(TODAY()," & RAVColunas.DataInclusao & PrimeiraLinha & "),""NE"")"

If SemFeriados = True Then

PlanilhaRav.Range(RAVColunas.DiasAguardandoAnalise & PrimeiraLinha).Formula = _
"=IF(" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & "=""NE"",""NE"",IF(AND(" & RAVColunas.Inadimpliencia & PrimeiraLinha & "<>""-""," _
& RAVColunas.FimInadimplencia & PrimeiraLinha & "=""-""),""-"",IF(" & RAVColunas.FimInadimplencia & PrimeiraLinha & "=""-""," _
& "CPDOD(" & RAVColunas.DataInclusao & PrimeiraLinha & "),IF(" & RAVColunas.FimInadimplencia & PrimeiraLinha & ">" _
& RAVColunas.DataInclusao & PrimeiraLinha & ",CPDOD(" & RAVColunas.FimInadimplencia & PrimeiraLinha & ")," _
& "CPDOD(" & RAVColunas.DataInclusao & PrimeiraLinha & ")))))"

Else
PlanilhaRav.Range(RAVColunas.DiasAguardandoAnalise & PrimeiraLinha).Formula = _
"=IF(" & RAVColunas.PrazoMaximoAnalise & PrimeiraLinha & "=""NE"",""NE"",IF(AND(" & RAVColunas.Inadimpliencia & PrimeiraLinha & "<>""-""," _
& RAVColunas.FimInadimplencia & PrimeiraLinha & "=""-""),""-"",IF(" & RAVColunas.FimInadimplencia & PrimeiraLinha & "=""-""," _
& "CPDOD(" & RAVColunas.DataInclusao & PrimeiraLinha & ",$" & RAVColunas.Feriados & "$" & PrimeiraLinha & ":$" _
& RAVColunas.Feriados & "$" & aux & "),IF(" & RAVColunas.FimInadimplencia & PrimeiraLinha & ">" _
& RAVColunas.DataInclusao & PrimeiraLinha & ",CPDOD(" & RAVColunas.FimInadimplencia & PrimeiraLinha & "," _
& "$" & RAVColunas.Feriados & "$" & PrimeiraLinha & ":$" & RAVColunas.Feriados & "$" & aux & ")," _
& "CPDOD(" & RAVColunas.DataInclusao & PrimeiraLinha & ",$" & RAVColunas.Feriados & "$" & PrimeiraLinha & "" _
& ":$" & RAVColunas.Feriados & "$" & aux & ")))))"
End If

PlanilhaRav.Range(RAVColunas.Tipo & PrimeiraLinha).FormulaArray = _
"=IFERROR(INDEX(Repositório!$D$2:$D$" & UltimaLinhaRepositorio & ",MATCH(Plan1!" & RAVColunas.Cliente & PrimeiraLinha & "" _
& ":B" & UltimaLinha & "&Plan1!" & RAVColunas.Documento & PrimeiraLinha & ":" & RAVColunas.Feriados & UltimaLinha & "," _
& "Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" & UltimaLinhaRepositorio & ",0)),""NE"")"


PlanilhaRav.Range(RAVColunas.Status & PrimeiraLinha).Formula = _
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


PlanilhaRav.Range(RAVColunas.DocumentoComplexidade & PrimeiraLinha).FormulaArray = _
"=IFERROR(INDEX(Repositório!$E$2:$E$" & UltimaLinhaRepositorio & ",MATCH(Plan1!" & RAVColunas.Cliente & PrimeiraLinha & "" _
& ":" & RAVColunas.Cliente & UltimaLinha & "&Plan1!" & RAVColunas.Documento & PrimeiraLinha & ":" & RAVColunas.Documento & UltimaLinha & "," _
& "Repositório!$A$2:$A$" & UltimaLinhaRepositorio & "&Repositório!$B$2:$B$" & UltimaLinhaRepositorio & ",0)),""NE"")"

PlanilhaRav.Range(RAVColunas.Linha & PrimeiraLinha).Formula = "=ROW()"

PlanilhaRav.Range(RAVColunas.DiasNoSistema & PrimeiraLinha & ":" & RAVColunas.Linha & UltimaLinha).FillDown
PlanilhaRav.Range(RAVColunas.DocumentoComplexidade & PrimeiraLinha & ":" & RAVColunas.DocumentoComplexidade & UltimaLinha).FillDown
PlanilhaRav.Columns(RAVColunas.DiasNoSistema & ":" & RAVColunas.Linha).AutoFit

PastadeTrabalhoRAV.Application.Calculate

Do While PastadeTrabalhoRAV.Application.CalculationState <> xlDone
     DoEvents
Loop
Set PlanilhaTemp = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, True, UltimaLinha, RAVColunas)
PlanilhaTemp.Application.DisplayAlerts = False
UltimaLinha = PlanilhaTemp.Range("A:A").End(xlDown).Row

DocumentosDisponiveis = PlanilhaTemp.Range(RAVColunas.Documento & 2 & ":" & RAVColunas.Documento & UltimaLinha).Value2
Status = PlanilhaTemp.Range(RAVColunas.Status & 2 & ":" & RAVColunas.Status & UltimaLinha).Value2
Linhas = PlanilhaTemp.Range(RAVColunas.Linha & 2 & ":" & RAVColunas.Linha & UltimaLinha).Value2
Empresa = PlanilhaTemp.Range(RAVColunas.Cliente & 2 & ":" & RAVColunas.Cliente & UltimaLinha).Value2
PlanilhaTemp.Delete


aux = 1
DocumentosForadaBase.RemoveAll
DocumentosLinhas.RemoveAll
For Each DocumentoDisponivel In DocumentosDisponiveis
If Status(aux, 1) = "NE" And Not DocumentosForadaBase.Exists(DocumentoDisponivel) Then
DocumentosForadaBase.Add DocumentoDisponivel, Empresa(aux, 1)
End If
'Else
'If Not DocumentosLinhas.Exists(Linhas(aux, 1)) And Status(aux, 1) <> "" Then
DocumentosLinhas.Add Linhas(aux, 1), 0
'End If

aux = aux + 1
Next DocumentoDisponivel

'FiltroExistente = FiltrarUmaColuna(PastadeTrabalhoRAV, Range(RAVColunas.Status & ":" & RAVColunas.Status), "NE", _
'GetColunaIndiceNumerico(RAVColunas.Status))
If DocumentosForadaBase.Count > 0 Then
Dim FiltroAtual As Excel.Range
Set FiltroAtual = PlanilhaRav.UsedRange.SpecialCells(xlCellTypeVisible).CurrentRegion
EditDocumentoAPI DocumentosForadaBase, RAVPreferencias(1).celula, RAVColunas
DistribuirDocumentos PastadeTrabalhoRAV, FiltroAtual, DocumentosLinhas, RAVColunas
PlanilhaDinamica PastadeTrabalhoRAV
PlanilhaRav.Calculate
DadosEmailDocumentosAusentes PastadeTrabalhoRAV, RAVColunas
'GetDocumentosForaDaBase PastadeTrabalhoRAV, FiltroAtual, GetColunaIndiceNumerico(RAVColunas.Documento), RAVColunas
Else
If PlanilhaRav.AutoFilter.FilterMode Then
PlanilhaRav.ShowAllData
End If
Dim raXL As Excel.Range
Set raXL = PlanilhaRav.Range("A1").CurrentRegion

If RAVPreferencias(1).Cliente = "Todos os clientes" Then
Dim qtdlinha As Long
Dim Nome As String
Dim Cliente, Clientes As Variant
Nome = RAVPreferencias(1).celula
'MsgBox RAVPreferencias(1).Celula
Clientes = GetClienteAPI(Nome)
Dim ListaClientes As New Dictionary


For Each Cliente In Clientes
ListaClientes.Add Cliente, RAVPreferencias(1).celula
Next Cliente
PlanilhaRav.Range(raXL.Address).AutoFilter Field:=2, Criteria1:=ListaClientes.Keys, Operator:=xlFilterValues
Else
PlanilhaRav.Range(raXL.Address).AutoFilter Field:=2, Criteria1:=Clientes.List(0)
End If
'PlanilhaAtiva.Range ("A1")

qtdlinha = PlanilhaRav.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row

Set FiltroAtual = PlanilhaRav.Range("A1").SpecialCells(xlCellTypeVisible).CurrentRegion

ColunaStatus = GetColunaIndiceNumerico(RAVColunas.Status)

PlanilhaRav.Range(FiltroAtual.Address).AutoFilter Field:=ColunaStatus



Set FiltroAtual = PlanilhaRav.UsedRange.SpecialCells(xlCellTypeVisible).CurrentRegion
DistribuirDocumentos PastadeTrabalhoRAV, FiltroAtual, DocumentosLinhas, RAVColunas

PlanilhaDinamica PastadeTrabalhoRAV
QtdFiltros = FiltrarVariasColunas(PastadeTrabalhoRAV, Range(RAVColunas.Status & ":" & RAVColunas.Status), 1)
'Set RAVColunas = Nothing
If QtdFiltros = 1 Then
DadosEmailDocumentosAusentes PastadeTrabalhoRAV, RAVColunas
Else
Set RAVColunas = Nothing
'Pensar no que pode acontencer aqui
MsgBox "Algo deu errado.", vbCritical
End If
End If
Exit Function
Catch:
MsgBox "Algo deu errado. " & Err.Description & " " & Err.Source & " " & Erl, vbCritical
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

Function GetDocumentosForaDaBase(PastadeTrabalhoRAV As Excel.Workbook, FiltroAnterior As Excel.Range, NumColuna As Integer, _
ByVal DocumentosAusentes As Dictionary, ByVal RAVColunas As cRAVColunasXL) As Integer
'Prepara documentos fora do banco de dados para inserção rápida
DoEvents
Dim PlanilhaRav As Excel.Worksheet
Dim PlanilhaEspelho As Excel.Worksheet
Set PlanilhaRav = PastadeTrabalhoRAV.Worksheets(1)
Dim QtdFiltros, ColunaStatus As Integer
Dim NumLinha, LinhaInicial, UltimaLinha, aux As Long
Dim DocumentosAusentes As New Dictionary
Dim Documentos, Documento, Clientes As Variant

'LinhaInicial = PlanilhaRAV.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
'UltimaLinha = PlanilhaRAV.Range("A:A").SpecialCells(xlCellTypeVisible).End(xlDown).Row
'Set PlanilhaEspelho = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, True, UltimaLinha, RAVColunas)
'UltimaLinha = PlanilhaEspelho.Range("A:A").End(xlDown).Row
'Documentos = PlanilhaEspelho.Range(RAVColunas.Documento & 2 & ":" & RAVColunas.Documento & UltimaLinha).Value2
'Clientes = PlanilhaEspelho.Range(RAVColunas.Cliente & 2 & ":" & RAVColunas.Cliente & UltimaLinha).Value2
'Set PlanilhaEspelho = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, False, UltimaLinha, RAVColunas)
'aux = 1
'For Each Documento In Documentos
'If Not DocumentosAusentes.Exists(Documento) Then
'DocumentosAusentes.Add Documento, Clientes(aux, 1)
'End If
'aux = aux + 1
'Next Documento


If DocumentosAusentes.Count > 0 Then
 EditDocumentoAPI DocumentosAusentes, RAVPreferencias(1).celula, RAVColunas


ColunaStatus = GetColunaIndiceNumerico(RAVColunas.Status)
PlanilhaRav.Application.ScreenUpdating = False
    PlanilhaRav.Range(FiltroAnterior.Address).AutoFilter Field:=ColunaStatus
    PlanilhaRav.Range(FiltroAnterior.Address).AutoFilter Field:=GetColunaIndiceNumerico(RAVColunas.DiasAguardandoAnalise), Criteria1:="<>NE"

Set FiltroAnterior = PlanilhaRav.UsedRange.SpecialCells(xlCellTypeVisible).CurrentRegion
    DistribuirDocumentos PastadeTrabalhoRAV, FiltroAnterior, RAVColunas
PlanilhaDinamica PastadeTrabalhoRAV
PlanilhaRav.Calculate

QtdFiltros = FiltrarVariasColunas(PastadeTrabalhoRAV, Range(RAVColunas.Status & ":" & RAVColunas.Status), 1)
If QtdFiltros = 1 Then
DadosEmailDocumentosAusentes PastadeTrabalhoRAV, RAVColunas
Else
'Pensar no que pode acontencer aqui
MsgBox "Algo deu errado.", vbCritical
End If
End If
End Function


Public Function DadosEmailDocumentosAusentes(PastadeTrabalhoRAV As Excel.Workbook, RAVColunas As cRAVColunasXL)

On Error GoTo Catch
PastadeTrabalhoRAV.Application.ScreenUpdating = False
PastadeTrabalhoRAV.Application.Calculation = xlCalculationManual
PastadeTrabalhoRAV.Application.EnableEvents = False
PastadeTrabalhoRAV.Application.DisplayAlerts = False

Dim PastadeTrabalhoCriada As Excel.Workbook
Dim PlanilhaAtiva, PlanilhaEspelho, PlanilhaCriada As Excel.Worksheet
Dim PlanilhaCriadaLinhas As Excel.Range
Set PlanilhaAtiva = PastadeTrabalhoRAV.Worksheets(1)
Dim IntervaloDocumentos, Linha, LinhaReferencia As Excel.Range
Set LinhaReferencia = PlanilhaAtiva.UsedRange.SpecialCells(xlCellTypeVisible).CurrentRegion
Set IntervaloDocumentos = PlanilhaAtiva.UsedRange.Offset(1, 0). _
Resize(LinhaReferencia.Rows.Count - 1, LinhaReferencia.Columns.Count).SpecialCells(xlCellTypeVisible).Rows
Dim preto As Long
preto = RGB(51, 51, 51)
Dim DepositoDadosBasicos, Empregado, AnalistaeData, Dia, PrazoMaximoAnalise, Status, Tipo, AnalistaAtualEnvioEmail, _
Lideres, Consulta, ArquivoCriado, Item As Variant

Dim resposta, AnalistaAtual, ItemAtual, QtdResultadosConsulta, EmailAtual, lider, QtdLideres, _
QtdPrazoFatal, QtdPrazoPerdido, QtdEmDia, QtdDocumentosSemAnalista, SomaQtdPrazoFatal, SomaQtdPrazoPerdido, _
SomaQtdEmDia, LinhaAtual As Integer

Dim LinhaInicial, NumLinha As Long
Dim Email, TabelaHTML, EmailPara, EmailCopia, EmailAssunto, EmailMensagem, _
EmailLiderNome, EmailLiderEmail, MeioTabela, AnalistaNome As String

Dim Criado As Boolean

'Consulta = GetColaboradorInfo(1, , RAVPreferencias(1).celula)
Consulta = GetColaboradorDadosResumidosAPI(RAVPreferencias(1).celula, 1)
NomesEmailsCriados = Consulta
QtdResultadosConsulta = UBound(Consulta, 2)
Dim AnalistasEncontrados As New Dictionary
Dim AnalistasHomonimos As New Dictionary
Dim PlanilhaAnalista As New Dictionary
AnalistasEncontrados.CompareMode = TextCompare
AnalistasHomonimos.CompareMode = TextCompare

Lideres = GetColaboradorDadosResumidosAPI(RAVPreferencias(1).celula, 5)
QtdLideres = UBound(Lideres, 2)
For AnalistaAtual = 0 To QtdResultadosConsulta
If Not AnalistasEncontrados.Exists(Consulta(0, AnalistaAtual)) Then
AnalistasEncontrados.Add Consulta(0, AnalistaAtual), Consulta(1, AnalistaAtual)
Else
AnalistasHomonimos.Add Consulta(0, AnalistaAtual), Consulta(1, AnalistaAtual)
End If
Next AnalistaAtual


'Dim RAVDeposito As CRAVDeposito
'Set RAVDeposito = New CRAVDeposito


Dim DocDepositadoAtribuido As New Collection
Dim DocDepositadoSemAnalista As New Collection
Dim DocumentosparaEnvioAnalista As New Collection

NumLinha = PlanilhaAtiva.UsedRange.SpecialCells(xlCellTypeVisible).End(xlDown).Row
Set PlanilhaEspelho = GerarCopiaPlanilhaTemporia(PastadeTrabalhoRAV, 1, True, NumLinha, RAVColunas)
NumLinha = PlanilhaEspelho.UsedRange.SpecialCells(xlCellTypeVisible).End(xlDown).Row
LinhaInicial = PlanilhaEspelho.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row

DepositoDadosBasicos = PlanilhaEspelho.Range(RAVColunas.Protocolo & LinhaInicial & _
":" & RAVColunas.Documento & Cells(NumLinha, 1).End(xlEnd).Row).Value2

Empregado = PlanilhaEspelho.Range(RAVColunas.Empregado & LinhaInicial & ":" _
& RAVColunas.Empregado & Cells(NumLinha, 1).End(xlEnd).Row).Value2

AnalistaeData = PlanilhaEspelho.Range(RAVColunas.Analista & LinhaInicial & ":" _
& RAVColunas.DataInclusao & Cells(NumLinha, 1).End(xlEnd).Row).Value2

Dia = PlanilhaEspelho.Range(RAVColunas.DiasAguardandoAnalise & LinhaInicial & ":" _
& RAVColunas.DiasAguardandoAnalise & Cells(NumLinha, 1).End(xlEnd).Row).Value2

PrazoMaximoAnalise = PlanilhaEspelho.Range(RAVColunas.PrazoMaximoAnalise & LinhaInicial & ":" _
& RAVColunas.PrazoMaximoAnalise & Cells(NumLinha, 1).End(xlEnd).Row).Value2

Status = PlanilhaEspelho.Range(RAVColunas.Status & LinhaInicial & ":" _
& RAVColunas.Status & Cells(NumLinha, 1).End(xlEnd).Row).Value2

Tipo = PlanilhaEspelho.Range(RAVColunas.Tipo & LinhaInicial & ":" & RAVColunas.Tipo & Cells(NumLinha, 4).End(xlEnd).Row).Value2

NumLinha = NumLinha - 1

For ItemAtual = 1 To NumLinha

If AnalistaeData(ItemAtual, 1) <> "-" And Status(ItemAtual, 1) <> "NE" Then
RAVColunas.Protocolo = DepositoDadosBasicos(ItemAtual, 1)
RAVColunas.Cliente = DepositoDadosBasicos(ItemAtual, 2)
RAVColunas.Fornecedor = DepositoDadosBasicos(ItemAtual, 3)
RAVColunas.Unidade = DepositoDadosBasicos(ItemAtual, 4)
RAVColunas.Documento = DepositoDadosBasicos(ItemAtual, 5)
RAVColunas.Empregado = Empregado(ItemAtual, 1)

If AnalistasEncontrados.Exists(AnalistaeData(ItemAtual, 1)) Then
Email = AnalistaeData(ItemAtual, 1)
RAVColunas.Analista = AnalistaeData(ItemAtual, 1)
RAVColunas.AnalistaEmail = AnalistasEncontrados.Item(AnalistaeData(ItemAtual, 1))
Else
RAVColunas.AnalistaEmail = "-"
End If

If IsNumeric(Dia(ItemAtual, 1)) Then
RAVColunas.DiasAguardandoAnalise = CInt(Dia(ItemAtual, 1))
Else
RAVColunas.DiasAguardandoAnalise = -1
End If

RAVColunas.DataInclusao = CDate(AnalistaeData(ItemAtual, 2))

If IsNumeric(PrazoMaximoAnalise(ItemAtual, 1)) Then
RAVColunas.PrazoMaximoAnalise = CInt(PrazoMaximoAnalise(ItemAtual, 1))
Else
RAVColunas.PrazoMaximoAnalise = -1
End If
RAVColunas.Status = Status(ItemAtual, 1)
RAVColunas.Tipo = Tipo(ItemAtual, 1)

DocDepositadoAtribuido.Add RAVColunas
Set RAVColunas = New cRAVColunasXL
Else
RAVColunas.Protocolo = DepositoDadosBasicos(ItemAtual, 1)
RAVColunas.Cliente = DepositoDadosBasicos(ItemAtual, 2)
RAVColunas.Fornecedor = DepositoDadosBasicos(ItemAtual, 3)
RAVColunas.Unidade = DepositoDadosBasicos(ItemAtual, 4)
RAVColunas.Documento = DepositoDadosBasicos(ItemAtual, 5)
RAVColunas.Empregado = Empregado(ItemAtual, 1)
RAVColunas.Analista = AnalistaeData(ItemAtual, 1)
RAVColunas.AnalistaEmail = "-"
RAVColunas.DataInclusao = CDate(AnalistaeData(ItemAtual, 2))

If IsNumeric(Dia(ItemAtual, 1)) Then
RAVColunas.DiasAguardandoAnalise = CInt(Dia(ItemAtual, 1))
Else
RAVColunas.DiasAguardandoAnalise = -1
End If

If IsNumeric(PrazoMaximoAnalise(ItemAtual, 1)) Then
RAVColunas.PrazoMaximoAnalise = CInt(PrazoMaximoAnalise(ItemAtual, 1))
Else
RAVColunas.PrazoMaximoAnalise = -1
End If
RAVColunas.Status = Status(ItemAtual, 1)
RAVColunas.Tipo = Tipo(ItemAtual, 1)

DocDepositadoSemAnalista.Add RAVColunas
Set RAVColunas = New cRAVColunasXL
End If
Next ItemAtual

QtdDocumentosSemAnalista = DocDepositadoSemAnalista.Count
GerarCopiaPlanilhaTemporia PastadeTrabalhoRAV, 1, False, NumLinha, RAVColunas

resposta = MsgBox("Deseja visualizar os e-mails antes de enviar-los ? " _
& "Caso não, serão enviados automáticamente.", vbYesNoCancel + vbQuestion)

If resposta = vbYes Or resposta = vbNo Then

NumLinha = DocDepositadoAtribuido.Count

If resposta = vbYes Then
Dim RAVEmail As CRAVEmailDocumento
Set RAVEmail = New CRAVEmailDocumento
End If

For Each AnalistaAtualEnvioEmail In AnalistasEncontrados.Keys

For AnalistaAtual = 1 To NumLinha
If DocDepositadoAtribuido(AnalistaAtual).Analista = AnalistaAtualEnvioEmail Then

If Not PlanilhaAnalista.Exists(AnalistaAtualEnvioEmail) Then
If Not PastadeTrabalhoCriada Is Nothing Then
PastadeTrabalhoCriada.Close SaveChanges:=True
Set PastadeTrabalhoCriada = Nothing
Set PlanilhaCriada = Nothing
Else
'PastadeTrabalhoCriada.Close SaveChanges:=True
Set PastadeTrabalhoCriada = Nothing
Set PlanilhaCriada = Nothing
End If
Criado = False
ArquivoCriado = Environ$("temp") & "\MEUS DOCUMENTOS - " & DocDepositadoAtribuido(AnalistaAtual).Analista & Format(Now(), "ddmmyyyyhhmmss") & ".xlsx"
Set PastadeTrabalhoCriada = Workbooks.Add
With PastadeTrabalhoCriada
PastadeTrabalhoCriada.SaveAs Filename:=ArquivoCriado
End With
PlanilhaAnalista.Add AnalistaAtualEnvioEmail, ArquivoCriado
PastadeTrabalhoCriada.Application.Calculation = xlCalculationManual
PastadeTrabalhoCriada.Application.ScreenUpdating = False
PastadeTrabalhoCriada.Application.EnableEvents = False
PastadeTrabalhoCriada.Application.AskToUpdateLinks = False
PastadeTrabalhoCriada.Application.DisplayAlerts = False

Set PlanilhaCriada = PastadeTrabalhoCriada.Worksheets(1)
PlanilhaCriada.Cells(1, 1).Value2 = "Protocolo"
PlanilhaCriada.Cells(1, 1).Interior.Color = preto
PlanilhaCriada.Cells(1, 1).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 2).Value2 = "Cliente"
PlanilhaCriada.Cells(1, 2).Interior.Color = preto
PlanilhaCriada.Cells(1, 2).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 3).Value2 = "Fornecedor"
PlanilhaCriada.Cells(1, 3).Interior.Color = preto
PlanilhaCriada.Cells(1, 3).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 4).Value2 = "Unidade"
PlanilhaCriada.Cells(1, 4).Interior.Color = preto
PlanilhaCriada.Cells(1, 4).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 5).Value2 = "Documento"
PlanilhaCriada.Cells(1, 5).Interior.Color = preto
PlanilhaCriada.Cells(1, 5).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 6).Value2 = "Empregado"
PlanilhaCriada.Cells(1, 6).Interior.Color = preto
PlanilhaCriada.Cells(1, 6).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 7).Value2 = "Analista"
PlanilhaCriada.Cells(1, 7).Interior.Color = preto
PlanilhaCriada.Cells(1, 7).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 8).Value2 = "Data de Depósito"
PlanilhaCriada.Cells(1, 8).Interior.Color = preto
PlanilhaCriada.Cells(1, 8).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 9).Value2 = "Dias em Análise"
PlanilhaCriada.Cells(1, 9).Interior.Color = preto
PlanilhaCriada.Cells(1, 9).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 10).Value2 = "Prazo Máximo de Análise"
PlanilhaCriada.Cells(1, 10).Interior.Color = preto
PlanilhaCriada.Cells(1, 10).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 11).Value2 = "Status"
PlanilhaCriada.Cells(1, 11).Interior.Color = preto
PlanilhaCriada.Cells(1, 11).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 12).Value2 = "Tipo"
PlanilhaCriada.Cells(1, 12).Interior.Color = preto
PlanilhaCriada.Cells(1, 12).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 15).Value2 = "Informações Adicionais"
PlanilhaCriada.Cells(1, 12).Interior.Color = preto
PlanilhaCriada.Cells(1, 12).Font.Color = vbWhite
PlanilhaCriada.Range("O1:Q1").Merge
PlanilhaCriada.Range("O1:Q1").Interior.Color = preto
PlanilhaCriada.Range("O1:Q1").Font.Color = vbWhite
PlanilhaCriada.Range("S1").Value2 = "Métricas"
PlanilhaCriada.Range("S1:U1").Merge
PlanilhaCriada.Range("S1:U1").Interior.Color = preto
PlanilhaCriada.Range("S1:U1").Font.Color = vbWhite
PlanilhaCriada.Range("O2").Value2 = "Tipo"
PlanilhaCriada.Range("P2").Value2 = "Dias"
PlanilhaCriada.Range("Q2").Value2 = "Cor"
PlanilhaCriada.Range("O3").Value2 = "Empresa"
PlanilhaCriada.Range("O4").Value2 = "Empresa"
PlanilhaCriada.Range("O5").Value2 = "Empresa"
PlanilhaCriada.Range("O6").Value2 = "Empresa"
PlanilhaCriada.Range("O7").Value2 = "Empresa"
PlanilhaCriada.Range("O8").Value2 = "Empresa"
PlanilhaCriada.Range("O9").Value2 = "Empresa"
PlanilhaCriada.Range("O10").Value2 = "Admissão"
PlanilhaCriada.Range("O11").Value2 = "Admissão"
PlanilhaCriada.Range("O12").Value2 = "Admissão"
PlanilhaCriada.Range("O13").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O14").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O15").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O16").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O17").Value2 = "ACT e CCT"
PlanilhaCriada.Range("P3").Value2 = "0"
PlanilhaCriada.Range("P4").Value2 = "1"
PlanilhaCriada.Range("P5").Value2 = "2"
PlanilhaCriada.Range("P6").Value2 = "3"
PlanilhaCriada.Range("P7").Value2 = "4"
PlanilhaCriada.Range("P8").Value2 = "5"
PlanilhaCriada.Range("P9").Value2 = ">= 6"
PlanilhaCriada.Range("P10").Value2 = "0"
PlanilhaCriada.Range("P11").Value2 = "1"
PlanilhaCriada.Range("P12").Value2 = ">= 2"
PlanilhaCriada.Range("P13").Value2 = "0"
PlanilhaCriada.Range("P14").Value2 = "1"
PlanilhaCriada.Range("P15").Value2 = "2"
PlanilhaCriada.Range("P16").Value2 = "3"
PlanilhaCriada.Range("P17").Value2 = ">= 4"
PlanilhaCriada.Range("Q3").Interior.Color = RGB(76, 175, 80)
PlanilhaCriada.Range("Q4").Interior.Color = RGB(217, 255, 0)
PlanilhaCriada.Range("Q5").Interior.Color = RGB(255, 255, 0)
PlanilhaCriada.Range("Q6").Interior.Color = RGB(255, 229, 0)
PlanilhaCriada.Range("Q7").Interior.Color = RGB(255, 120, 0)
PlanilhaCriada.Range("Q8").Interior.Color = RGB(255, 0, 0)
PlanilhaCriada.Range("Q9").Interior.Color = preto
PlanilhaCriada.Range("Q10").Interior.Color = RGB(76, 175, 80)
PlanilhaCriada.Range("Q11").Interior.Color = RGB(255, 0, 0)
PlanilhaCriada.Range("Q12").Interior.Color = preto
PlanilhaCriada.Range("Q13").Interior.Color = RGB(76, 175, 80)
PlanilhaCriada.Range("Q14").Interior.Color = RGB(255, 220, 153)
PlanilhaCriada.Range("Q15").Interior.Color = RGB(255, 158, 0)
PlanilhaCriada.Range("Q16").Interior.Color = RGB(255, 0, 0)
PlanilhaCriada.Range("Q17").Interior.Color = preto
PlanilhaCriada.Range("O:Q").HorizontalAlignment = xlCenter
PlanilhaCriada.Range("O:Q").VerticalAlignment = xlCenter
PlanilhaCriada.Range("S2").Value2 = "Status do Documento"
PlanilhaCriada.Range("T2").Value2 = "Quantidade"
PlanilhaCriada.Range("U2").Value2 = "Porcentagem"
PlanilhaCriada.Range("S3").Value2 = "Em dia"
PlanilhaCriada.Range("S4").Value2 = "Prazo Fatal"
PlanilhaCriada.Range("S5").Value2 = "Atrasado"
PlanilhaCriada.Range("S:U").HorizontalAlignment = xlCenter
PlanilhaCriada.Range("S:U").VerticalAlignment = xlCenter
LinhaAtual = 2
Criado = True
End If

RAVColunas.Protocolo = DocDepositadoAtribuido(AnalistaAtual).Protocolo
PlanilhaCriada.Cells(LinhaAtual, 1).Value2 = DocDepositadoAtribuido(AnalistaAtual).Protocolo
RAVColunas.Cliente = DocDepositadoAtribuido(AnalistaAtual).Cliente
PlanilhaCriada.Cells(LinhaAtual, 2).Value2 = DocDepositadoAtribuido(AnalistaAtual).Cliente
RAVColunas.Fornecedor = DocDepositadoAtribuido(AnalistaAtual).Fornecedor
PlanilhaCriada.Cells(LinhaAtual, 3).Value2 = DocDepositadoAtribuido(AnalistaAtual).Fornecedor
RAVColunas.Unidade = DocDepositadoAtribuido(AnalistaAtual).Unidade
PlanilhaCriada.Cells(LinhaAtual, 4).Value2 = DocDepositadoAtribuido(AnalistaAtual).Unidade
RAVColunas.Documento = DocDepositadoAtribuido(AnalistaAtual).Documento
PlanilhaCriada.Cells(LinhaAtual, 5).Value2 = DocDepositadoAtribuido(AnalistaAtual).Documento
RAVColunas.Empregado = DocDepositadoAtribuido(AnalistaAtual).Empregado
PlanilhaCriada.Cells(LinhaAtual, 6).Value2 = DocDepositadoAtribuido(AnalistaAtual).Empregado
RAVColunas.Analista = DocDepositadoAtribuido(AnalistaAtual).Analista
PlanilhaCriada.Cells(LinhaAtual, 7).Value2 = DocDepositadoAtribuido(AnalistaAtual).Analista
RAVColunas.AnalistaEmail = DocDepositadoAtribuido(AnalistaAtual).AnalistaEmail
RAVColunas.DataInclusao = DocDepositadoAtribuido(AnalistaAtual).DataInclusao
PlanilhaCriada.Cells(LinhaAtual, 8).Value2 = DocDepositadoAtribuido(AnalistaAtual).DataInclusao
RAVColunas.DiasAguardandoAnalise = DocDepositadoAtribuido(AnalistaAtual).DiasAguardandoAnalise
PlanilhaCriada.Cells(LinhaAtual, 9).Value2 = DocDepositadoAtribuido(AnalistaAtual).DiasAguardandoAnalise
RAVColunas.PrazoMaximoAnalise = DocDepositadoAtribuido(AnalistaAtual).PrazoMaximoAnalise
PlanilhaCriada.Cells(LinhaAtual, 10).Value2 = DocDepositadoAtribuido(AnalistaAtual).PrazoMaximoAnalise
RAVColunas.Status = DocDepositadoAtribuido(AnalistaAtual).Status
PlanilhaCriada.Cells(LinhaAtual, 11).Value2 = DocDepositadoAtribuido(AnalistaAtual).Status
RAVColunas.Tipo = DocDepositadoAtribuido(AnalistaAtual).Tipo
PlanilhaCriada.Cells(LinhaAtual, 12).Value2 = DocDepositadoAtribuido(AnalistaAtual).Tipo
TingePrazo PlanilhaCriada, Range("A" & LinhaAtual & ":L" & LinhaAtual), _
IIf(RAVColunas.Tipo = "COMUM", 0, 1), DocDepositadoAtribuido(AnalistaAtual).DiasAguardandoAnalise, _
DocDepositadoAtribuido(AnalistaAtual).Documento
If RAVColunas.Status = "Atrasado" Then
QtdPrazoPerdido = QtdPrazoPerdido + 1
ElseIf RAVColunas.Status = "Prazo Fatal" Then
QtdPrazoFatal = QtdPrazoFatal + 1
Else
QtdEmDia = QtdEmDia + 1
End If

LinhaAtual = LinhaAtual + 1
DocumentosparaEnvioAnalista.Add RAVColunas

 
Set RAVColunas = New cRAVColunasXL
End If
Next AnalistaAtual



If DocumentosparaEnvioAnalista.Count = 0 Then
Set DocumentosparaEnvioAnalista = Nothing
Set RAVColunas = New cRAVColunasXL
GoTo NovoAnalista
End If
        
        PlanilhaCriada.Range("T3").Formula = "=COUNTIF(K:K,""Em dia"")"
        PlanilhaCriada.Range("T4").Formula = "=COUNTIF(K:K,""Prazo Fatal"")"
        PlanilhaCriada.Range("T5").Formula = "=COUNTIF(K:K,""Atrasado"")"
        PlanilhaCriada.Range("U3").Formula = "=T3/COUNTIF(A:A,"">0"")"
        PlanilhaCriada.Range("U4").Formula = "=T4/COUNTIF(A:A,"">0"")"
        PlanilhaCriada.Range("U5").Formula = "=T5/COUNTIF(A:A,"">0"")"
        PlanilhaCriada.Range("U:U").NumberFormat = "0.00%"
        Set PlanilhaCriadaLinhas = PlanilhaCriada.Range("A1").CurrentRegion
        PlanilhaCriada.Range(PlanilhaCriadaLinhas.Address).Sort Key1:=Range("L1"), Order1:=xlAscending, Header:=xlYes
        PlanilhaCriada.Columns("A:L").AutoFit
        PlanilhaCriada.Columns("S:U").AutoFit
        PastadeTrabalhoCriada.Save
        
        TabelaHTML = GerarTabelaDocumentosHTML(1, DocumentosparaEnvioAnalista)
        EmailPara = DocumentosparaEnvioAnalista(1).AnalistaEmail
        If resposta = vbYes Then
        If Not EmailsCriados.Exists(EmailPara) Then
        EmailsCriados.Add EmailPara, TabelaHTML
        End If
        End If
        For lider = 0 To QtdLideres
        If QtdLideres <> lider Then
        AnalistaNome = Left(Lideres(0, lider), InStr(Lideres(0, lider), " ") - 1)
        EmailLiderNome = AnalistaNome & " e "
        EmailCopia = EmailCopia & Lideres(1, lider) & ";"
        Else
        AnalistaNome = Left(Lideres(0, lider), InStr(Lideres(0, lider), " ") - 1)
        EmailLiderNome = EmailLiderNome & AnalistaNome
        EmailCopia = EmailCopia & Lideres(1, lider)
        End If
        Next lider
        
        EmailAssunto = "E-MAIL AUTOMÁTICO - Relatório de Documentos Aguardando Validação. Você possui " & QtdEmDia & "" _
        & " documento(s) em dia " & QtdPrazoFatal & " documento(s) em prazo fatal e " & QtdPrazoPerdido & " documento(s) atrasado(s)"
        
        SomaQtdPrazoFatal = SomaQtdPrazoFatal + QtdPrazoFatal
        SomaQtdPrazoPerdido = SomaQtdPrazoFatal + QtdPrazoPerdido
        SomaQtdEmDia = SomaQtdEmDia + QtdEmDia
        
        Dim momento, HoraPlanilha As String
        Dim Hora As Date: Hora = TimeValue(Now())
        If (Hora < TimeValue("12:00") And Hora >= TimeValue("06:00")) Then
        momento = "bom dia"
        ElseIf (Hora >= TimeValue("13:00") And Hora <= TimeValue("17:59")) Then
        momento = "boa tarde"
        Else
        momento = "boa noite"
        End If
        
      
        HoraPlanilha = HorarioCriacaoArquivo(PastadeTrabalhoRAV.FullName)
       
      AnalistaNome = Left(DocumentosparaEnvioAnalista(1).Analista, InStr(DocumentosparaEnvioAnalista(1).Analista, " ") - 1)
        
        If TabelaHTML <> "" Then
        EmailMensagem = _
         "<h3 style=text-align:center;>RELATÓRIO DE DOCUMENTOS AGUARDANDO VALIDAÇÃO</h3> " _
        & " Olá, " & momento & ", " & AnalistaNome & "! Tudo bem ? <br><br>  Precisamos que você analise ainda hoje, " _
        & "os seguintes documentos vinculados a você: " & TabelaHTML & " <br><br> " _
        & " Contamos com a sua dedicação e comprometimento para juntos alcarçarmos a meta do dia." _
        & "<br><br> " & EmailLiderNome & " <br><br> Horário de emissão do relatório: " _
        & HoraPlanilha & ""
        Else
         EmailMensagem = _
         "<h3 style=text-align:center;>RELATÓRIO DE DOCUMENTOS AGUARDANDO VALIDAÇÃO</h3> " _
        & " Olá, " & momento & ", " & AnalistaNome & "! Tudo bem ? <br><br>  Distribuímos alguns documentos para você, " _
        & "eles estão contidos na planilha anexa. Contamos com sua ajuda! " _
        & "<br><br> " & EmailLiderNome & " <br><br> Horário de emissão do relatório: " _
        & HoraPlanilha & ""
        End If
        
        
        If resposta = vbNo Then
        EnvioEmail EmailPara, EmailCopia, EmailAssunto, EmailMensagem, , , , PlanilhaAnalista.Item(AnalistaAtualEnvioEmail)
        On Error Resume Next
        PastadeTrabalhoCriada.Close SaveChanges:=True
        Set PastadeTrabalhoCriada = Nothing
        Kill (PlanilhaAnalista.Item(AnalistaAtualEnvioEmail))
        If Err.Number > 0 Then
        MsgBox "Não foi possível deletar a planilha temporária, contacte o desenvolvedor, Weverson Rafael Moreira. Aconteceu o seguinte" _
        & " erro " & Err.Description, vbCritical
        Err.Clear
        End If
        Else
        RAVEmail.Para = EmailPara
        RAVEmail.Copia = EmailCopia
        RAVEmail.Assunto = EmailAssunto
        RAVEmail.Mensagem = EmailMensagem
        RAVEmail.Anexo = PlanilhaAnalista.Item(AnalistaAtualEnvioEmail)
        EmailsParaVisualizacao.Add RAVEmail
        Set RAVEmail = New CRAVEmailDocumento
        End If
        Set DocumentosparaEnvioAnalista = Nothing
        Set RAVColunas = New cRAVColunasXL
        TabelaHTML = ""
        EmailLiderNome = ""
        EmailCopia = ""
        QtdPrazoFatal = 0
        QtdPrazoPerdido = 0
        QtdEmDia = 0
NovoAnalista:
Next AnalistaAtualEnvioEmail

PastadeTrabalhoCriada.Close SaveChanges:=True
Set PastadeTrabalhoCriada = Nothing

If QtdDocumentosSemAnalista > 0 Then
        For EmailAtual = 0 To QtdResultadosConsulta
        If EmailAtual <> QtdResultadosConsulta Then
            EmailPara = EmailPara & Consulta(1, EmailAtual) & ";"
        Else
            EmailPara = EmailPara & Consulta(1, EmailAtual)
        End If
        Next EmailAtual
End If
         
If DocDepositadoSemAnalista.Count > 0 Then
Dim PlanilhaAnexa As String
If DocDepositadoSemAnalista.Count > 750 Then DoEvents
If Not PastadeTrabalhoCriada Is Nothing Then
PastadeTrabalhoCriada.Close SaveChanges:=True
Set PastadeTrabalhoCriada = Nothing
Set PlanilhaCriada = Nothing
Else
'PastadeTrabalhoCriada.Close SaveChanges:=True
Set PastadeTrabalhoCriada = Nothing
Set PlanilhaCriada = Nothing
End If
Criado = False
ArquivoCriado = Environ$("temp") & "\Documentação - " & RAVPreferencias(1).celula & " - " & Format(Now(), "ddmmyyyyhhmmss") & ".xlsx"
Set PastadeTrabalhoCriada = Workbooks.Add
With PastadeTrabalhoCriada
PastadeTrabalhoCriada.SaveAs Filename:=ArquivoCriado
PlanilhaAnexa = PastadeTrabalhoCriada.FullName
End With
PlanilhaAnalista.Add "Documentação - " & RAVPreferencias(1).celula & " - " & Format(Now(), "ddmmyyyyhhmmss"), ArquivoCriado
PastadeTrabalhoCriada.Application.Calculation = xlCalculationManual
PastadeTrabalhoCriada.Application.ScreenUpdating = False
PastadeTrabalhoCriada.Application.EnableEvents = False
PastadeTrabalhoCriada.Application.AskToUpdateLinks = False
PastadeTrabalhoCriada.Application.DisplayAlerts = False
'PastadeTrabalhoCriada.Application.EnableEvents = False
Set PlanilhaCriada = PastadeTrabalhoCriada.Worksheets(1)

PlanilhaCriada.Cells(1, 1).Value2 = "Protocolo"
PlanilhaCriada.Cells(1, 1).Interior.Color = preto
PlanilhaCriada.Cells(1, 1).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 2).Value2 = "Cliente"
PlanilhaCriada.Cells(1, 2).Interior.Color = preto
PlanilhaCriada.Cells(1, 2).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 3).Value2 = "Fornecedor"
PlanilhaCriada.Cells(1, 3).Interior.Color = preto
PlanilhaCriada.Cells(1, 3).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 4).Value2 = "Unidade"
PlanilhaCriada.Cells(1, 4).Interior.Color = preto
PlanilhaCriada.Cells(1, 4).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 5).Value2 = "Documento"
PlanilhaCriada.Cells(1, 5).Interior.Color = preto
PlanilhaCriada.Cells(1, 5).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 6).Value2 = "Empregado"
PlanilhaCriada.Cells(1, 6).Interior.Color = preto
PlanilhaCriada.Cells(1, 6).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 7).Value2 = "Analista"
PlanilhaCriada.Cells(1, 7).Interior.Color = preto
PlanilhaCriada.Cells(1, 7).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 8).Value2 = "Data de Depósito"
PlanilhaCriada.Cells(1, 8).Interior.Color = preto
PlanilhaCriada.Cells(1, 8).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 9).Value2 = "Dias em Análise"
PlanilhaCriada.Cells(1, 9).Interior.Color = preto
PlanilhaCriada.Cells(1, 9).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 10).Value2 = "Prazo Máximo de Análise"
PlanilhaCriada.Cells(1, 10).Interior.Color = preto
PlanilhaCriada.Cells(1, 10).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 11).Value2 = "Status"
PlanilhaCriada.Cells(1, 11).Interior.Color = preto
PlanilhaCriada.Cells(1, 11).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 12).Value2 = "Tipo"
PlanilhaCriada.Cells(1, 12).Interior.Color = preto
PlanilhaCriada.Cells(1, 12).Font.Color = vbWhite
PlanilhaCriada.Cells(1, 15).Value2 = "Informações Adicionais"
PlanilhaCriada.Cells(1, 12).Interior.Color = preto
PlanilhaCriada.Cells(1, 12).Font.Color = vbWhite
PlanilhaCriada.Range("O1:Q1").Merge
PlanilhaCriada.Range("O1:Q1").Interior.Color = preto
PlanilhaCriada.Range("O1:Q1").Font.Color = vbWhite
PlanilhaCriada.Range("S1").Value2 = "Métricas"
PlanilhaCriada.Range("S1:U1").Merge
PlanilhaCriada.Range("S1:U1").Interior.Color = preto
PlanilhaCriada.Range("S1:U1").Font.Color = vbWhite
PlanilhaCriada.Range("O2").Value2 = "Tipo"
PlanilhaCriada.Range("P2").Value2 = "Dias"
PlanilhaCriada.Range("Q2").Value2 = "Cor"
PlanilhaCriada.Range("O3").Value2 = "Empresa"
PlanilhaCriada.Range("O4").Value2 = "Empresa"
PlanilhaCriada.Range("O5").Value2 = "Empresa"
PlanilhaCriada.Range("O6").Value2 = "Empresa"
PlanilhaCriada.Range("O7").Value2 = "Empresa"
PlanilhaCriada.Range("O8").Value2 = "Empresa"
PlanilhaCriada.Range("O9").Value2 = "Empresa"
PlanilhaCriada.Range("O10").Value2 = "Admissão"
PlanilhaCriada.Range("O11").Value2 = "Admissão"
PlanilhaCriada.Range("O12").Value2 = "Admissão"
PlanilhaCriada.Range("O13").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O14").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O15").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O16").Value2 = "ACT e CCT"
PlanilhaCriada.Range("O17").Value2 = "ACT e CCT"
PlanilhaCriada.Range("P3").Value2 = "0"
PlanilhaCriada.Range("P4").Value2 = "1"
PlanilhaCriada.Range("P5").Value2 = "2"
PlanilhaCriada.Range("P6").Value2 = "3"
PlanilhaCriada.Range("P7").Value2 = "4"
PlanilhaCriada.Range("P8").Value2 = "5"
PlanilhaCriada.Range("P9").Value2 = ">= 6"
PlanilhaCriada.Range("P10").Value2 = "0"
PlanilhaCriada.Range("P11").Value2 = "1"
PlanilhaCriada.Range("P12").Value2 = ">= 2"
PlanilhaCriada.Range("P13").Value2 = "0"
PlanilhaCriada.Range("P14").Value2 = "1"
PlanilhaCriada.Range("P15").Value2 = "2"
PlanilhaCriada.Range("P16").Value2 = "3"
PlanilhaCriada.Range("P17").Value2 = ">= 4"
PlanilhaCriada.Range("Q3").Interior.Color = RGB(76, 175, 80)
PlanilhaCriada.Range("Q4").Interior.Color = RGB(217, 255, 0)
PlanilhaCriada.Range("Q5").Interior.Color = RGB(255, 255, 0)
PlanilhaCriada.Range("Q6").Interior.Color = RGB(255, 229, 0)
PlanilhaCriada.Range("Q7").Interior.Color = RGB(255, 120, 0)
PlanilhaCriada.Range("Q8").Interior.Color = RGB(255, 0, 0)
PlanilhaCriada.Range("Q9").Interior.Color = RGB(89, 89, 89)
PlanilhaCriada.Range("Q10").Interior.Color = RGB(76, 175, 80)
PlanilhaCriada.Range("Q11").Interior.Color = RGB(255, 0, 0)
PlanilhaCriada.Range("Q12").Interior.Color = RGB(97, 97, 97)
PlanilhaCriada.Range("Q13").Interior.Color = RGB(76, 175, 80)
PlanilhaCriada.Range("Q14").Interior.Color = RGB(255, 220, 153)
PlanilhaCriada.Range("Q15").Interior.Color = RGB(255, 158, 0)
PlanilhaCriada.Range("Q16").Interior.Color = RGB(255, 0, 0)
PlanilhaCriada.Range("Q17").Interior.Color = RGB(97, 97, 97)
PlanilhaCriada.Range("O:Q").HorizontalAlignment = xlCenter
PlanilhaCriada.Range("O:Q").VerticalAlignment = xlCenter
PlanilhaCriada.Range("S2").Value2 = "Status do Documento"
PlanilhaCriada.Range("S3").Value2 = "Em dia"
PlanilhaCriada.Range("S4").Value2 = "Prazo Fatal"
PlanilhaCriada.Range("S5").Value2 = "Atrasado"
PlanilhaCriada.Range("T2").Value2 = "Quantidade"
PlanilhaCriada.Range("U2").Value2 = "Porcentagem"
PlanilhaCriada.Range("S:U").HorizontalAlignment = xlCenter
PlanilhaCriada.Range("S:U").VerticalAlignment = xlCenter
LinhaAtual = 2
Criado = True


For Each Item In DocDepositadoSemAnalista

PlanilhaCriada.Cells(LinhaAtual, 1).Value2 = Item.Protocolo
PlanilhaCriada.Cells(LinhaAtual, 2).Value2 = Item.Cliente
PlanilhaCriada.Cells(LinhaAtual, 3).Value2 = Item.Fornecedor
PlanilhaCriada.Cells(LinhaAtual, 4).Value2 = Item.Unidade
PlanilhaCriada.Cells(LinhaAtual, 5).Value2 = Item.Documento
PlanilhaCriada.Cells(LinhaAtual, 6).Value2 = Item.Empregado
PlanilhaCriada.Cells(LinhaAtual, 7).Value2 = Item.Analista
PlanilhaCriada.Cells(LinhaAtual, 8).Value2 = Item.DataInclusao
PlanilhaCriada.Cells(LinhaAtual, 9).Value2 = Item.DiasAguardandoAnalise
PlanilhaCriada.Cells(LinhaAtual, 10).Value2 = Item.PrazoMaximoAnalise
PlanilhaCriada.Cells(LinhaAtual, 11).Value2 = Item.Status
PlanilhaCriada.Cells(LinhaAtual, 12).Value2 = Item.Tipo
TingePrazo PlanilhaCriada, Range("A" & LinhaAtual & ":L" & LinhaAtual), _
IIf(Item.Tipo = "COMUM", 0, 1), Item.DiasAguardandoAnalise, _
Item.Documento
If Item.Status = "Atrasado" Then
QtdPrazoPerdido = QtdPrazoPerdido + 1
ElseIf Item.Status = "Prazo Fatal" Then
QtdPrazoFatal = QtdPrazoFatal + 1
Else
QtdEmDia = QtdEmDia + 1
End If

'PastadeTrabalhoCriada.Save
LinhaAtual = LinhaAtual + 1
'DocumentosparaEnvioAnalista.Add RAVColunas
Set RAVColunas = New cRAVColunasXL
Next Item



        PlanilhaCriada.Range("T3").Formula = "=COUNTIF(K:K,""Em dia"")"
        PlanilhaCriada.Range("T4").Formula = "=COUNTIF(K:K,""Prazo Fatal"")"
        PlanilhaCriada.Range("T5").Formula = "=COUNTIF(K:K,""Atrasado"")"
        PlanilhaCriada.Range("U3").Formula = "=T3/COUNTIF(A:A,"">0"")"
        PlanilhaCriada.Range("U4").Formula = "=T4/COUNTIF(A:A,"">0"")"
        PlanilhaCriada.Range("U5").Formula = "=T5/COUNTIF(A:A,"">0"")"
        PlanilhaCriada.Range("U:U").NumberFormat = "0.00%"
        Set PlanilhaCriadaLinhas = PlanilhaCriada.Range("A1").CurrentRegion
        PlanilhaCriada.Range(PlanilhaCriadaLinhas.Address).Sort Key1:=Range("L1"), Order1:=xlAscending, Header:=xlYes
        PlanilhaCriada.Columns("A:L").AutoFit
        PlanilhaCriada.Columns("S:U").AutoFit
        PastadeTrabalhoCriada.Close SaveChanges:=True
End If
        
        
        EmailCopia = ""
        EmailAssunto = "E-MAIL AUTOMÁTICO - Relatório de Documentos Aguardando Validação. Documentos não distribuídos"
        EmailMensagem = "<h3 style=text-align:center;>RELATÓRIO DE DOCUMENTOS AGUARDANDO VALIDAÇÃO</h3>" _
        & "Olá, " & momento & " analistas! Existe(m) " & DocDepositadoSemAnalista.Count & " documento(s) que ainda não foi(ram)" _
        & "distribuídos. Veja a tabela em anexo. Horário de emissão do relatório: " & HoraPlanilha & ""
        
        If resposta = vbNo Then
        EnvioEmail EmailPara, EmailCopia, EmailAssunto, EmailMensagem, , , , PlanilhaAnexa
        End If
        If resposta = vbYes Then
        RAVEmail.Para = EmailPara
        RAVEmail.Copia = EmailCopia
        RAVEmail.Assunto = EmailAssunto
        RAVEmail.Mensagem = EmailMensagem
        RAVEmail.Anexo = PlanilhaAnexa
        'RAVEmail.AnalistaNome = "E-mail de conhecimento geral"
     
        Dim mqe As Variant
        EmailsCriados.Add "E-mail de conhecimento geral", EmailPara
        EmailsParaVisualizacao.Add RAVEmail
        Set RAVEmail = New CRAVEmailDocumento
     
        FrmDisparoEmail.Show
        End If
        
        MsgBox "Todos e-mails enviados : )", vbInformation
        Set EmailsParaVisualizacao = Nothing
        Set EmailsCriados = Nothing
        Set PlanilhaAtiva = Nothing
        
        PastadeTrabalhoRAV.Application.ScreenUpdating = True
        PastadeTrabalhoRAV.Application.Calculation = xlCalculationAutomatic
        PastadeTrabalhoRAV.Application.EnableEvents = True
        PastadeTrabalhoRAV.Application.DisplayAlerts = True
        Set SelecaoFeriados = Nothing
End If

'Else
PastadeTrabalhoRAV.Application.ScreenUpdating = True
PastadeTrabalhoRAV.Application.Calculation = xlCalculationAutomatic
PastadeTrabalhoRAV.Application.EnableEvents = True
PastadeTrabalhoRAV.Application.DisplayAlerts = True
Set SelecaoFeriados = Nothing
'End If
Exit Function
Catch:
MsgBox "Algo deu errado. " & Err.Description & " " & Err.Source, vbCritical
End Function
