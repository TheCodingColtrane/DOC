Attribute VB_Name = "FO"
Option Explicit

Function CPDOD(DiaDepositado As Excel.Range, Optional FeriadosSelecionados As Excel.Range, Optional DiaLimite As Excel.Range) As Long


Dim QtdDiasAguardandoValidacao As Long
Dim QtdDiaUtilPosFeriado As Integer
QtdDiasAguardandoValidacao = 0
Dim DataDeposito, DataLimite, DataVerificada, ChecagemDiaUtil As Date

Dim FoiDepositadoEmFinalSemana, FoiDepositadoEmFeriado, EDiaZero As Boolean


If Not FeriadosSelecionados Is Nothing Then
Dim Linha As Range
Dim Feriados As New Dictionary
For Each Linha In FeriadosSelecionados
Feriados.Add CDate(Linha.Value2), Feriados.Count + 1
Next Linha

If Not DiaLimite Is Nothing Then
DataLimite = CDate(Int(DiaLimite))
Else
DataLimite = Date
End If


FoiDepositadoEmFeriado = False

DataDeposito = CDate(Int(DiaDepositado.Value2))

If Feriados.Exists(DataDeposito) Then
FoiDepositadoEmFeriado = True
EDiaZero = True
End If

If Weekday(DataDeposito) = vbSaturday Or Weekday(DataDeposito) = vbSunday Then
FoiDepositadoEmFinalSemana = True
EDiaZero = True
End If


For DataVerificada = DataDeposito To DataLimite

If FoiDepositadoEmFeriado = True Then

If Not Feriados.Exists(DataVerificada) _
And (Weekday(DataVerificada) <> vbSaturday) _
And (Weekday(DataVerificada) <> vbSunday) _
And EDiaZero = True Then
QtdDiaUtilPosFeriado = 1
ChecagemDiaUtil = DataVerificada
End If

End If

If Not Feriados.Exists(DataVerificada) _
And (Weekday(DataVerificada) <> vbSaturday) _
And (Weekday(DataVerificada) <> vbSunday) _
And DataVerificada <> DataDeposito _
And QtdDiaUtilPosFeriado = 0 _
And EDiaZero = False Then

QtdDiasAguardandoValidacao = QtdDiasAguardandoValidacao + 1
End If

If FoiDepositadoEmFinalSemana = True And EDiaZero = True Then
If Weekday(DataVerificada) <> vbSaturday And Weekday(DataVerificada) <> vbSunday Then
EDiaZero = False
End If
End If

If FoiDepositadoEmFeriado = True And DataVerificada = ChecagemDiaUtil Then

QtdDiaUtilPosFeriado = QtdDiaUtilPosFeriado - 1
EDiaZero = False
If QtdDiaUtilPosFeriado = 0 Then
QtdDiaUtilPosFeriado = 0
End If

End If

Next DataVerificada

CPDOD = QtdDiasAguardandoValidacao
Else

If Not DiaLimite Is Nothing Then
DataLimite = CDate(Int(DiaLimite))
Else
DataLimite = Date
End If

DataDeposito = CDate(Int(DiaDepositado.Value2))

If Weekday(DataDeposito) = vbSaturday Or Weekday(DataDeposito) = vbSunday Then
FoiDepositadoEmFinalSemana = True
EDiaZero = True

Else
FoiDepositadoEmFinalSemana = False
EDiaZero = False

End If

For DataVerificada = DataDeposito To DataLimite

If Weekday(DataVerificada) <> vbSaturday _
And Weekday(DataVerificada) <> vbSunday _
And DataVerificada <> DataDeposito _
And QtdDiaUtilPosFeriado = 0 _
And EDiaZero = False Then

QtdDiasAguardandoValidacao = QtdDiasAguardandoValidacao + 1
End If

If FoiDepositadoEmFinalSemana = True And EDiaZero = True Then

If Weekday(DataVerificada) <> vbSaturday And Weekday(DataVerificada) <> vbSunday Then
EDiaZero = False
End If
End If

'End If

Next DataVerificada
End If

CPDOD = QtdDiasAguardandoValidacao
End Function

Function teste(num1 As Excel.Range) As Integer
Dim num1a As Integer: num1a = num1.Value
num1a = num1a + 70
teste = num1a
End Function

Private Sub RegistroFormula()
    Application.MacroOptions _
        Macro:="CPDOD", _
        Category:=14, _
        Description:="Retorna o Cálculo de Prazo de Documento Operacional Depositado", _
        ArgumentDescriptions:=Array( _
            "data_deposito.  Data em que o documento foi depositado", _
         "Feriados.  Intervalo de Feriados trabalhados. Não é necessário fornecer sábados e domingos.")
End Sub

