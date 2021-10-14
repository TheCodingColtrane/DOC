Attribute VB_Name = "Distrib"
Option Explicit

Public Function AutoDistrib(AnalistaCargoComplexidade As Variant, QtdAnalistas As Variant, _
QtdDocumentos As Variant, DocumentosNivel1 As Dictionary, DocumentosNivel2 As Dictionary, _
DocumentosNivel3 As Dictionary, DocumentosNivel4 As Dictionary, DocumentosNivel5 As Dictionary, _
Modo As Integer, PlanilhaRav As Excel.Worksheet, RAVColunas As cRAVColunasXL)

Dim QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, AnalistaAtual, QtdAnalistasCompensados, QtdDocumentosCompensados, _
QtdDocumentosReservados, Indice, Limite, LimiteAuxiliar, IndiceNivel1, IndiceNivel2, IndiceNivel3, IndiceNivel4, _
QtdDocumentosAuxiliar, QtdAnalistaNivel1, QtdAnalistaNivel2, QtdAnalistaNivel3, QtdAnalistaNivel4  As Integer
Dim LinhaAtual, AnalistaNivel1, AnalistaNivel2, AnalistaNivel3, AnalistaNivel4, AnalistaNivel5 As Variant
Dim QtdDocumentosNivel1PorNivelAnalista As New Dictionary
Dim QtdDocumentosNivel2PorNivelAnalista As New Dictionary
Dim QtdDocumentosNivel3PorNivelAnalista As New Dictionary
Dim QtdDocumentosNivel4PorNivelAnalista As New Dictionary
Dim QtdDocumentosNivel5PorNivelAnalista As New Dictionary


ReDim AnalistaNivel1(QtdAnalistas(0) - 1)
ReDim AnalistaNivel2(QtdAnalistas(1) - 1)
ReDim AnalistaNivel3(QtdAnalistas(2) - 1)
ReDim AnalistaNivel4(QtdAnalistas(3) - 1)
QtdAnalistas = UBound(AnalistaCargoComplexidade, 2)
For Indice = 0 To QtdAnalistas

If AnalistaCargoComplexidade(1, Indice) = 2 Then
AnalistaNivel1(IndiceNivel1) = AnalistaCargoComplexidade(0, Indice)
IndiceNivel1 = IndiceNivel1 + 1
ElseIf AnalistaCargoComplexidade(1, Indice) = 3 Then
AnalistaNivel2(IndiceNivel2) = AnalistaCargoComplexidade(0, Indice)
IndiceNivel2 = IndiceNivel2 + 1
ElseIf AnalistaCargoComplexidade(1, Indice) = 4 Then
AnalistaNivel3(IndiceNivel3) = AnalistaCargoComplexidade(0, Indice)
IndiceNivel3 = IndiceNivel3 + 1
Else
AnalistaNivel4(IndiceNivel4) = AnalistaCargoComplexidade(0, Indice)
IndiceNivel4 = IndiceNivel4 + 1
End If


Next Indice


QtdAnalistaNivel1 = UBound(AnalistaNivel1) + 1
QtdAnalistaNivel2 = UBound(AnalistaNivel2) + 1
QtdAnalistaNivel3 = UBound(AnalistaNivel3) + 1
QtdAnalistaNivel4 = UBound(AnalistaNivel4) + 1

If Modo = 1 Then
QtdAnalistasNivelAtual = Int(QtdAnalistas(0))
QtdDocumentosNivelAtual = DocumentosNivel1.Count
QtdDocumentosReservados = CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.25)
QtdDocumentosNivel1PorNivelAnalista.Add 1, QtdDocumentosReservados
QtdAnalistasCompensados = QtdAnalistas(1) + QtdAnalistas(2) + QtdAnalistas(3)
QtdDocumentosCompensados = QtdDocumentosNivelAtual - QtdDocumentosNivel1PorNivelAnalista.Item(1) * QtdAnalistasNivelAtual
QtdAnalistasNivelAtual = QtdAnalistas(1)
QtdDocumentosNivel1PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosCompensados, 0.4) * QtdAnalistasNivelAtual
QtdDocumentosCompensados = QtdDocumentosCompensados - QtdDocumentosNivel1PorNivelAnalista.Item(2)
QtdAnalistasNivelAtual = QtdAnalistas(2)
QtdDocumentosNivel1PorNivelAnalista.Add 3, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosCompensados, 0.5) * QtdAnalistasNivelAtual
QtdDocumentosCompensados = QtdDocumentosCompensados - QtdDocumentosNivel1PorNivelAnalista.Item(3)
QtdAnalistasNivelAtual = QtdAnalistas(3)
QtdDocumentosNivel1PorNivelAnalista.Add 4, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosCompensados, 1) * QtdAnalistasNivelAtual


Limite = QtdDocumentosNivel1PorNivelAnalista.Item(1) * QtdAnalistaNivel1
LimiteAuxiliar = Limite
'Distribuição Nível 1. Possibilidade de distribuição universal.
CalculoDistribIgualitaria DocumentosNivel1, AnalistaNivel1, 0, Limite, PlanilhaRav, RAVColunas
Limite = QtdDocumentosNivel1PorNivelAnalista.Item(1) * QtdAnalistaNivel1 + QtdDocumentosNivel1PorNivelAnalista.Item(2)
CalculoDistribIgualitaria DocumentosNivel1, AnalistaNivel2, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite
Limite = Limite + QtdDocumentosNivel1PorNivelAnalista.Item(3)
CalculoDistribIgualitaria DocumentosNivel1, AnalistaNivel3, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite
Limite = Limite + QtdDocumentosNivel1PorNivelAnalista.Item(4)
CalculoDistribIgualitaria DocumentosNivel1, AnalistaNivel4, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas

'Preparação de Distribuição Nível 2. Preparação das respectivas quotas.
QtdAnalistasNivelAtual = QtdAnalistaNivel1
QtdDocumentosNivelAtual = DocumentosNivel2.Count
QtdDocumentosNivel2PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.35) * QtdAnalistaNivel1
QtdAnalistasNivelAtual = QtdAnalistaNivel2
QtdDocumentosNivel2PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.35) * QtdAnalistaNivel2
QtdDocumentosNivelAtual = QtdDocumentosNivelAtual - (QtdDocumentosNivel2PorNivelAnalista.Item(1) + QtdDocumentosNivel2PorNivelAnalista.Item(2))
QtdAnalistasNivelAtual = QtdAnalistaNivel3
QtdDocumentosAuxiliar = QtdDocumentosNivel2PorNivelAnalista.Item(1) + QtdDocumentosNivel2PorNivelAnalista.Item(2)
QtdDocumentosCompensados = DocumentosNivel2.Count - QtdDocumentosAuxiliar
QtdDocumentosNivel2PorNivelAnalista.Add 3, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.6) * QtdAnalistaNivel3
QtdDocumentosNivelAtual = DocumentosNivel2.Count - (QtdDocumentosNivel2PorNivelAnalista.Item(1) + QtdDocumentosNivel2PorNivelAnalista.Item(2) + QtdDocumentosNivel2PorNivelAnalista.Item(3))
QtdAnalistasNivelAtual = QtdAnalistaNivel4
QtdDocumentosNivel2PorNivelAnalista.Add 4, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 1) * QtdAnalistaNivel4

'Distribuição Nível 2. Possibilidade de distribuição universal.
Limite = QtdDocumentosNivel2PorNivelAnalista.Item(1)
CalculoDistribIgualitaria DocumentosNivel2, AnalistaNivel1, 0, QtdDocumentosNivel2PorNivelAnalista.Item(1), PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite
Limite = QtdDocumentosNivel2PorNivelAnalista.Item(1) + QtdDocumentosNivel2PorNivelAnalista.Item(2)
CalculoDistribIgualitaria DocumentosNivel2, AnalistaNivel2, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite
Limite = Limite + QtdDocumentosNivel2PorNivelAnalista.Item(3)
CalculoDistribIgualitaria DocumentosNivel2, AnalistaNivel3, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite
Limite = Limite + QtdDocumentosNivel2PorNivelAnalista.Item(4)
CalculoDistribIgualitaria DocumentosNivel2, AnalistaNivel4, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas

'Preparação de Distribuição Nível 3. Preparação das respectivas quotas
QtdAnalistasNivelAtual = QtdAnalistaNivel2
QtdDocumentosNivelAtual = DocumentosNivel3.Count
QtdDocumentosNivel3PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.3) * QtdAnalistaNivel2
QtdAnalistasNivelAtual = QtdAnalistaNivel3
QtdDocumentosNivel3PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.35) * QtdAnalistaNivel3
QtdDocumentosAuxiliar = QtdDocumentosNivel3PorNivelAnalista.Item(1) + QtdDocumentosNivel3PorNivelAnalista.Item(2)
QtdDocumentosCompensados = QtdDocumentosNivelAtual - QtdDocumentosAuxiliar
QtdAnalistasNivelAtual = QtdAnalistaNivel4
QtdDocumentosNivel3PorNivelAnalista.Add 3, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosCompensados, 1) * QtdAnalistaNivel4
'Distribuição Nível 3. Possibilidade de distribuição universal.
Limite = QtdDocumentosNivel3PorNivelAnalista.Item(1)
CalculoDistribIgualitaria DocumentosNivel3, AnalistaNivel2, 0, Limite, PlanilhaRav, RAVColunas
Limite = QtdDocumentosNivel3PorNivelAnalista.Item(1) + QtdDocumentosNivel3PorNivelAnalista.Item(2)
CalculoDistribIgualitaria DocumentosNivel3, AnalistaNivel3, QtdDocumentosNivel3PorNivelAnalista.Item(1), Limite, PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite
Limite = Limite + QtdDocumentosNivel3PorNivelAnalista.Item(3)
CalculoDistribIgualitaria DocumentosNivel3, AnalistaNivel3, LimiteAuxiliar, Limite, PlanilhaRav, RAVColunas
LimiteAuxiliar = Limite

'Preparação de Distribuição Nível 4. Preparação das respectivas quotas
QtdAnalistasNivelAtual = QtdAnalistaNivel3
QtdDocumentosNivelAtual = DocumentosNivel4.Count
QtdDocumentosNivel4PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.6) * QtdAnalistaNivel3
QtdAnalistasNivelAtual = QtdAnalistaNivel4
QtdDocumentosAuxiliar = QtdDocumentosNivelAtual - QtdDocumentosNivel4PorNivelAnalista.Item(1)
QtdDocumentosNivel4PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosAuxiliar, 1) * QtdAnalistaNivel4
'Distribuição Nível 4. Possibilidade de distribuição universal.
Limite = QtdDocumentosNivel4PorNivelAnalista.Item(1) * QtdAnalistaNivel4
CalculoDistribIgualitaria DocumentosNivel4, AnalistaNivel3, 0, Limite, PlanilhaRav, RAVColunas
Limite = QtdDocumentosNivel4PorNivelAnalista.Item(1) + QtdDocumentosNivel4PorNivelAnalista.Item(2)
CalculoDistribIgualitaria DocumentosNivel4, AnalistaNivel4, QtdDocumentosNivel4PorNivelAnalista.Item(1), Limite, PlanilhaRav, RAVColunas

'Distribuição Nível 5. Possibilidade de distribuição universal.
CalculoDistribIgualitaria DocumentosNivel5, AnalistaNivel4, 0, CalculoCotaDistrib(QtdAnalistaNivel4, DocumentosNivel5.Count, 1) * QtdAnalistaNivel4, PlanilhaRav, RAVColunas

Else
'Distribuição integral de documentos nível 1
QtdAnalistasNivelAtual = QtdAnalistaNivel1
QtdDocumentosNivelAtual = DocumentosNivel1.Count
QtdDocumentosNivel1PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 1) * QtdAnalistaNivel1
CalculoDistribIgualitaria DocumentosNivel1, AnalistaNivel1, 0, QtdDocumentosNivelAtual, PlanilhaRav, RAVColunas
'Distribuição integral de documentos nível 2
QtdDocumentosNivelAtual = DocumentosNivel2.Count
QtdDocumentosNivel2PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.4) * QtdAnalistaNivel1
CalculoDistribIgualitaria DocumentosNivel2, AnalistaNivel1, 0, QtdDocumentosNivel2PorNivelAnalista.Item(1), PlanilhaRav, RAVColunas
QtdDocumentosNivelAtual = DocumentosNivel2.Count - QtdDocumentosNivel2PorNivelAnalista.Item(1)
QtdAnalistasNivelAtual = QtdAnalistaNivel2
QtdDocumentosNivel2PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 1) * QtdAnalistaNivel2
CalculoDistribIgualitaria DocumentosNivel2, AnalistaNivel2, QtdDocumentosNivel2PorNivelAnalista.Item(1), DocumentosNivel2.Count, PlanilhaRav, RAVColunas
'Distribuição Nível 3. Possibilidade de distribuição universal.
QtdDocumentosNivelAtual = DocumentosNivel3.Count
QtdDocumentosNivel3PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.45) * QtdAnalistaNivel2
CalculoDistribIgualitaria DocumentosNivel3, AnalistaNivel2, 0, QtdDocumentosNivel3PorNivelAnalista.Item(1), PlanilhaRav, RAVColunas
QtdDocumentosNivelAtual = DocumentosNivel3.Count - QtdDocumentosNivel3PorNivelAnalista.Item(1)
QtdAnalistasNivelAtual = QtdAnalistaNivel3
QtdDocumentosNivel3PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 1) * QtdAnalistaNivel3
CalculoDistribIgualitaria DocumentosNivel3, AnalistaNivel3, QtdDocumentosNivel3PorNivelAnalista.Item(1), DocumentosNivel3.Count, PlanilhaRav, RAVColunas
'Distribuição Nível 4. Possibilidade de distribuição universal.
QtdDocumentosNivelAtual = DocumentosNivel4.Count
QtdDocumentosNivel4PorNivelAnalista.Add 1, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 0.5) * QtdAnalistaNivel3
CalculoDistribIgualitaria DocumentosNivel4, AnalistaNivel3, 0, QtdDocumentosNivel4PorNivelAnalista.Item(1), PlanilhaRav, RAVColunas
QtdDocumentosNivelAtual = DocumentosNivel4.Count - QtdDocumentosNivel4PorNivelAnalista.Item(1)
QtdAnalistasNivelAtual = QtdAnalistaNivel4
QtdDocumentosNivel4PorNivelAnalista.Add 2, CalculoCotaDistrib(QtdAnalistasNivelAtual, QtdDocumentosNivelAtual, 1) * QtdAnalistaNivel4
CalculoDistribIgualitaria DocumentosNivel4, AnalistaNivel4, QtdDocumentosNivel4PorNivelAnalista.Item(1), DocumentosNivel4.Count, PlanilhaRav, RAVColunas
'Distribuição Nível 5. Possibilidade de distribuição universal.
CalculoDistribIgualitaria DocumentosNivel5, AnalistaNivel4, 0, CalculoCotaDistrib(QtdAnalistaNivel4, DocumentosNivel5.Count, 1) * QtdAnalistaNivel4, PlanilhaRav, RAVColunas
End If




End Function

Public Function CalculoCotaDistrib(QtdAnalistas As Variant, QtdDocumentos As Variant, Percentual As Double) As Integer
Dim resto As Double
Dim QtdDocumentosReservados, QtdDocumentosArrendondos As Integer

QtdDocumentosArrendondos = Application.WorksheetFunction.Round(QtdDocumentos * Percentual, 0)

'QtdDocumentos = Int(QtdDocumentos - QtdDocumentosArrendondos)

resto = QtdDocumentosArrendondos Mod QtdAnalistas

If resto > 0 Then
QtdDocumentosReservados = QtdDocumentosArrendondos - resto
QtdDocumentosReservados = QtdDocumentosReservados / QtdAnalistas
Else
QtdDocumentosReservados = QtdDocumentosArrendondos / QtdAnalistas
resto = 0
End If

CalculoCotaDistrib = QtdDocumentosReservados
End Function

Public Function CalculoDistribIgualitaria(DocumentosLinhas As Dictionary, Analistas As Variant, _
Indice As Variant, Limite As Variant, PlanilhaRav As Excel.Worksheet, RAVColunas As cRAVColunasXL)
Dim IndiceAtual, QuotaAnalista, IndiceQuota, Auxiliar As Long
Dim IndiceAnalista As Integer
Dim EIncioDistrib As Boolean
Dim QtdAnalista As Integer: QtdAnalista = UBound(Analistas) + 1
QuotaAnalista = Limite - Indice
QuotaAnalista = QuotaAnalista / QtdAnalista
Limite = Limite - 1
Auxiliar = 1
EIncioDistrib = True
For IndiceAtual = Indice To Limite
PlanilhaRav.Range(RAVColunas.Analista & DocumentosLinhas.Keys(IndiceAtual)).Value2 = Analistas(IndiceAnalista)
IndiceQuota = IndiceQuota + 1
Auxiliar = Auxiliar + 1
If Auxiliar = QuotaAnalista + 1 Then
IndiceAnalista = IndiceAnalista + 1
Auxiliar = 1
End If
Next IndiceAtual
Limite = Limite + 1
End Function
Public Function QuotaDocumentos(QtdDocumentos As Integer, QtdAnalistas As Integer, QtdAnalistasPrioritarios As Integer, _
Percentual As Double, Percentual2 As Double) As Dictionary

End Function


