Attribute VB_Name = "modPreencherFrequencia"
Option Explicit
' ==============================================================================
' MÓDULO: modPreencherFrequencia
' DESCRIÇÃO: Preenche FREQUENCIA com "MENSAL" em linhas onde está vazia
'            mas que tenham dados em outras colunas.
'
' REGRA:
'   FREQUENCIA vazia + linha com dados -> FREQUENCIA = "MENSAL"
'
' TABELA: tbl_Lancamentos_Fixos
' ==============================================================================
Sub Preencher_Frequencia_Padrao()

    ' =========================
    ' CONFIGURAÇÕES
    ' =========================
    Const COL_FREQUENCIA  As String = "FREQUENCIA"
    Const FREQ_PADRAO     As String = "MENSAL"

    Dim tabelasProcessar As Variant
    tabelasProcessar = Array("tbl_Lancamentos_Fixos")

    ' =========================
    ' VARIÁVEIS
    ' =========================
    Dim tbl          As ListObject
    Dim ws           As Worksheet
    Dim nomeTabela   As Variant
    Dim colFreq      As ListColumn
    Dim i            As Long
    Dim j            As Long
    Dim linhaTemDados As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' =========================
    ' PROCESSAR CADA TABELA
    ' =========================
    For Each nomeTabela In tabelasProcessar

        Set tbl = Nothing
        Set colFreq = Nothing

        On Error Resume Next
        For Each ws In ThisWorkbook.Worksheets
            Set tbl = ws.ListObjects(CStr(nomeTabela))
            If Not tbl Is Nothing Then Exit For
        Next ws
        On Error GoTo 0

        If tbl Is Nothing Then GoTo ProximaTabela
        If tbl.DataBodyRange Is Nothing Then GoTo ProximaTabela

        On Error Resume Next
        Set colFreq = tbl.ListColumns(COL_FREQUENCIA)
        On Error GoTo 0

        If colFreq Is Nothing Then GoTo ProximaTabela

        For i = 1 To tbl.ListRows.Count

            ' Só age se FREQUENCIA estiver vazia
            If Trim(CStr(tbl.DataBodyRange.Cells(i, colFreq.Index).Value)) = "" Then

                ' Verifica se a linha tem dados em alguma outra coluna
                linhaTemDados = False
                For j = 1 To tbl.ListColumns.Count
                    If j <> colFreq.Index Then
                        If Trim(CStr(tbl.DataBodyRange.Cells(i, j).Value)) <> "" Then
                            linhaTemDados = True
                            Exit For
                        End If
                    End If
                Next j

                If linhaTemDados Then
                    tbl.DataBodyRange.Cells(i, colFreq.Index).Value = FREQ_PADRAO
                End If

            End If

        Next i

ProximaTabela:
    Next nomeTabela

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
