Attribute VB_Name = "modStatusPendente"
Option Explicit

' ==============================================================================
' MACRO: PREENCHER STATUS AUTOMATICAMENTE
' DESCRIÇÃO: Preenche STATUS em linhas onde está vazio MAS que tenham dados.
'
' REGRA:
'   METODO_PAGAMENTO = "CARTAO CREDITO" -> STATUS = "EMITIDO"
'   Qualquer outro método               -> STATUS = "PENDENTE"
'
' TABELAS: tbl_Lancamentos_Fixos, tbl_Lancamentos_Eventuais
' ==============================================================================
Sub Preencher_Status_Pendente()

    ' =========================
    ' CONFIGURAÇÕES
    ' =========================
    Const COL_STATUS As String = "STATUS"
    Const COL_METODO As String = "METODO_PAGAMENTO"
    Const METODO_CARTAO   As String = "CARTAO CREDITO"
    Const STATUS_EMITIDO  As String = "EMITIDO"
    Const STATUS_PENDENTE As String = "PENDENTE"

    Dim tabelasProcessar As Variant
    tabelasProcessar = Array("tbl_Lancamentos_Fixos", "tbl_Lancamentos_Eventuais")

    ' =========================
    ' VARIÁVEIS
    ' =========================
    Dim tbl          As ListObject
    Dim ws           As Worksheet
    Dim nomeTabela   As Variant
    Dim colStatus    As ListColumn
    Dim colMetodo    As ListColumn
    Dim i            As Long
    Dim j            As Long
    Dim linhaTemDados As Boolean
    Dim metodoVal    As String
    Dim statusValor  As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' =========================
    ' PROCESSAR CADA TABELA
    ' =========================
    For Each nomeTabela In tabelasProcessar

        Set tbl = Nothing
        Set colStatus = Nothing
        Set colMetodo = Nothing

        On Error Resume Next
        For Each ws In ThisWorkbook.Worksheets
            Set tbl = ws.ListObjects(CStr(nomeTabela))
            If Not tbl Is Nothing Then Exit For
        Next ws
        On Error GoTo 0

        If tbl Is Nothing Then GoTo ProximaTabela
        If tbl.DataBodyRange Is Nothing Then GoTo ProximaTabela

        On Error Resume Next
        Set colStatus = tbl.ListColumns(COL_STATUS)
        On Error GoTo 0
        If colStatus Is Nothing Then GoTo ProximaTabela

        On Error Resume Next
        Set colMetodo = tbl.ListColumns(COL_METODO)
        On Error GoTo 0

        For i = 1 To tbl.ListRows.Count

            If Trim(CStr(tbl.DataBodyRange.Cells(i, colStatus.Index).Value)) = "" Then

                linhaTemDados = False
                For j = 1 To tbl.ListColumns.Count
                    If j <> colStatus.Index Then
                        If Trim(CStr(tbl.DataBodyRange.Cells(i, j).Value)) <> "" Then
                            linhaTemDados = True
                            Exit For
                        End If
                    End If
                Next j

                If linhaTemDados Then

                    If Not colMetodo Is Nothing Then
                        metodoVal = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, colMetodo.Index).Value)))
                        If metodoVal = METODO_CARTAO Then
                            statusValor = STATUS_EMITIDO
                        Else
                            statusValor = STATUS_PENDENTE
                        End If
                    Else
                        statusValor = STATUS_PENDENTE
                    End If

                    tbl.DataBodyRange.Cells(i, colStatus.Index).Value = statusValor

                End If

            End If

        Next i

ProximaTabela:
    Next nomeTabela

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
