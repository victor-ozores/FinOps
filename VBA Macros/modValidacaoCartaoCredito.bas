Attribute VB_Name = "modValidacaoCartaoCredito"
Option Explicit

' ==============================================================================
' VALIDAÇÃO DE CARTÃO DE CRÉDITO - LANÇAMENTOS FIXOS E EVENTUAIS
'
' Limpa APENAS CARTAO quando METODO_PAGAMENTO <> "CARTAO CREDITO"
' Previne inconsistências antes da atualização no Power BI
' Executa no Workbook_Open para garantir dados limpos desde o início da sessão
' ==============================================================================

Public Sub Validar_Cartao_Credito_Lancamentos()

    ' =========================
    ' CONFIGURAÇÕES
    ' =========================
    Const COL_METODO    As String = "METODO_PAGAMENTO"
    Const COL_CARTAO    As String = "CARTAO"
    Const METODO_CARTAO As String = "CARTAO CREDITO"

    Dim totalLimpezas As Long
    totalLimpezas = 0

    ' =========================
    ' OTIMIZAÇÃO
    ' =========================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' =========================
    ' PROCESSAR FIXOS
    ' =========================
    totalLimpezas = totalLimpezas + ValidarTabela("tbl_Lancamentos_Fixos", COL_METODO, COL_CARTAO, METODO_CARTAO)

    ' =========================
    ' PROCESSAR EVENTUAIS
    ' =========================
    totalLimpezas = totalLimpezas + ValidarTabela("tbl_Lancamentos_Eventuais", COL_METODO, COL_CARTAO, METODO_CARTAO)

    ' =========================
    ' RESTAURAR CONFIGURAÇÕES
    ' =========================
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

' ==============================================================================
' VALIDAR TABELA INDIVIDUAL
' ==============================================================================
Private Function ValidarTabela( _
    nomeTbl As String, _
    colMetodo As String, _
    colCartao As String, _
    metodoValido As String _
) As Long

    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim colMet As ListColumn
    Dim colCA As ListColumn
    Dim i As Long
    Dim limpezas As Long
    Dim metodo As String
    Dim valorMetodo As Variant
    Dim valorCartao As Variant
    Dim temCartao As Boolean

    limpezas = 0

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = ws.ListObjects(nomeTbl)
        If Not tbl Is Nothing Then Exit For
    Next ws
    On Error GoTo 0

    If tbl Is Nothing Then ValidarTabela = 0: Exit Function
    If tbl.DataBodyRange Is Nothing Then ValidarTabela = 0: Exit Function

    On Error Resume Next
    Set colMet = tbl.ListColumns(colMetodo)
    Set colCA = tbl.ListColumns(colCartao)
    On Error GoTo 0

    If colMet Is Nothing Or colCA Is Nothing Then ValidarTabela = 0: Exit Function

    For i = 1 To tbl.ListRows.Count

        valorMetodo = tbl.DataBodyRange.Cells(i, colMet.Index).Value
        metodo = Trim$(UCase$(CStr(valorMetodo)))

        temCartao = False
        valorCartao = tbl.DataBodyRange.Cells(i, colCA.Index).Value
        If Not IsEmpty(valorCartao) Then
            If Trim$(CStr(valorCartao)) <> "" Then temCartao = True
        End If

        If metodo <> "" And metodo <> metodoValido And temCartao Then
            tbl.DataBodyRange.Cells(i, colCA.Index).ClearContents
            limpezas = limpezas + 1
        End If

    Next i

    ValidarTabela = limpezas

End Function

