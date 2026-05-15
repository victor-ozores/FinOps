Attribute VB_Name = "modStatusAtrasado"
Option Explicit

' ==============================================================================
' MÓDULO: modStatusAtrasado
' DESCRIÇÃO: Atualiza STATUS para "ATRASADO" quando o prazo venceu.
'
' REGRAS:
'   PENDENTE + não cartão  -> ATRASADO se DATA_EFETIVA < Hoje
'   EMITIDO  + cartão      -> ATRASADO se vencimento da fatura < Hoje
'   PENDENTE + cartão      -> não toca (fatura ainda não fechou)
'   PAGO / ATRASADO        -> nunca toca
'
' LÓGICA DO CICLO DO CARTÃO:
'   Se Day(DATA_EFETIVA) <= DIA_FECH -> compra entrou na fatura deste ciclo
'     Se DIA_VENC >= DIA_FECH -> vence no mesmo mês da DATA_EFETIVA
'     Se DIA_VENC <  DIA_FECH -> vence no mês seguinte
'   Se Day(DATA_EFETIVA) >  DIA_FECH -> compra entrou no próximo ciclo
'     Se DIA_VENC >= DIA_FECH -> vence no mês seguinte à DATA_EFETIVA
'     Se DIA_VENC <  DIA_FECH -> vence em dois meses à frente
' ==============================================================================

Public Sub Atualizar_Status_Atrasado()

    ' =========================
    ' CONSTANTES
    ' =========================
    Const STATUS_EMITIDO    As String = "EMITIDO"
    Const STATUS_PENDENTE   As String = "PENDENTE"
    Const STATUS_ATRASADO   As String = "ATRASADO"
    Const METODO_CARTAO     As String = "CARTAO CREDITO"

    Const COL_STATUS        As String = "STATUS"
    Const COL_METODO        As String = "METODO_PAGAMENTO"
    Const COL_CARTAO        As String = "CARTAO"
    Const COL_DATA          As String = "DATA_EFETIVA"

    Const TBL_CARTOES       As String = "tbl_Base_Cartoes_Credito"
    Const CART_COL_DESC     As String = "NOME"
    Const CART_COL_FECH     As String = "DIA_FECH"
    Const CART_COL_VENC     As String = "DIA_VENC"

    Dim tabelasAlvo(1)      As String
    tabelasAlvo(0) = "tbl_Lancamentos_Fixos"
    tabelasAlvo(1) = "tbl_Lancamentos_Eventuais"

    ' =========================
    ' PERFORMANCE
    ' =========================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' =========================
    ' CARREGAR CARTÕES EM DICIONÁRIO
    ' chave = NOME (texto), valor = Array(DIA_FECH, DIA_VENC)
    ' =========================
    Dim dictCartoes As Object
    Set dictCartoes = CarregarCartoes(TBL_CARTOES, CART_COL_DESC, CART_COL_FECH, CART_COL_VENC)

    Dim hoje As Date
    hoje = Date

    ' =========================
    ' PROCESSAR CADA TABELA
    ' =========================
    Dim nomeTbl As Variant
    For Each nomeTbl In tabelasAlvo
        ProcessarTabela _
            CStr(nomeTbl), _
            COL_STATUS, COL_METODO, COL_CARTAO, COL_DATA, _
            STATUS_PENDENTE, STATUS_EMITIDO, STATUS_ATRASADO, METODO_CARTAO, _
            dictCartoes, hoje
    Next nomeTbl

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

' ==============================================================================
' PROCESSAR TABELA
' ==============================================================================
Private Sub ProcessarTabela( _
    nomeTbl As String, _
    colStatus As String, colMetodo As String, colCartao As String, colData As String, _
    STATUS_PENDENTE As String, STATUS_EMITIDO As String, STATUS_ATRASADO As String, _
    METODO_CARTAO As String, _
    dictCartoes As Object, _
    hoje As Date _
)
    Dim tbl As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = ws.ListObjects(nomeTbl)
        If Not tbl Is Nothing Then Exit For
    Next ws
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim colST As ListColumn, colMT As ListColumn
    Dim colCA As ListColumn, colDT As ListColumn

    On Error Resume Next
    Set colST = tbl.ListColumns(colStatus)
    Set colMT = tbl.ListColumns(colMetodo)
    Set colCA = tbl.ListColumns(colCartao)
    Set colDT = tbl.ListColumns(colData)
    On Error GoTo 0

    If colST Is Nothing Or colMT Is Nothing Or colDT Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To tbl.ListRows.Count

        Dim statusVal As String
        statusVal = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, colST.Index).Value)))

        If statusVal <> STATUS_PENDENTE And statusVal <> STATUS_EMITIDO Then GoTo ProximaLinha

        Dim metodoVal As String
        metodoVal = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, colMT.Index).Value)))

        Dim dataVal As Variant
        dataVal = tbl.DataBodyRange.Cells(i, colDT.Index).Value

        If Not IsDate(dataVal) Then GoTo ProximaLinha

        Dim dataEfetiva As Date
        dataEfetiva = CDate(dataVal)

        ' -------------------------------------------------------
        ' CASO 1: Não é cartão de crédito -> usa DATA_EFETIVA
        ' -------------------------------------------------------
        If metodoVal <> UCase$(METODO_CARTAO) Then
            If statusVal = STATUS_PENDENTE Then
                If dataEfetiva < hoje Then
                    tbl.DataBodyRange.Cells(i, colST.Index).Value = STATUS_ATRASADO
                End If
            End If
            GoTo ProximaLinha
        End If

        ' -------------------------------------------------------
        ' CASO 2: Cartão de crédito
        ' -------------------------------------------------------
        If statusVal = STATUS_PENDENTE Then GoTo ProximaLinha

        If statusVal = STATUS_EMITIDO Then

            If colCA Is Nothing Then GoTo ProximaLinha

            Dim descCartao As String
            descCartao = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, colCA.Index).Value)))

            If descCartao = "" Then GoTo ProximaLinha

            If Not dictCartoes.Exists(descCartao) Then GoTo ProximaLinha

            Dim infoCartao As Variant
            infoCartao = dictCartoes(descCartao)

            Dim diaFech As Long
            Dim diaVenc As Long
            diaFech = CLng(infoCartao(0))
            diaVenc = CLng(infoCartao(1))

            Dim dataVenc As Date
            dataVenc = CalcularVencimentoFatura(dataEfetiva, diaFech, diaVenc)

            If dataVenc < hoje Then
                tbl.DataBodyRange.Cells(i, colST.Index).Value = STATUS_ATRASADO
            End If

        End If

ProximaLinha:
    Next i

End Sub

' ==============================================================================
' CALCULAR VENCIMENTO DA FATURA
' ==============================================================================
Private Function CalcularVencimentoFatura( _
    dataEfetiva As Date, _
    diaFech As Long, _
    diaVenc As Long _
) As Date

    Dim diaCompra As Long
    diaCompra = Day(dataEfetiva)

    Dim offsetMeses As Long

    If diaCompra < diaFech Then
        If diaVenc >= diaFech Then
            offsetMeses = 0
        Else
            offsetMeses = 1
        End If
    Else
        If diaVenc >= diaFech Then
            offsetMeses = 1
        Else
            offsetMeses = 2
        End If
    End If

    CalcularVencimentoFatura = DateSerial( _
        Year(dataEfetiva), _
        Month(dataEfetiva) + offsetMeses, _
        diaVenc _
    )

End Function

' ==============================================================================
' CARREGAR CARTÕES EM DICIONÁRIO
' chave = NOME (UCase), valor = Array(DIA_FECH, DIA_VENC)
' ==============================================================================
Private Function CarregarCartoes( _
    nomeTbl As String, _
    colDesc As String, colFech As String, colVenc As String _
) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim tbl As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = ws.ListObjects(nomeTbl)
        If Not tbl Is Nothing Then Exit For
    Next ws
    On Error GoTo 0

    If tbl Is Nothing Then Set CarregarCartoes = dict: Exit Function
    If tbl.DataBodyRange Is Nothing Then Set CarregarCartoes = dict: Exit Function

    Dim colD As ListColumn, colF As ListColumn, colV As ListColumn

    On Error Resume Next
    Set colD = tbl.ListColumns(colDesc)
    Set colF = tbl.ListColumns(colFech)
    Set colV = tbl.ListColumns(colVenc)
    On Error GoTo 0

    If colD Is Nothing Or colF Is Nothing Or colV Is Nothing Then
        Set CarregarCartoes = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To tbl.ListRows.Count

        Dim vDesc As String
        Dim vFech As Variant
        Dim vVenc As Variant

        vDesc = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, colD.Index).Value)))
        vFech = tbl.DataBodyRange.Cells(i, colF.Index).Value
        vVenc = tbl.DataBodyRange.Cells(i, colV.Index).Value

        If vDesc <> "" And IsNumeric(vFech) And IsNumeric(vVenc) Then
            If Not dict.Exists(vDesc) Then
                dict.Add vDesc, Array(CLng(vFech), CLng(vVenc))
            End If
        End If

    Next i

    Set CarregarCartoes = dict

End Function

