Attribute VB_Name = "modParcelamento"
Option Explicit

' ==============================================================================
' GERAR PARCELAMENTO - AUTOMÁTICO (SILENCIOSO) + PADRONIZAÇÃO FORÇADA DE IDs
'
' OBJETIVO
' - Gerar parcelas automaticamente quando:
'   - N_PARCELAS > 1
'   - ID_PARCELAMENTO vazio (ainda não gerado)
'   - DATA_EFETIVA válida
'
' PADRÃO OFICIAL (ÚNICO PADRÃO)
' - ID_PARCELAMENTO = GUID 32 HEX (sem chaves e sem hífens)
'   Ex: 2F66B702460B44F18B5C7D0A1B2C3D4E
'
' REGRAS DE STATUS (ATUALIZADO)
' - A linha original (1ª parcela) mantém o STATUS que o usuário definiu.
' - Parcelas futuras (2..N):
'     CARTAO CREDITO  -> STATUS = "EMITIDO"  (já estão na fatura futura)
'     Outros métodos  -> STATUS = "PENDENTE" (aguardando pagamento)
'
' SEGURANÇA / PADRONIZAÇÃO
' - Antes de gerar, varre todas as linhas e força que:
'   - Se ID_PARCELAMENTO não for GUID, ele é migrado para GUID
'   - Mantém agrupamento: mesmo ID antigo -> mesmo GUID em todas as linhas
'   - Se o ID antigo contém um GUID (ex: "PARC_..._<GUID>"), ele extrai e usa
'   - Se for ID curto (ex: "PARC_20260115_182311"), cria GUID novo e aplica em lote
'
' CORREÇÃO: Fórmulas e validações de dados agora são preservadas nas parcelas geradas.
' ==============================================================================

Public Sub Gerar_Parcelamento_Automatico()

    ' =========================
    ' CONFIGURAÇÕES (TABELA / PLANILHA)
    ' =========================
    Const TBL_EVENTUAIS  As String = "tbl_Lancamentos_Eventuais"
    Const PLAN_EVENTUAIS As String = "Lancamentos_Eventuais"

    ' =========================
    ' CONFIGURAÇÕES (COLUNAS)
    ' =========================
    Const COL_N_PARCELAS      As String = "N_PARCELAS"
    Const COL_PARCELA_ATUAL   As String = "PARCELA_ATUAL"
    Const COL_PARCELAMENTO_ID As String = "ID_PARCELAMENTO"
    Const COL_DATA_EFETIVA    As String = "DATA_EFETIVA"
    Const COL_STATUS          As String = "STATUS"
    Const COL_METODO          As String = "METODO_PAGAMENTO"

    ' =========================
    ' VARIÁVEIS
    ' =========================
    Dim tbl          As ListObject
    Dim ws           As Worksheet
    Dim i            As Long
    Dim nParcelas    As Long
    Dim parcelamentoID As String
    Dim dataOriginal As Variant
    Dim initialCount As Long
    Dim v            As Variant
    Dim metodoVal    As String

    Dim mapOldToGuid As Object
    Set mapOldToGuid = CreateObject("Scripting.Dictionary")
    mapOldToGuid.CompareMode = vbTextCompare

    ' =========================
    ' LOCALIZAR TABELA
    ' =========================
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PLAN_EVENTUAIS)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(TBL_EVENTUAIS)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' =========================
    ' PERFORMANCE / SEGURANÇA
    ' =========================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp

    initialCount = tbl.ListRows.Count

    ' ==============================================================================
    ' 1) PADRONIZAR TODOS OS ID_PARCELAMENTO EXISTENTES (FORÇAR GUID)
    ' ==============================================================================
    For i = 1 To initialCount

        parcelamentoID = Trim$(CStr(tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_PARCELAMENTO_ID).Index).Value))
        If parcelamentoID = vbNullString Then GoTo NextNorm

        If IsGuid32(parcelamentoID) Then GoTo NextNorm

        If mapOldToGuid.Exists(parcelamentoID) Then
            tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_PARCELAMENTO_ID).Index).Value = mapOldToGuid(parcelamentoID)
            GoTo NextNorm
        End If

        Dim extracted As String
        extracted = ExtractGuid32FromAny(parcelamentoID)

        If extracted <> vbNullString Then
            mapOldToGuid.Add parcelamentoID, extracted
            tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_PARCELAMENTO_ID).Index).Value = extracted
        Else
            Dim g As String
            g = CreateGuidNoBraces()
            mapOldToGuid.Add parcelamentoID, g
            tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_PARCELAMENTO_ID).Index).Value = g
        End If

NextNorm:
    Next i

    ' ==============================================================================
    ' 2) GERAR NOVOS PARCELAMENTOS (APENAS QUANDO ID_PARCELAMENTO ESTIVER VAZIO)
    ' ==============================================================================
    For i = 1 To initialCount

        nParcelas = 0
        parcelamentoID = vbNullString
        dataOriginal = Empty
        metodoVal = vbNullString

        v = tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_N_PARCELAS).Index).Value
        If IsNumeric(v) Then nParcelas = CLng(v)

        parcelamentoID = Trim$(CStr(tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_PARCELAMENTO_ID).Index).Value))
        dataOriginal = tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_DATA_EFETIVA).Index).Value
        metodoVal = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, tbl.ListColumns(COL_METODO).Index).Value)))

        If nParcelas > 1 And parcelamentoID = "" And IsDate(dataOriginal) Then
            ProcessarParcelamento tbl, i, nParcelas, CDate(dataOriginal), metodoVal, _
                                  COL_N_PARCELAS, COL_PARCELA_ATUAL, COL_PARCELAMENTO_ID, _
                                  COL_DATA_EFETIVA, COL_STATUS
        End If

    Next i

CleanUp:
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

' ==============================================================================
' PROCESSAR PARCELAMENTO (GERAÇÃO DAS LINHAS)
' ==============================================================================
Private Sub ProcessarParcelamento( _
    tbl As ListObject, _
    idxLinha As Long, _
    nParcelas As Long, _
    dataOriginal As Date, _
    metodoOriginal As String, _
    colNParcelas As String, _
    colParcelaAtual As String, _
    colParcelamentoID As String, _
    colDataEfetiva As String, _
    colStatus As String)

    Const METODO_CARTAO  As String = "CARTAO CREDITO"
    Const STATUS_EMITIDO As String = "EMITIDO"
    Const STATUS_PENDENTE As String = "PENDENTE"

    Dim parcelamentoID  As String
    Dim i               As Long
    Dim novaLinha       As ListRow
    Dim colIdx          As Long
    Dim celParcelaAtual As Range
    Dim srcCell         As Range
    Dim dstCell         As Range

    Dim statusParcelas As String
    If metodoOriginal = METODO_CARTAO Then
        statusParcelas = STATUS_EMITIDO
    Else
        statusParcelas = STATUS_PENDENTE
    End If

    parcelamentoID = CreateGuidNoBraces()

    tbl.DataBodyRange.Cells(idxLinha, tbl.ListColumns(colParcelamentoID).Index).Value = parcelamentoID

    Set celParcelaAtual = tbl.DataBodyRange.Cells(idxLinha, tbl.ListColumns(colParcelaAtual).Index)
    celParcelaAtual.NumberFormat = "@"
    celParcelaAtual.Value = "1/" & nParcelas

    For i = 2 To nParcelas

        Set novaLinha = tbl.ListRows.Add

        For colIdx = 1 To tbl.ListColumns.Count
            Set srcCell = tbl.DataBodyRange.Cells(idxLinha, colIdx)
            Set dstCell = novaLinha.Range.Cells(1, colIdx)

            If srcCell.HasFormula Then
                dstCell.Formula = srcCell.Formula
            Else
                dstCell.Value = srcCell.Value
            End If
        Next colIdx

        tbl.DataBodyRange.Rows(idxLinha).Copy
        novaLinha.Range.PasteSpecial Paste:=xlPasteValidation
        Application.CutCopyMode = False

        novaLinha.Range.Cells(1, tbl.ListColumns(colDataEfetiva).Index).Value = DateAdd("m", i - 1, dataOriginal)

        Set celParcelaAtual = novaLinha.Range.Cells(1, tbl.ListColumns(colParcelaAtual).Index)
        celParcelaAtual.NumberFormat = "@"
        celParcelaAtual.Value = i & "/" & nParcelas

        novaLinha.Range.Cells(1, tbl.ListColumns(colParcelamentoID).Index).Value = parcelamentoID
        novaLinha.Range.Cells(1, tbl.ListColumns(colStatus).Index).Value = statusParcelas

    Next i

End Sub

' ==============================================================================
' HELPERS: GUID / VALIDAÇÃO / EXTRAÇÃO
' ==============================================================================

Private Function IsGuid32(ByVal s As String) As Boolean
    Dim t As String
    t = KeepHexOnly(UCase$(Trim$(s)))
    IsGuid32 = (Len(t) = 32)
End Function

Private Function ExtractGuid32FromAny(ByVal s As String) As String
    Dim t As String
    t = KeepHexOnly(UCase$(Trim$(s)))

    If Len(t) = 32 Then
        ExtractGuid32FromAny = t
    ElseIf Len(t) > 32 Then
        ExtractGuid32FromAny = Right$(t, 32)
    Else
        ExtractGuid32FromAny = vbNullString
    End If
End Function

Private Function KeepHexOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "F") Then
            out = out & ch
        End If
    Next i
    KeepHexOnly = out
End Function

Private Function CreateGuidNoBraces() As String
    Dim g As String
    g = CreateObject("Scriptlet.TypeLib").GUID
    g = Replace$(g, "{", "")
    g = Replace$(g, "}", "")
    g = Replace$(g, "-", "")
    CreateGuidNoBraces = UCase$(g)
End Function

