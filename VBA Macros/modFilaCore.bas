Attribute VB_Name = "modFilaCore"
Option Explicit

' ==============================================================================
' MÓDULO: modProjecaoFila
' DESCRIÇÃO: Projeção de lançamentos fixos por frequência
'
' FREQUÊNCIAS SUPORTADAS:
'   SEMANAL    -> +7 dias   (compara dia exato)
'   QUINZENAL  -> +15 dias  (compara dia exato)
'   MENSAL     -> +1 mês    (compara mês/ano — projeta no 1º dia do mês de competência)
'   BIMESTRAL  -> +2 meses  (compara mês/ano)
'   TRIMESTRAL -> +3 meses  (compara mês/ano)
'   SEMESTRAL  -> +6 meses  (compara mês/ano)
'   ANUAL      -> +12 meses (compara mês/ano)
'
' CHAVE DE IDENTIFICAÇÃO: RESPONSAVEL|SERVICO|VALOR|TIPO
' PADRÃO SE SEM FREQUÊNCIA: MENSAL
' ==============================================================================

Public Sub ProjetarLancamentosFixos_Fila()

    ' =========================
    ' CONSTANTES
    ' =========================
    Const TBL_DADOS         As String = "tbl_Lancamentos_Fixos"
    Const TBL_PROJECAO      As String = "tbl_Projecao_Lancamentos_Fixos"

    Const COL_DATA          As String = "DATA_EFETIVA"
    Const COL_STATUS        As String = "STATUS"
    Const COL_FREQUENCIA    As String = "FREQUENCIA"
    Const COL_RESPONSAVEL   As String = "RESPONSAVEL"
    Const COL_SERVICO       As String = "SERVICO"
    Const COL_VALOR         As String = "VALOR"
    Const COL_TIPO          As String = "TIPO"

    Const PROJ_COL_CHAVE    As String = "CHAVE"
    Const PROJ_COL_FREQ     As String = "FREQUENCIA"
    Const PROJ_COL_DATA     As String = "DATA_PROJETADA"
    Const PROJ_COL_DEC      As String = "DECISAO"
    Const PROJ_COL_DATA_DEC As String = "DATA_DECISAO"

    Const FREQ_PADRAO       As String = "MENSAL"

    ' =========================
    ' VARIÁVEIS
    ' =========================
    Dim tblDados    As ListObject
    Dim tblProj     As ListObject
    Dim hoje        As Date
    Dim totalCriado As Long

    hoje = Date
    totalCriado = 0

    ' =========================
    ' LOCALIZAR TABELAS
    ' =========================
    Set tblDados = EncontrarTabela(TBL_DADOS)
    If tblDados Is Nothing Then
        MsgBox "Tabela " & TBL_DADOS & " não encontrada!", vbCritical
        Exit Sub
    End If
    If tblDados.DataBodyRange Is Nothing Then
        MsgBox "Tabela " & TBL_DADOS & " está vazia!", vbCritical
        Exit Sub
    End If

    Set tblProj = EncontrarTabela(TBL_PROJECAO)
    If tblProj Is Nothing Then
        MsgBox "Tabela " & TBL_PROJECAO & " não encontrada!", vbCritical
        Exit Sub
    End If

    ' =========================
    ' PERFORMANCE
    ' =========================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' =========================
    ' MAPEAR COLUNAS DA TABELA DADOS
    ' =========================
    Dim colData     As ListColumn
    Dim colStatus   As ListColumn
    Dim colFreq     As ListColumn
    Dim colResp     As ListColumn
    Dim colServ     As ListColumn
    Dim colValor    As ListColumn
    Dim colTipo     As ListColumn

    On Error Resume Next
    Set colData = tblDados.ListColumns(COL_DATA)
    Set colStatus = tblDados.ListColumns(COL_STATUS)
    Set colFreq = tblDados.ListColumns(COL_FREQUENCIA)
    Set colResp = tblDados.ListColumns(COL_RESPONSAVEL)
    Set colServ = tblDados.ListColumns(COL_SERVICO)
    Set colValor = tblDados.ListColumns(COL_VALOR)
    Set colTipo = tblDados.ListColumns(COL_TIPO)
    On Error GoTo CleanUp

    If colData Is Nothing Or colResp Is Nothing Or colServ Is Nothing Or _
       colValor Is Nothing Or colTipo Is Nothing Then
        MsgBox "Colunas obrigatórias não encontradas em " & TBL_DADOS, vbCritical
        GoTo CleanUp
    End If

    ' =========================
    ' CARREGAR ÚLTIMA DATA POR CHAVE
    ' =========================
    Dim dictUltimaData  As Object
    Dim dictFrequencia  As Object
    Set dictUltimaData = CreateObject("Scripting.Dictionary")
    Set dictFrequencia = CreateObject("Scripting.Dictionary")
    dictUltimaData.CompareMode = vbTextCompare
    dictFrequencia.CompareMode = vbTextCompare

    Dim i           As Long
    Dim chave       As String
    Dim freq        As String
    Dim dataVal     As Date
    Dim valorStr    As String
    Dim numVal      As Double

    For i = 1 To tblDados.ListRows.Count

        If Not IsDate(tblDados.DataBodyRange.Cells(i, colData.Index).Value) Then GoTo ProximaLinha1

        Dim respVal As String
        Dim servVal As String
        Dim tipoVal As String

        respVal = Trim$(UCase$(CStr(tblDados.DataBodyRange.Cells(i, colResp.Index).Value)))
        servVal = Trim$(UCase$(CStr(tblDados.DataBodyRange.Cells(i, colServ.Index).Value)))
        tipoVal = Trim$(UCase$(CStr(tblDados.DataBodyRange.Cells(i, colTipo.Index).Value)))

        numVal = 0
        On Error Resume Next
        numVal = CDbl(tblDados.DataBodyRange.Cells(i, colValor.Index).Value)
        On Error GoTo CleanUp
        valorStr = Format(numVal, "#,##0.00")

        chave = respVal & "|" & servVal & "|" & valorStr & "|" & tipoVal

        freq = FREQ_PADRAO
        If Not colFreq Is Nothing Then
            Dim freqRaw As String
            freqRaw = Trim$(UCase$(CStr(tblDados.DataBodyRange.Cells(i, colFreq.Index).Value)))
            If freqRaw <> "" Then freq = freqRaw
        End If

        dataVal = CDate(tblDados.DataBodyRange.Cells(i, colData.Index).Value)

        If Not dictUltimaData.Exists(chave) Then
            dictUltimaData.Add chave, dataVal
            dictFrequencia.Add chave, freq
        Else
            If dataVal > CDate(dictUltimaData(chave)) Then
                dictUltimaData(chave) = dataVal
                dictFrequencia(chave) = freq
            End If
        End If

ProximaLinha1:
    Next i

    ' =========================
    ' CALCULAR PRÓXIMA DATA E VERIFICAR PENDENTES
    ' SEMANAL/QUINZENAL  -> compara dia exato  (proxData <= hoje)
    ' MENSAL em diante   -> compara mês/ano    (mês de competência já chegou)
    ' =========================
    Dim dictPendentes As Object
    Set dictPendentes = CreateObject("Scripting.Dictionary")
    dictPendentes.CompareMode = vbTextCompare

    Dim chaveItem   As Variant
    Dim ultimaData  As Date
    Dim proxData    As Date
    Dim freqItem    As String
    Dim novaCol     As Collection

    For Each chaveItem In dictUltimaData.Keys

        ultimaData = CDate(dictUltimaData(chaveItem))
        freqItem = CStr(dictFrequencia(chaveItem))
        proxData = CalcularProximaData(ultimaData, freqItem)

        If DeveProjetar(proxData, freqItem, hoje) Then
            If Not JaFoiProjetada(tblProj, CStr(chaveItem), proxData, PROJ_COL_CHAVE, PROJ_COL_DATA) Then
                If Not dictPendentes.Exists(freqItem) Then
                    Set novaCol = New Collection
                    dictPendentes.Add freqItem, novaCol
                End If
                dictPendentes(freqItem).Add Array(CStr(chaveItem), proxData)
            End If
        End If

    Next chaveItem

    ' =========================
    ' NADA PARA PROJETAR
    ' =========================
    If dictPendentes.Count = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "Nenhum lançamento fixo precisa ser projetado no momento.", _
               vbInformation, "Projeção de Lançamentos Fixos"
        Exit Sub
    End If

    ' =========================
    ' PERGUNTAR EM BLOCO POR FREQUÊNCIA
    ' =========================
    Dim freqBloco      As Variant
    Dim cancelado      As Boolean
    Dim houveInteracao As Boolean
    cancelado = False
    houveInteracao = False

    For Each freqBloco In dictPendentes.Keys

        Dim itensBloco  As Collection
        Set itensBloco = dictPendentes(freqBloco)

        Dim qtd As Long
        qtd = itensBloco.Count

        Dim r As VbMsgBoxResult
        r = MsgBox( _
            "========================================" & vbCrLf & _
            "   PROJEÇÃO DE LANÇAMENTOS FIXOS" & vbCrLf & _
            "========================================" & vbCrLf & vbCrLf & _
            "Existem " & qtd & " item(ns) com frequência " & CStr(freqBloco) & " prontos para projeção." & vbCrLf & vbCrLf & _
            "----------------------------------------" & vbCrLf & _
            "SIM      = projeta e registra como decidido" & vbCrLf & _
            "NÃO      = não projeta e registra como decidido" & vbCrLf & _
            "CANCELAR = decidir depois (pausa processo)", _
            vbYesNoCancel + vbQuestion + vbDefaultButton3, _
            "Projeção - " & CStr(freqBloco))

        If r = vbCancel Then
            cancelado = True
            Exit For
        End If

        houveInteracao = True

        If r = vbYes Then

            Dim itemArr As Variant
            For Each itemArr In itensBloco

                Dim chaveProj   As String
                Dim dataProj    As Date
                chaveProj = CStr(itemArr(0))
                dataProj = CDate(itemArr(1))

                Dim linhaOrig As Long
                linhaOrig = EncontrarLinhaOrigem( _
                    tblDados, chaveProj, _
                    colData.Index, colResp.Index, colServ.Index, colValor.Index, colTipo.Index, _
                    CDate(dictUltimaData(chaveProj)))

                If linhaOrig > 0 Then
                    Dim nl As ListRow
                    Set nl = tblDados.ListRows.Add
                    tblDados.DataBodyRange.Rows(linhaOrig).Copy nl.Range
                    Application.CutCopyMode = False
                    nl.Range.Cells(1, colData.Index).Value = dataProj
                    If Not colStatus Is Nothing Then
                        nl.Range.Cells(1, colStatus.Index).Value = "PENDENTE"
                    End If
                    If Not colFreq Is Nothing Then
                        nl.Range.Cells(1, colFreq.Index).Value = CStr(dictFrequencia(chaveProj))
                    End If
                    totalCriado = totalCriado + 1

                    GravarProjecao tblProj, chaveProj, CStr(freqBloco), dataProj, "SIM", _
                        PROJ_COL_CHAVE, PROJ_COL_FREQ, PROJ_COL_DATA, PROJ_COL_DEC, PROJ_COL_DATA_DEC
                End If

            Next itemArr

        ElseIf r = vbNo Then

            For Each itemArr In itensBloco
                GravarProjecao tblProj, CStr(itemArr(0)), CStr(freqBloco), CDate(itemArr(1)), "NAO", _
                    PROJ_COL_CHAVE, PROJ_COL_FREQ, PROJ_COL_DATA, PROJ_COL_DEC, PROJ_COL_DATA_DEC
            Next itemArr

        End If

    Next freqBloco

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If houveInteracao Or cancelado Then
        If cancelado Then
            MsgBox "Processo pausado pelo usuário." & vbCrLf & _
                   "Lançamentos criados até agora: " & totalCriado, _
                   vbInformation, "Pausado - Lançamentos Fixos"
        Else
            MsgBox "Processo concluído!" & vbCrLf & vbCrLf & _
                   "Total de lançamentos fixos criados: " & totalCriado, _
                   vbInformation, "Concluído - Lançamentos Fixos"
        End If
    End If

End Sub

' ==============================================================================
' DECIDE SE DEVE PROJETAR CONFORME TIPO DE FREQUÊNCIA
'   SEMANAL / QUINZENAL -> dia exato (proxData <= hoje)
'   MENSAL em diante    -> mês de competência já chegou (mes/ano <= mes/ano de hoje)
' ==============================================================================
Private Function DeveProjetar(proxData As Date, freq As String, hoje As Date) As Boolean
    Select Case UCase$(Trim$(freq))
        Case "SEMANAL", "QUINZENAL"
            DeveProjetar = (proxData <= hoje)
        Case Else
            ' MENSAL, BIMESTRAL, TRIMESTRAL, SEMESTRAL, ANUAL
            Dim anoProx As Long: anoProx = Year(proxData)
            Dim mesProx As Long: mesProx = Month(proxData)
            Dim anoHoje As Long: anoHoje = Year(hoje)
            Dim mesHoje As Long: mesHoje = Month(hoje)
            DeveProjetar = (anoProx < anoHoje) Or (anoProx = anoHoje And mesProx <= mesHoje)
    End Select
End Function

' ==============================================================================
' CALCULAR PRÓXIMA DATA CONFORME FREQUÊNCIA
' ==============================================================================
Private Function CalcularProximaData(ultimaData As Date, freq As String) As Date
    Select Case UCase$(Trim$(freq))
        Case "SEMANAL"
            CalcularProximaData = ultimaData + 7
        Case "QUINZENAL"
            CalcularProximaData = ultimaData + 15
        Case "MENSAL"
            CalcularProximaData = DateSerial(Year(ultimaData), Month(ultimaData) + 1, Day(ultimaData))
        Case "BIMESTRAL"
            CalcularProximaData = DateSerial(Year(ultimaData), Month(ultimaData) + 2, Day(ultimaData))
        Case "TRIMESTRAL"
            CalcularProximaData = DateSerial(Year(ultimaData), Month(ultimaData) + 3, Day(ultimaData))
        Case "SEMESTRAL"
            CalcularProximaData = DateSerial(Year(ultimaData), Month(ultimaData) + 6, Day(ultimaData))
        Case "ANUAL"
            CalcularProximaData = DateSerial(Year(ultimaData) + 1, Month(ultimaData), Day(ultimaData))
        Case Else
            CalcularProximaData = DateSerial(Year(ultimaData), Month(ultimaData) + 1, Day(ultimaData))
    End Select
End Function

' ==============================================================================
' VERIFICA SE JÁ FOI PROJETADA (CHAVE + DATA_PROJETADA)
' ==============================================================================
Private Function JaFoiProjetada( _
    tblProj As ListObject, _
    chave As String, _
    dataProj As Date, _
    colChaveNome As String, _
    colDataNome As String _
) As Boolean

    JaFoiProjetada = False
    If tblProj.DataBodyRange Is Nothing Then Exit Function

    Dim colCH As ListColumn
    Dim colDT As ListColumn

    On Error Resume Next
    Set colCH = tblProj.ListColumns(colChaveNome)
    Set colDT = tblProj.ListColumns(colDataNome)
    On Error GoTo 0

    If colCH Is Nothing Or colDT Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To tblProj.ListRows.Count
        If Trim$(CStr(tblProj.DataBodyRange.Cells(i, colCH.Index).Value)) = chave Then
            If IsDate(tblProj.DataBodyRange.Cells(i, colDT.Index).Value) Then
                If CDate(tblProj.DataBodyRange.Cells(i, colDT.Index).Value) = dataProj Then
                    JaFoiProjetada = True
                    Exit Function
                End If
            End If
        End If
    Next i

End Function

' ==============================================================================
' ENCONTRAR LINHA ORIGEM NA TABELA DE DADOS
' ==============================================================================
Private Function EncontrarLinhaOrigem( _
    tbl As ListObject, _
    chave As String, _
    idxData As Long, idxResp As Long, idxServ As Long, idxValor As Long, idxTipo As Long, _
    ultimaData As Date _
) As Long

    EncontrarLinhaOrigem = 0

    Dim i           As Long
    Dim numVal      As Double
    Dim valorStr    As String
    Dim chaveAtual  As String

    For i = 1 To tbl.ListRows.Count
        If IsDate(tbl.DataBodyRange.Cells(i, idxData).Value) Then
            If CDate(tbl.DataBodyRange.Cells(i, idxData).Value) = ultimaData Then

                numVal = 0
                On Error Resume Next
                numVal = CDbl(tbl.DataBodyRange.Cells(i, idxValor).Value)
                On Error GoTo 0
                valorStr = Format(numVal, "#,##0.00")

                chaveAtual = Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, idxResp).Value))) & "|" & _
                             Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, idxServ).Value))) & "|" & _
                             valorStr & "|" & _
                             Trim$(UCase$(CStr(tbl.DataBodyRange.Cells(i, idxTipo).Value)))

                If chaveAtual = chave Then
                    EncontrarLinhaOrigem = i
                    Exit Function
                End If

            End If
        End If
    Next i

End Function

' ==============================================================================
' GRAVAR PROJEÇÃO NA tbl_Projecao_Lancamentos_Fixos
' ==============================================================================
Private Sub GravarProjecao( _
    tblProj As ListObject, _
    chave As String, _
    freq As String, _
    dataProj As Date, _
    decisao As String, _
    colChaveNome As String, _
    colFreqNome As String, _
    colDataNome As String, _
    colDecNome As String, _
    colDataDecNome As String _
)
    Dim colCH   As ListColumn
    Dim colFR   As ListColumn
    Dim colDT   As ListColumn
    Dim colDEC  As ListColumn
    Dim colDDEC As ListColumn

    On Error Resume Next
    Set colCH = tblProj.ListColumns(colChaveNome)
    Set colFR = tblProj.ListColumns(colFreqNome)
    Set colDT = tblProj.ListColumns(colDataNome)
    Set colDEC = tblProj.ListColumns(colDecNome)
    Set colDDEC = tblProj.ListColumns(colDataDecNome)
    On Error GoTo 0

    If colCH Is Nothing Or colDT Is Nothing Or colDEC Is Nothing Then Exit Sub

    Dim lr      As ListRow
    Dim i       As Long
    Dim achou   As Boolean
    achou = False

    If Not tblProj.DataBodyRange Is Nothing Then
        For i = 1 To tblProj.ListRows.Count
            If Trim$(CStr(tblProj.DataBodyRange.Cells(i, colCH.Index).Value)) = "" Then
                Set lr = tblProj.ListRows(i)
                achou = True
                Exit For
            End If
        Next i
    End If

    If Not achou Then Set lr = tblProj.ListRows.Add

    lr.Range.Cells(1, colCH.Index).Value = chave
    If Not colFR Is Nothing Then lr.Range.Cells(1, colFR.Index).Value = freq
    lr.Range.Cells(1, colDT.Index).Value = dataProj
    lr.Range.Cells(1, colDEC.Index).Value = decisao
    If Not colDDEC Is Nothing Then lr.Range.Cells(1, colDDEC.Index).Value = Date

End Sub

' ==============================================================================
' HELPER: ENCONTRAR TABELA EM QUALQUER ABA
' ==============================================================================
Private Function EncontrarTabela(nome As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set EncontrarTabela = ws.ListObjects(nome)
        On Error GoTo 0
        If Not EncontrarTabela Is Nothing Then Exit Function
    Next ws
End Function

