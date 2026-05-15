Attribute VB_Name = "modCoresCategoriasServicos"
Option Explicit

' ==============================================================================
' MÓDULO: modCoresCategoriasServicos
' DESCRIÇÃO: Aplicar cores e organizar tabelas de lançamentos (Fixos e Eventuais)
' ==============================================================================

Public Sub AplicarCores_Categorias_Servicos_Automatico()
    ' ==============================================================================
    ' MACRO: APLICAR CORES E ORGANIZAR TABELAS DE LANÇAMENTOS
    ' ==============================================================================

    ' ==============================================================================
    ' CONFIGURAÇÕES CENTRALIZADAS
    ' ==============================================================================
    Const TABELA_1 As String = "tbl_Lancamentos_Fixos"
    Const TABELA_2 As String = "tbl_Lancamentos_Eventuais"
    Const COLUNA_CATEGORIA As String = "CATEGORIA"
    Const COLUNA_SERVICO As String = "SERVICO"
    Const COLUNA_DATA As String = "DATA_EFETIVA"

    ' ==============================================================================
    ' OBTER PALETA DE CORES CENTRALIZADA
    ' ==============================================================================
    Dim paleta As Object
    Set paleta = ObterPaletaCores()

    ' ==============================================================================
    ' OTIMIZAÇÃO
    ' ==============================================================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ==============================================================================
    ' PROCESSAR TABELAS
    ' ==============================================================================
    ProcessarTabela TABELA_1, COLUNA_CATEGORIA, COLUNA_SERVICO, COLUNA_DATA, paleta
    ProcessarTabela TABELA_2, COLUNA_CATEGORIA, COLUNA_SERVICO, COLUNA_DATA, paleta

    ' ==============================================================================
    ' RESTAURAR CONFIGURAÇÕES
    ' ==============================================================================
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

' ==============================================================================
' PROCESSAR UMA TABELA
' ==============================================================================
Private Sub ProcessarTabela(nomeTabela As String, nomeColCategoria As String, nomeColServico As String, nomeColData As String, paleta As Object)

    Dim tbl As ListObject
    Dim colCat As ListColumn, colServ As ListColumn, colData As ListColumn
    Dim ws As Worksheet

    ' =========================
    ' LOCALIZAR TABELA
    ' =========================
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = ws.ListObjects(nomeTabela)
        If Not tbl Is Nothing Then Exit For
    Next ws
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' =========================
    ' LOCALIZAR COLUNAS
    ' =========================
    On Error Resume Next
    Set colCat = tbl.ListColumns(nomeColCategoria)
    Set colServ = tbl.ListColumns(nomeColServico)
    Set colData = tbl.ListColumns(nomeColData)
    On Error GoTo 0

    If colCat Is Nothing Then Exit Sub

    ' =========================
    ' ORDENAR
    ' =========================
    If Not colData Is Nothing Then
        OrdenarTabela tbl, colData, colCat
    End If

    ' =========================
    ' APLICAR CORES
    ' =========================
    colCat.DataBodyRange.FormatConditions.Delete
    If Not colServ Is Nothing Then colServ.DataBodyRange.FormatConditions.Delete

    AplicarCoresFormatacao tbl, colCat, colServ, paleta

End Sub

' ==============================================================================
' ORDENAR TABELA
' ==============================================================================
Private Sub OrdenarTabela(tbl As ListObject, colData As ListColumn, colCat As ListColumn)

    tbl.Sort.SortFields.Clear

    tbl.Sort.SortFields.Add2 _
        Key:=colData.DataBodyRange, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    tbl.Sort.SortFields.Add2 _
        Key:=colCat.DataBodyRange, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    With tbl.Sort
        .Header = xlYes
        .Apply
    End With

End Sub

' ==============================================================================
' APLICAR FORMATAÇÃO CONDICIONAL
' ==============================================================================
Private Sub AplicarCoresFormatacao(tbl As ListObject, colCat As ListColumn, colServ As ListColumn, paleta As Object)

    Dim celReferencia As String
    celReferencia = colCat.DataBodyRange.Cells(1, 1).Address(False, False)

    Dim categoria As Variant
    For Each categoria In paleta.Keys

        With colCat.DataBodyRange.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=" & celReferencia & "=""" & categoria & """")
            .Interior.Color = paleta(categoria)
        End With

        If Not colServ Is Nothing Then
            With colServ.DataBodyRange.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:="=" & celReferencia & "=""" & categoria & """")
                .Interior.Color = paleta(categoria)
            End With
        End If

    Next categoria

End Sub


