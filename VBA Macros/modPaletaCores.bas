Attribute VB_Name = "modPaletaCores"
Option Explicit
' ==============================================================================
' MÓDULO: modPaletaCores
' DESCRIÇÃO: Paleta de cores lida dinamicamente da aba Configs
' O usuário pinta o FUNDO da coluna CATEGORIA_LANCAMENTO com a cor desejada.
' A cor lida do FUNDO será aplicada como fundo nas linhas dos lançamentos.
' ==============================================================================
Public Function ObterPaletaCores() As Object
    Dim paleta As Object
    Set paleta = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colCat As ListColumn
    Dim cel As Range
    Dim categoria As String
    Dim corFundo As Long
    
    ' =========================
    ' LOCALIZAR TABELA
    ' =========================
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = ws.ListObjects("tbl_Base_Sistema")
        If Not tbl Is Nothing Then Exit For
    Next ws
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Set ObterPaletaCores = paleta
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        Set ObterPaletaCores = paleta
        Exit Function
    End If
    
    ' =========================
    ' LOCALIZAR COLUNA
    ' =========================
    On Error Resume Next
    Set colCat = tbl.ListColumns("CATEGORIA_LANCAMENTO")
    On Error GoTo 0
    
    If colCat Is Nothing Then
        Set ObterPaletaCores = paleta
        Exit Function
    End If
    
    ' =========================
    ' LER CATEGORIAS E COR DE FUNDO
    ' =========================
    For Each cel In colCat.DataBodyRange
        categoria = Trim(UCase(CStr(cel.Value)))
        corFundo = cel.Interior.Color
        
        ' Ignora células sem preenchimento (ColorIndex = xlNone) e categorias vazias
        If categoria <> "" And cel.Interior.ColorIndex <> xlNone And Not paleta.Exists(categoria) Then
            paleta.Add categoria, corFundo
        End If
    Next cel
    
    Set ObterPaletaCores = paleta
End Function
