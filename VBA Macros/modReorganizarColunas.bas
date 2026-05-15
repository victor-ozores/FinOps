Attribute VB_Name = "modReorganizarColunas"
Sub Reorganizar_Colunas_E_Valores_Alfabetico()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblNomes(1) As String
    Dim t As Integer
    Dim i As Integer, j As Integer
    Dim numColunas As Integer
    Dim nomesColunas() As String
    Dim colunasOrdenadas() As String
    Dim tentativas As Integer
    Dim posicaoAtual As Integer
    Dim posicaoDesejada As Integer
    Dim col As ListColumn
    Dim rng As Range
    Dim arr As Variant
    Dim k As Long, m As Long
    Dim tempVal As Variant
    
    ' Variáveis novas para lidar com as cores
    Dim checkColor As Boolean
    Dim arrColors() As Variant
    Dim tempColor As Variant
    Dim cIdx As Long

    ' ========== TABELAS ALVO ==========
    tblNomes(0) = "tbl_Base_User"
    tblNomes(1) = "tbl_Base_Sistema"

    ' ========== COLUNAS DA tbl_Base_Sistema QUE NAO SERAO ORDENADAS ==========
    Dim colunasIgnorar(4) As String
    colunasIgnorar(0) = "METODO_PAGAMENTO"
    colunasIgnorar(1) = "STATUS_GENERICO"
    colunasIgnorar(2) = "STATUS_PAGAMENTO"
    colunasIgnorar(3) = "TIPO_LANCAMENTO"
    colunasIgnorar(4) = "FREQUENCIA"

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Base_Opcoes")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Planilha Base_Opcoes não encontrada!", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ========== LOOP NAS 2 TABELAS ==========
    For t = 0 To 1

        On Error Resume Next
        Set tbl = ws.ListObjects(tblNomes(t))
        On Error GoTo 0

        If tbl Is Nothing Then
            MsgBox "Tabela " & tblNomes(t) & " não encontrada!", vbCritical
            GoTo ProximaTabela
        End If

        ' ========== PARTE 1: REORGANIZAR COLUNAS EM ORDEM ALFABÉTICA ==========

        numColunas = tbl.ListColumns.Count
        ReDim nomesColunas(1 To numColunas)
        ReDim colunasOrdenadas(1 To numColunas)

        For i = 1 To numColunas
            nomesColunas(i) = tbl.ListColumns(i).Name
            colunasOrdenadas(i) = nomesColunas(i)
        Next i

        ' Ordena nomes das colunas
        For i = 1 To numColunas - 1
            For j = i + 1 To numColunas
                If colunasOrdenadas(i) > colunasOrdenadas(j) Then
                    Dim tmp As String
                    tmp = colunasOrdenadas(i)
                    colunasOrdenadas(i) = colunasOrdenadas(j)
                    colunasOrdenadas(j) = tmp
                End If
            Next j
        Next i

        ' Reorganiza colunas fisicamente
        For tentativas = 1 To 3
            For i = 1 To numColunas
                posicaoDesejada = i
                On Error Resume Next
                posicaoAtual = tbl.ListColumns(colunasOrdenadas(i)).Index
                On Error GoTo 0

                If posicaoAtual > 0 And posicaoAtual <> posicaoDesejada Then
                    tbl.ListColumns(colunasOrdenadas(i)).Range.Cut
                    tbl.Range.Cells(1, posicaoDesejada).Insert Shift:=xlToRight
                    Application.CutCopyMode = False
                End If
            Next i
        Next tentativas

        ' ========== PARTE 2: LIMPAR, PADRONIZAR E ORDENAR VALORES ==========

        If Not tbl.DataBodyRange Is Nothing Then

            For Each col In tbl.ListColumns
                Set rng = col.DataBodyRange

                If Not rng Is Nothing And rng.Cells.Count > 1 Then

                    arr = rng.Value
                    
                    ' Verifica se é a coluna que precisa levar a cor junto
                    checkColor = (UCase(col.Name) = "CATEGORIA_LANCAMENTO")
                    If checkColor Then
                        ReDim arrColors(1 To rng.Rows.Count)
                        For cIdx = 1 To rng.Rows.Count
                            ' Salva -1 se a célula não tiver cor, caso contrário salva a cor exata
                            If rng.Cells(cIdx, 1).Interior.ColorIndex = xlNone Then
                                arrColors(cIdx) = -1
                            Else
                                arrColors(cIdx) = rng.Cells(cIdx, 1).Interior.Color
                            End If
                        Next cIdx
                    End If

                    ' LIMPEZA + PADRONIZAÇÃO
                    For k = LBound(arr) To UBound(arr)
                        If Not IsEmpty(arr(k, 1)) And arr(k, 1) <> "" Then
                            arr(k, 1) = Trim(arr(k, 1))
                            arr(k, 1) = UCase(arr(k, 1))
                            arr(k, 1) = RemoverAcentos(CStr(arr(k, 1)))
                        End If
                    Next k

                    ' ========== VERIFICA SE A COLUNA DEVE SER IGNORADA NA ORDENAÇÃO ==========
                    Dim ignorarOrdenacao As Boolean
                    ignorarOrdenacao = False

                    If t = 1 Then ' Apenas para tbl_Base_Sistema
                        Dim c As Integer
                        For c = 0 To 4
                            If UCase(col.Name) = colunasIgnorar(c) Then
                                ignorarOrdenacao = True
                                Exit For
                            End If
                        Next c
                    End If

                    ' ORDENA VALORES (apenas se a coluna não estiver na lista de ignoradas)
                    If Not ignorarOrdenacao Then
                        For k = LBound(arr) To UBound(arr) - 1
                            For m = k + 1 To UBound(arr)
                                If arr(k, 1) = "" Then
                                    ' Troca Valores
                                    tempVal = arr(k, 1)
                                    arr(k, 1) = arr(m, 1)
                                    arr(m, 1) = tempVal
                                    
                                    ' Troca Cores (Se for CATEGORIA_LANCAMENTO)
                                    If checkColor Then
                                        tempColor = arrColors(k)
                                        arrColors(k) = arrColors(m)
                                        arrColors(m) = tempColor
                                    End If
                                    
                                ElseIf arr(m, 1) <> "" Then
                                    If arr(k, 1) > arr(m, 1) Then
                                        ' Troca Valores
                                        tempVal = arr(k, 1)
                                        arr(k, 1) = arr(m, 1)
                                        arr(m, 1) = tempVal
                                        
                                        ' Troca Cores (Se for CATEGORIA_LANCAMENTO)
                                        If checkColor Then
                                            tempColor = arrColors(k)
                                            arrColors(k) = arrColors(m)
                                            arrColors(m) = tempColor
                                        End If
                                    End If
                                End If
                            Next m
                        Next k
                    End If

                    ' Devolve valores para a planilha
                    rng.Value = arr
                    
                    ' Devolve cores para a planilha (Se for CATEGORIA_LANCAMENTO)
                    If checkColor Then
                        For cIdx = 1 To rng.Rows.Count
                            ' Limpa a cor do fundo SE a linha estiver sem texto ("") OU se a cor original salva era nula (-1)
                            If arr(cIdx, 1) = "" Or arrColors(cIdx) = -1 Then
                                rng.Cells(cIdx, 1).Interior.ColorIndex = xlNone
                            Else
                                rng.Cells(cIdx, 1).Interior.Color = arrColors(cIdx)
                            End If
                        Next cIdx
                    End If

                End If
            Next col

        End If

ProximaTabela:
        Set tbl = Nothing

    Next t

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

' ========== FUNÇÃO AUXILIAR: REMOVER ACENTOS ==========
Function RemoverAcentos(ByVal texto As String) As String
    Dim comAcento As Variant, semAcento As Variant
    Dim i As Long

    comAcento = Array("Á", "À", "Â", "Ã", "Ä", "É", "È", "Ê", "Ë", _
                      "Í", "Ì", "Î", "Ï", "Ó", "Ò", "Ô", "Õ", "Ö", _
                      "Ú", "Ù", "Û", "Ü", "Ç", _
                      "á", "à", "â", "ã", "ä", "é", "è", "ê", "ë", _
                      "í", "ì", "î", "ï", "ó", "ò", "ô", "õ", "ö", _
                      "ú", "ù", "û", "ü", "ç")

    semAcento = Array("A", "A", "A", "A", "A", "E", "E", "E", "E", _
                      "I", "I", "I", "I", "O", "O", "O", "O", "O", _
                      "U", "U", "U", "U", "C", _
                      "A", "A", "A", "A", "A", "E", "E", "E", "E", _
                      "I", "I", "I", "I", "O", "O", "O", "O", "O", _
                      "U", "U", "U", "U", "C")

    For i = LBound(comAcento) To UBound(comAcento)
        texto = Replace(texto, comAcento(i), semAcento(i))
    Next i

    RemoverAcentos = texto
End Function

