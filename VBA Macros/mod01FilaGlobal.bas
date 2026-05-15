Attribute VB_Name = "mod01FilaGlobal"
Option Explicit

' ==============================================================================
' MACRO: PROCESSAMENTO DE FILA GLOBAL
' O parâmetro ExibirAviso define se a confirmação do Ctrl+Z deve aparecer.
' ==============================================================================
Public Sub Processamento_Fila_Global(Optional ByVal ExibirAviso As Boolean = True)
    
    ' 1. VERIFICAÇÃO DE MENSAGEM (TRAVA)
    If ExibirAviso Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("ATENÇÃO: A execução desta fila limpará todo o histórico de 'Desfazer' (Ctrl+Z) do Excel." & vbCrLf & vbCrLf & _
                          "Deseja continuar com o processamento?", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Aviso de Segurança")
        
        If resposta = vbNo Then Exit Sub
    End If

    On Error GoTo ErroHandler
    
    ' PERFORMANCE E SEGURANÇA
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    ' --- FILA DE TAREFAS ---
    Application.StatusBar = "Processando: Reorganizando Colunas e Estrutura..."
    Call Reorganizar_Colunas_E_Valores_Alfabetico

    Application.StatusBar = "Processando: Validando Frequências..."
    Call Preencher_Frequencia_Padrao

    Application.StatusBar = "Processando: Executando Projeção de Lançamentos..."
    Call ProjetarLancamentosFixos_Fila

    Application.StatusBar = "Processando: Gerando Parcelamentos Automáticos..."
    Call Gerar_Parcelamento_Automatico

    Application.StatusBar = "Processando: Validando Cartões de Crédito..."
    Call Validar_Cartao_Credito_Lancamentos

    Application.StatusBar = "Processando: Verificando Status Pendentes..."
    Call Preencher_Status_Pendente

    Application.StatusBar = "Processando: Atualizando Lançamentos Atrasados..."
    Call Atualizar_Status_Atrasado

    Application.StatusBar = "Processando: Aplicando Identidade Visual..."
    Call AplicarCores_Categorias_Servicos_Automatico

    ' FINALIZAÇÃO
    Application.StatusBar = False
    MsgBox "Processamento de fila global finalizado!", vbInformation, "Sincronização"

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

ErroHandler:
    MsgBox "Erro na fila global: " & Err.Description, vbCritical, "Erro de Sistema"
    Resume CleanUp
End Sub


' ==============================================================================
' CHAMADA PELO BOTÃO (INTERFACE)
' ==============================================================================
Public Sub Trigger_Executar_Fila_Global()
    Call Processamento_Fila_Global(True)
End Sub
