Attribute VB_Name = "ImportadorAGHU"
Option Explicit

' =====================================================================
' Importador de dados do AGHU - Painel de Produtividade
' =====================================================================
' Este modulo contem as macros acionadas pelos botoes da aba PAINEL:
'   - "Selecionar arquivo do Periodo 1" -> ImportarPeriodo1
'   - "Selecionar arquivo do Periodo 2" -> ImportarPeriodo2
'
' O que cada macro faz:
'   1) Abre a janela padrao do Windows para o usuario escolher o
'      arquivo .xlsx exportado do AGHU (em qualquer pasta do computador).
'   2) Confere se o arquivo tem as colunas esperadas (Cod. SUS,
'      Descricao, Prontuario, Data Realizado, Qtd, Consulta).
'   3) Substitui o conteudo das colunas A:F da aba correspondente
'      (Dados_Periodo1 ou Dados_Periodo2), apagando linhas antigas
'      que sobrarem se o arquivo novo tiver menos linhas.
'   4) NAO toca nas colunas G:J (formulas auxiliares) - elas
'      recalculam sozinhas.
'   5) Mostra uma mensagem confirmando quantas linhas foram
'      importadas, e grava esse resumo na aba PAINEL.
' =====================================================================

Private Const EXPECTED_HEADERS As String = _
    "Cód. SUS|Descrição|Prontuário|Data Realizado|Qtd|Consulta"
Private Const MAX_ROWS As Long = 12000   ' linhas de dados suportadas (linha 2 ate 12001)

Sub ImportarPeriodo1()
    ImportarDados "Dados_Periodo1", "Período 1", "B60"
End Sub

Sub ImportarPeriodo2()
    ImportarDados "Dados_Periodo2", "Período 2", "B61"
End Sub

Private Sub ImportarDados(ByVal nomeAba As String, ByVal rotulo As String, ByVal celulaStatus As String)
    Dim caminhoArquivo As Variant
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long
    Dim numLinhas As Long
    Dim dadosArray As Variant

    On Error GoTo TratarErro

    ' 1) Selecionar arquivo (janela padrao do Windows/Excel)
    caminhoArquivo = Application.GetOpenFilename( _
        FileFilter:="Arquivos Excel (*.xlsx;*.xls), *.xlsx;*.xls", _
        Title:="Selecione o arquivo exportado do AGHU - " & rotulo, _
        MultiSelect:=False)

    If VarType(caminhoArquivo) = vbBoolean Then
        Exit Sub ' usuario clicou em Cancelar
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 2) Abrir o arquivo de origem, somente leitura, sem atualizar links
    Set wbOrigem = Workbooks.Open(Filename:=caminhoArquivo, ReadOnly:=True, _
                                   UpdateLinks:=0, AddToMru:=False)
    Set wsOrigem = wbOrigem.Sheets(1)

    ' 3) Validar cabecalho das colunas A a F
    If Not CabecalhoValido(wsOrigem) Then
        wbOrigem.Close SaveChanges:=False
        RestaurarAplicacao
        MsgBox "O arquivo selecionado não tem o formato esperado." & vbCrLf & vbCrLf & _
               "As colunas A a F da primeira planilha do arquivo devem ser, nesta ordem:" & vbCrLf & _
               Replace(EXPECTED_HEADERS, "|", ", ") & vbCrLf & vbCrLf & _
               "Verifique se você selecionou o arquivo certo, exportado do AGHU para " & rotulo & ", e tente novamente.", _
               vbExclamation, "Arquivo inválido"
        Exit Sub
    End If

    ' 4) Determinar quantas linhas de dados existem (coluna A, a partir da linha 2)
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    numLinhas = ultimaLinhaOrigem - 1

    If numLinhas <= 0 Then
        wbOrigem.Close SaveChanges:=False
        RestaurarAplicacao
        MsgBox "O arquivo selecionado não contém nenhuma linha de dados abaixo do cabeçalho.", _
               vbExclamation, "Arquivo vazio"
        Exit Sub
    End If

    If numLinhas > MAX_ROWS Then
        wbOrigem.Close SaveChanges:=False
        RestaurarAplicacao
        MsgBox "O arquivo selecionado tem " & numLinhas & " linhas de dados, acima do limite" & vbCrLf & _
               "atualmente preparado nesta planilha (" & MAX_ROWS & " linhas)." & vbCrLf & vbCrLf & _
               "Fale com o suporte técnico antes de importar este arquivo.", _
               vbCritical, "Limite excedido"
        Exit Sub
    End If

    ' 5) Ler os dados de origem (colunas A a F) para memoria
    dadosArray = wsOrigem.Range("A2:F" & ultimaLinhaOrigem).Value

    ' 5-b) Normalizar a coluna D (Data Realizado) para texto "dd/mm/aaaa hh:mm".
    '      O AGHU normalmente exporta essa coluna como TEXTO, mas algumas
    '      linhas do arquivo podem vir como data/hora de verdade (o Excel
    '      as vezes reconhece automaticamente parte dos valores ao gerar o
    '      arquivo). Quando isso acontece, o VBA le a celula como um valor
    '      de data (Date), e simplesmente copiar esse valor adiante faria o
    '      Excel reescreve-lo como texto de acordo com a configuracao
    '      regional do Windows (podendo inverter dia e mes). Aqui, qualquer
    '      valor que chegue como Date e explicitamente formatado como texto
    '      "dd/mm/aaaa hh:mm" usando o Format() do VBA - que usa sempre essa
    '      ordem fixa, independente da configuracao regional do computador.
    '      Linhas que ja chegam como texto (o caso mais comum) nao sao
    '      alteradas aqui.
    Dim linhaDados As Long
    For linhaDados = 1 To UBound(dadosArray, 1)
        If VarType(dadosArray(linhaDados, 4)) = vbDate Then
            dadosArray(linhaDados, 4) = Format(dadosArray(linhaDados, 4), "dd/mm/yyyy hh:nn")
        End If
    Next linhaDados

    ' 5-c) Normalizar as colunas numericas (Cod. SUS, Prontuario, Qtd, Consulta)
    '      para NUMERO de verdade. O AGHU pode exportar essas colunas como
    '      texto em algumas linhas e como numero em outras, dentro do MESMO
    '      arquivo (a mesma inconsistencia que acontece com a coluna Data).
    '      Se isso ficar misturado, as formulas de COUNTIFS/SUMIFS que
    '      comparam a coluna Consulta (F) e somam a coluna Qtd (E) deixam de
    '      reconhecer que duas linhas sao do mesmo atendimento - porque, no
    '      Excel, o texto "50404199" e o numero 50404199 nao sao sempre
    '      tratados como iguais nessas formulas. Isso fazia o indicador
    '      "Atendimento com mais de 1 procedimento" ficar zerado mesmo
    '      havendo procedimentos repetidos - de forma inconsistente entre
    '      Periodo 1 e Periodo 2, dependendo de como cada arquivo do AGHU
    '      veio exportado.
    Dim colunasNumericas As Variant
    Dim idx As Long
    colunasNumericas = Array(1, 3, 5, 6) ' A=Cod.SUS, C=Prontuario, E=Qtd, F=Consulta
    For linhaDados = 1 To UBound(dadosArray, 1)
        For idx = LBound(colunasNumericas) To UBound(colunasNumericas)
            If VarType(dadosArray(linhaDados, colunasNumericas(idx))) = vbString Then
                If IsNumeric(dadosArray(linhaDados, colunasNumericas(idx))) Then
                    dadosArray(linhaDados, colunasNumericas(idx)) = CDbl(dadosArray(linhaDados, colunasNumericas(idx)))
                End If
            End If
        Next idx
    Next linhaDados

    ' 6) Fechar o arquivo de origem sem salvar
    wbOrigem.Close SaveChanges:=False
    Set wbOrigem = Nothing

    ' 7) Limpar dados antigos nas colunas A:F (linhas 2 a 12001), sem tocar em G:J
    Set wsDestino = ThisWorkbook.Sheets(nomeAba)
    wsDestino.Range("A2:F" & (MAX_ROWS + 1)).ClearContents

    ' 7-b) Forcar a coluna D (Data Realizado) como TEXTO antes de escrever.
    '      Sem isso, o Excel reinterpreta o texto "dd/mm/aaaa hh:mm" vindo do
    '      AGHU de acordo com a configuracao regional do Windows: em um
    '      computador configurado para EUA (mm/dd/aaaa), datas ambiguas
    '      (dia <= 12) sao gravadas com dia e mes invertidos. Mantendo a
    '      coluna como texto, o valor original do AGHU e preservado
    '      exatamente, e a formula da coluna G (que ja sabe interpretar
    '      texto "dd/mm/aaaa") continua funcionando normalmente.
    wsDestino.Range("D2:D" & (MAX_ROWS + 1)).NumberFormat = "@"

    ' 8) Escrever os dados novos
    wsDestino.Range("A2").Resize(numLinhas, 6).Value = dadosArray

    ' 9) Registrar resumo da importação na aba PAINEL
    On Error Resume Next
    ThisWorkbook.Sheets("PAINEL").Range(celulaStatus).Value = _
        rotulo & ": " & numLinhas & " linha(s) importada(s) em " & Format(Now, "dd/mm/yyyy hh:mm")
    On Error GoTo TratarErro

    RestaurarAplicacao

    MsgBox "Importação concluída para " & rotulo & "!" & vbCrLf & vbCrLf & _
           numLinhas & " linha(s) importada(s) para a aba """ & nomeAba & """.", _
           vbInformation, "Importação concluída"

    Exit Sub

TratarErro:
    RestaurarAplicacao
    If Not wbOrigem Is Nothing Then
        On Error Resume Next
        wbOrigem.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "Ocorreu um erro durante a importação:" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Erro na importação"
End Sub

Private Function CabecalhoValido(ByVal ws As Worksheet) As Boolean
    Dim esperado() As String
    Dim i As Long
    esperado = Split(EXPECTED_HEADERS, "|")
    CabecalhoValido = True
    For i = 0 To UBound(esperado)
        If Trim$(CStr(ws.Cells(1, i + 1).Value)) <> esperado(i) Then
            CabecalhoValido = False
            Exit Function
        End If
    Next i
End Function

Private Sub RestaurarAplicacao()
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
