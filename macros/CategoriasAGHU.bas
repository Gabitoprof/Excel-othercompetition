Attribute VB_Name = "CategoriasAGHU"
Option Explicit

' =====================================================================
' Sincronizador da lista de categorias - Painel de Produtividade AGHU
' =====================================================================
' Este modulo contem a macro acionada pelo botao "Atualizar Lista de
' Categorias" na aba Categorias.
'
' O que a macro faz:
'   1) Percorre a lista de procedimentos distintos (aba Auxiliar,
'      coluna B - a mesma lista que alimenta o menu suspenso do
'      PAINEL).
'   2) Para cada procedimento que ainda NAO existe na aba Categorias,
'      acrescenta uma linha nova no final da tabela, com a categoria
'      "Não classificado".
'   3) Procedimentos que ja estao na aba Categorias NUNCA sao
'      alterados, movidos ou duplicados - a categoria que o usuario
'      ja marcou fica exatamente como estava.
'
' Isso garante que a classificacao por procedimento sobrevive a
' qualquer atualizacao mensal dos dados (mesmo que a ordem ou a
' quantidade de procedimentos mude de uma competencia para outra),
' porque a busca é sempre pelo NOME do procedimento, nunca pela
' posição/linha.
' =====================================================================

Private Const CAT_PRIMEIRA_LINHA As Long = 8
Private Const CAT_ULTIMA_LINHA As Long = 307   ' mesma capacidade da lista de procedimentos (300 itens)
Private Const CATEGORIA_PADRAO As String = "Não classificado"

Sub SincronizarCategorias()
    Dim wsAux As Worksheet
    Dim wsCat As Worksheet
    Dim ultimaLinhaAux As Long
    Dim ultimaLinhaCat As Long
    Dim i As Long
    Dim procedimento As String
    Dim jaExiste As Variant
    Dim novos As Long

    On Error GoTo TratarErro

    Set wsAux = ThisWorkbook.Sheets("Auxiliar")
    Set wsCat = ThisWorkbook.Sheets("Categorias")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    novos = 0
    ultimaLinhaAux = wsAux.Cells(wsAux.Rows.Count, "B").End(xlUp).Row

    ' posição livre atual na tabela de categorias
    ultimaLinhaCat = wsCat.Cells(wsCat.Rows.Count, "A").End(xlUp).Row
    If ultimaLinhaCat < CAT_PRIMEIRA_LINHA - 1 Then ultimaLinhaCat = CAT_PRIMEIRA_LINHA - 1

    For i = 2 To ultimaLinhaAux
        procedimento = Trim$(CStr(wsAux.Cells(i, "B").Value))
        If procedimento <> "" Then
            jaExiste = Application.Match(procedimento, _
                wsCat.Range("A" & CAT_PRIMEIRA_LINHA & ":A" & CAT_ULTIMA_LINHA), 0)
            If IsError(jaExiste) Then
                If ultimaLinhaCat + 1 > CAT_ULTIMA_LINHA Then
                    MsgBox "Limite de " & (CAT_ULTIMA_LINHA - CAT_PRIMEIRA_LINHA + 1) & _
                           " procedimentos classificáveis atingido." & vbCrLf & _
                           "Fale com o suporte técnico para aumentar esse limite.", _
                           vbExclamation, "Limite atingido"
                    Exit For
                End If
                ultimaLinhaCat = ultimaLinhaCat + 1
                wsCat.Cells(ultimaLinhaCat, "A").Value = procedimento
                wsCat.Cells(ultimaLinhaCat, "B").Value = CATEGORIA_PADRAO
                novos = novos + 1
            End If
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If novos > 0 Then
        MsgBox novos & " procedimento(s) novo(s) adicionado(s) à lista, marcados como """ & _
               CATEGORIA_PADRAO & """." & vbCrLf & vbCrLf & _
               "Role até o final da tabela nesta aba para classificá-los.", _
               vbInformation, "Lista de categorias atualizada"
    Else
        MsgBox "Nenhum procedimento novo encontrado." & vbCrLf & _
               "A lista de categorias já contém todos os procedimentos conhecidos.", _
               vbInformation, "Lista de categorias atualizada"
    End If

    Exit Sub

TratarErro:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro ao atualizar a lista de categorias:" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Erro"
End Sub
