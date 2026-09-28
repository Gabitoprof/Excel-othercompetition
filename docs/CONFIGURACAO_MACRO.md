# Configuração dos botões de importação (passo único)

O arquivo `Painel_Produtividade_AGHU.xlsm` já vem pronto com:

- Dois botões na aba **PAINEL**, seção "9. IMPORTAR DADOS DO AGHU": *Selecionar
  arquivo do Período 1* e *Selecionar arquivo do Período 2*.
- Os botões já estão associados aos nomes de macro `ImportadorAGHU.ImportarPeriodo1`
  e `ImportadorAGHU.ImportarPeriodo2`.
- Um terceiro botão na aba **Categorias**: *Atualizar Lista de Categorias*, já
  associado ao nome de macro `CategoriasAGHU.SincronizarCategorias`.

O que falta, **uma única vez**, é carregar o código dessas macros dentro do
arquivo. Isso não pode ser feito de fora do Excel (não é possível gravar um
projeto VBA compilado em um ambiente sem o Excel instalado), então quem for
configurar o arquivo (a própria unidade de saúde, ou o suporte de TI) precisa
fazer isto no Windows, com o Excel aberto:

## Passo a passo (leva menos de 2 minutos)

1. Abra `Painel_Produtividade_AGHU.xlsm` no Excel.
2. Se aparecer uma faixa amarela dizendo "Macros foram desabilitadas", **não
   clique em Habilitar Conteúdo ainda** — primeiro finalize os passos abaixo.
3. Pressione `Alt+F11` para abrir o Editor do VBA.
4. No painel à esquerda ("Project - VBAProject"), clique com o botão direito
   sobre o nome do arquivo (ex.: `VBAProject (Painel_Produtividade_AGHU.xlsm)`)
   → **Import File...**
5. Selecione o arquivo `macros/ImportadorAGHU.bas` (está na mesma pasta deste
   guia, dentro do repositório/pasta entregue).
6. Repita os passos 4 e 5 selecionando agora `macros/CategoriasAGHU.bas` — os
   dois módulos precisam ser importados (aparecem como dois itens separados em
   "Modules" no painel do projeto).
7. Feche o Editor do VBA (`Alt+Q` ou botão fechar).
8. Salve o arquivo normalmente (`Ctrl+S`), mantendo o formato **Excel
   Habilitado para Macro (*.xlsm)**.
9. Feche e reabra o arquivo. Agora, ao abrir, clique em **Habilitar
   Conteúdo** na faixa amarela — os três botões (dois na aba PAINEL, um na
   aba Categorias) já vão funcionar normalmente a partir daí, todo mês, sem
   repetir esse processo.

## Se os botões não funcionarem mesmo depois de importar o módulo

Isso quase sempre significa que a **política de TI do órgão bloqueia macros**
neste computador (bastante comum em redes de secretarias de saúde estaduais,
como a SESAB). Nesse caso:

- Peça ao suporte técnico para liberar macros para este arquivo específico,
  ou para a pasta onde ele fica salvo (em geral: Central de Confiabilidade
  do Excel → Locais Confiáveis, ou uma política de GPO que precisa ser
  ajustada pelo administrador de TI).
- Se não for possível liberar, use o método alternativo sem macro descrito
  na aba **Instruções**, seção 6 ("Importação automática com Power Query —
  alternativa sem macro"). Ele exige mais passos de configuração inicial,
  mas depois de configurado uma vez, a atualização mensal também é simples
  (substituir dois arquivos de mesmo nome e clicar em "Atualizar Tudo").

## Por que este passo não vem pronto de fábrica?

Um arquivo `.xlsm` com macro é, tecnicamente, um contêiner binário compilado
pelo próprio Excel. Não existe uma forma confiável de gerar esse binário
fora do Excel real — qualquer ferramenta que tente fazer isso corre o risco
de gerar um arquivo corrompido, que o Excel recusaria abrir ou "repararia"
apagando as macros na primeira tentativa de abertura. Por isso a estrutura
inteira da planilha (abas, fórmulas, formatação, os três botões já
posicionados e já referenciando os nomes de macro certos) veio pronta, e
apenas a importação do código-fonte (dois arquivos `.bas`, texto simples)
precisa ser feita dentro do Excel, uma única vez.
