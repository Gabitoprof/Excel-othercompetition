# Excel-othercompetition

Productivity panel in Excel for health units. Consolidates two exported files by
competence (in any period), calculates indicators by procedure and by service,
identifies services with multiple procedures, and allows automatic import of data.

## Arquivos

- `Painel_Produtividade_AGHU.xlsm` — planilha principal. Contém dois botões na
  aba PAINEL ("Selecionar arquivo do Período 1/2") para importar diretamente os
  arquivos exportados do AGHU, sem copiar e colar e sem Power Query.
- `macros/ImportadorAGHU.bas` — código-fonte das macros usadas pelos botões.
- `docs/CONFIGURACAO_MACRO.md` — passo a passo (único, ~2 minutos) para ativar
  os botões no Excel, e o que fazer caso as macros estejam bloqueadas por
  política de TI (com alternativa via Power Query, sem macro).

A aba **Instruções**, dentro da própria planilha, documenta o uso completo do
painel, incluindo as três formas de atualizar os dados a cada competência.
