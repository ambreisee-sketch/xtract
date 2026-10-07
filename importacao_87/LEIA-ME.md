# As 87 matrículas que não entraram na importação

A lista vem do CSV de faltantes de 06/10/2026: são as matrículas que estão em
alguma planilha e não estão como digitadas.

| Arquivo | Conteúdo |
|---|---|
| `Matriculas_85_PRONTAS_PADRAO_CONSOLIDADA.xlsx` | **Importar este.** 85 matrículas no padrão da consolidada. |
| `Matriculas_87_relatorio.xlsx` | Uma linha por matrícula: versão usada, confiança e o que foi feito. A aba Alteracoes traz cada valor alterado. |
| `resolucoes/` | A análise de cada uma das 29 matrículas resolvidas pelo conteúdo, com as provas. |
| `etapas/` | Arquivos intermediários. Não importar. |

Ficaram de fora duas matrículas:

- **2745:** a pedido.
- **3326:** é uma cópia exata da 3324, e o conteúdo verdadeiro dela não está em nenhuma planilha.

## Por que o sistema recusou essas matrículas

Comparamos as 3.592 matrículas que entraram com as 75 que falharam nas importações reais.
Quatro regras explicam todas as falhas, sem nenhuma exceção:

1. Dois atos da mesma matrícula com o mesmo número, por exemplo R-3 e AV-3.
2. Uma parte em ATO_PARTES cujo ID não corresponde a nenhum ato.
3. Um PROTOCOLO do ato que não é número de até 9 dígitos nem "NÃO CONSTA".
4. Uma DATA do ato que não existe no calendário, por exemplo 29/02/2011.

O script `scripts/validar_regras_importador.py` confere essas regras em qualquer planilha antes de importar.

## Como foi feito

- **58 matrículas, correção automática.** Protocolo inválido virou "NÃO CONSTA". Partes sem ato foram ligadas ao único ato de mesmo nome. Datas por extenso foram convertidas. A 16864 ganhou a data da nota de retificação. A 3576 estava gravada como 3516.
- **27 matrículas, resolvidas lendo o conteúdo.** Foram lidos o texto e o cabeçalho de cada ato, as datas, as notas de cancelamento e todas as versões da matrícula nas outras planilhas. Cada solução foi conferida por um verificador independente. São 14 com confiança média, listadas no relatório. Nelas a numeração saiu da ordem das datas, ou o cabeçalho da matrícula veio da nota de reabertura.
