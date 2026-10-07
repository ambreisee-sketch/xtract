# As 87 matrículas que não entraram na importação

A lista vem do CSV de faltantes de 06/10/2026: são as matrículas que estão em
alguma planilha e não estão como digitadas. Cada uma foi tirada da versão mais
recente: primeiro a "Planilha Única - 32 Faltantes", depois a "Matrículas
faltantes", a planilha gerada em importacao_corrigida e a consolidada.

| Arquivo | Conteúdo |
|---|---|
| `Matriculas_87_PADRAO_CONSOLIDADA.xlsx` | 58 matrículas prontas para importar. |
| `Matriculas_87_PADRAO_CONSOLIDADA_CONFERIR_NAO_IMPORTAR.xlsx` | 29 matrículas que precisam de conferência no livro. A aba LEIA traz o motivo de cada uma. |
| `Matriculas_87_PADRAO_CONSOLIDADA_relatorio.xlsx` | A versão usada de cada matrícula e cada valor alterado. |

## Por que o sistema recusou essas matrículas

Comparamos as 3.592 matrículas que entraram com as 75 que falharam nas importações reais.
Quatro regras explicam todas as falhas, sem nenhuma exceção:

1. Dois atos da mesma matrícula com o mesmo número, por exemplo R-3 e AV-3.
2. Uma parte em ATO_PARTES cujo ID não corresponde a nenhum ato.
3. Um PROTOCOLO do ato que não é número de até 9 dígitos nem "NÃO CONSTA", por exemplo "1549/80" ou "Ofício nº 583/92".
4. Uma DATA do ato que não existe no calendário, por exemplo 29/02/2011.

O script `scripts/validar_regras_importador.py` confere essas regras em qualquer planilha antes de importar.

## O que foi corrigido nas 58

- **Protocolo inválido:** virou "NÃO CONSTA" em 42 matrículas. O texto original continua no ato.
- **Parte sem ato:** foi ligada ao único ato de mesmo nome em 4 matrículas.
- **Data por extenso:** a DATA ABERTURA virou dd/mm/aaaa nessas mesmas 4.
- **16864:** a data do R-1 virou 29/02/2012, como manda a nota de retificação do próprio ato.
- **3576:** estava digitada com o número 3516 na "Matrículas faltantes". Entrou com o número certo.
  Confira no sistema se a matrícula 3516 recebeu por engano os dados da 3576.
