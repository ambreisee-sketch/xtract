# Planilhas de importação corrigidas

Geradas pelo `scripts/gerar_planilha_importacao.py` a partir das planilhas
"Cópia de PLANILHA SANTA ...". O formato é idêntico ao da "PLANILHA UNICA
CONSOLIDADA", que importa sem erro: tudo gravado como texto, datas dd/mm/aaaa,
sem células vazias ocultas e sem matrícula repetida.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `PILOTO_teste_32_matriculas_PADRAO_CONSOLIDADA.xlsx` | Importe primeiro. São 32 matrículas da planilha principal, incluindo os casos mais arriscados. |
| `Matriculas_faltantes_PADRAO_CONSOLIDADA.xlsx` | Planilha principal, com 1.381 matrículas que faltam no sistema. |
| `Matriculas_DIGITADAS_encontradas_nas_copias_PADRAO_CONSOLIDADA.xlsx` | 17 matrículas que o relatório chama de "digitadas". Confirme que não estão no sistema antes de importar. |
| `PENDENTES_conferencia_manual_NAO_IMPORTAR.xlsx` | 82 matrículas que precisam de decisão humana. A aba LEIA traz o motivo de cada uma. Não importe. |
| `Relatorio_ajustes.xlsx` | O que foi alterado, célula por célula, mais os avisos e as 1.369 matrículas sem nenhuma fonte. |

## Ordem sugerida

1. Importe o PILOTO e confira no sistema se as 32 matrículas entraram completas.
   Vale olhar com atenção a 6650, que tem atos numerados até 16, e uma matrícula com CPF repetido entre cônjuges.
2. Se o piloto entrar inteiro, importe a planilha principal. As 32 do piloto
   também estão nela, então tire-as antes ou confirme que o sistema ignora repetidas.
3. Não abra e salve os arquivos no Excel antes de importar. O Excel regrava o
   arquivo em outro formato.

## O que mudou em relação à planilha que falhou

- Datas que estavam como número (ex.: 31873) agora são texto (06/04/1987).
- Removidas cerca de 344 mil células vazias invisíveis.
- Matrículas repetidas com dados diferentes foram para PENDENTES.
- Protocolo com texto livre virou "NÃO CONSTA". O original está no relatório.
- NOTAS vazia virou "NÃO CONSTA", como na consolidada.
- Zeros à esquerda de CPF e CNPJ foram recolocados só quando o documento fica válido.
