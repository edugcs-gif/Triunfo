# gerar_atualizada.py

Cruza o cadastro imobiliário da prefeitura de Triunfo/RS com a coleta de campo da Solo e grava, no mesmo arquivo Excel, a aba `atualizada` com as células alteradas pintadas.

## Como rodar

1. Instale o Python 3 e a biblioteca: `pip install openpyxl`.
2. Coloque o script na mesma pasta de `tabela atualizacao.xlsx` e feche o arquivo no Excel.
3. Dê dois cliques no script ou rode `python gerar_atualizada.py`.

A rodada leva de 2 a 4 minutos. Antes de salvar, o script grava uma cópia `.backup` do arquivo.

## Entrada

| Aba | Conteúdo |
|---|---|
| `prefeitura` | Cadastro oficial, uma linha por unidade de avaliação. Cabeçalho na linha 2. |
| `bruta` | Coleta de campo, uma linha por edificação. Cabeçalho na linha 1. |
| `depara` | Qual coluna da coleta (`planilha_bruta__unidades`) atualiza qual coluna da prefeitura (`modelo`). Só entram as linhas com `montar` = "agora". |
| `cancelados` (opcional) | Cadastro cancelado na primeira coluna; na `Observação`, o cadastro que fica ou "baixado". |

## Saída

Quatro abas, recriadas a cada rodada:

| Aba | Conteúdo |
|---|---|
| `atualizada` | Os cadastros tocados pela coleta ou por alguma regra, na estrutura da prefeitura, seguidos das colunas da Solo e das colunas de controle. |
| `comparacao` | Resumo da rodada e comparação de áreas por cadastro (lote, matrícula e área construída). |
| `unidades sem coleta` | Unidades sem dado de campo em cadastro que teve outra unidade atualizada. |
| `cadastros no mesmo lote` | Lotes com mais de um cadastro. |

### Como ler a aba atualizada

A coluna `_situacao` classifica cada linha; `_situacao_historico` traz o detalhe e `_status` a explicação.

| `_situacao` | Significado |
|---|---|
| ATUALIZADO | Linha da prefeitura que recebeu ou confirmou o dado de campo (inclui complementar e condomínio). |
| UNIDADE NOVA | Unidade que não existia no cadastro (inclui linha principal e lote baldio). |
| CADASTRO NOVO | Número que não existe na prefeitura, ou coleta sem número. |
| UNIDADE BAIXADA | Unidade a baixar: lista manual, lote que virou baldio ou unidade sem coleta. |
| CADASTRO BAIXADO | Cadastro cancelado, ou que ficou sem atualizar porque outro cadastro do lote recebeu a coleta. |
| CONFERIR | `Edificacoes_num_cadastro` mal preenchido; a linha não foi vinculada. |

Cores: amarelo = alterado; cinza = confirmado em campo; laranja = mantido por regra; verde = linha nova; azul = principal; lilás = baldio novo; verde-água = condomínio; vermelho = baixa, não atualizar ou conferir.

## Regras principais

- **Vínculo**: primeiro por `Edificacoes_num_cadastro` (`cadastro.unidade`), depois por `Lotes_num_cadastr` + `unidade`. Sem vínculo, a linha entra como unidade nova ou cadastro novo.
- **Atualização**: valor diferente troca o da prefeitura; coleta vazia mantém. O número do cadastro e o da unidade nunca são trocados.
- **Cadastro desativado** (`17-Situação Cadastro` diferente de 1): fica fora do processo.
- **Edificação dividida** por mais de um cadastro ou unidade: a área da unidade (`140`) não é atualizada e entra uma linha PRINCIPAL com a área coletada.
- **Mais de um cadastro** na mesma célula ou no mesmo lote: fica um só, o que tem matrícula e, no empate, o de maior área construída.
- **Condomínio**: edificação com cadastros de 3 ou mais matrículas diferentes ganha uma linha própria; os membros mantêm área, proprietário e CPF da prefeitura.
- **Lote baldio** (`Lotes_cd_ocupacao` = 4): as unidades do cadastro são marcadas para baixa e entra uma linha de terreno.
- **Cancelados**: o cancelado é baixado e o cadastro que fica recebe a matrícula e a área de matrícula dele. Com "baixado" na `Observação`, só baixa.
- **Coleta em cadastro baixado ou cancelado**: é descartada e listada no relatório da rodada.
- **Limpeza**: "nulo", "null", "-" e similares contam como vazio. No `156-Responsável` também ".", "0" e "sem informação". O `157-CPF/CNPJ Responsável` fica só com os dígitos.

## Configuração

Tudo fica no topo do script, em constantes em maiúsculas. As mais usadas:

| Constante | Para que serve |
|---|---|
| `ARQUIVO` | Nome da planilha. Vazio = usa o único .xlsx da pasta. |
| `CADASTROS_BAIXADOS` | Lista manual de cadastros a baixar. |
| `TEXTOS_NULOS`, `TEXTOS_NULOS_POR_COLUNA` | Textos da coleta que contam como vazio. |
| `COLUNAS_SO_DIGITOS` | Colunas de documento que ficam só com os dígitos. |
| `CONDOMINIO_MIN_MATRICULAS` | Mínimo de matrículas diferentes para virar condomínio. |
| `COMPARACOES` | Pares de colunas e limites de alerta da aba `comparacao`. |

Para mapear uma coluna nova não é preciso mexer no código: basta incluir a linha na aba `depara`.
