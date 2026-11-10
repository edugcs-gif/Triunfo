# Triunfo

Script para atualização da tabela de cadastros imobiliários de Triunfo/RS a partir dos dados coletados em campo.

O `gerar_atualizada.py` cruza a aba **prefeitura** (cadastro atual) com a aba **bruta** (dados de campo) e grava o resultado na aba **atualizada** do mesmo arquivo `.xlsx`. As células alteradas ficam pintadas.

## Requisitos

- Python 3
- [openpyxl](https://pypi.org/project/openpyxl/)

```bash
pip install openpyxl
```

## Como rodar

1. Feche o `.xlsx` no Excel, porque o Windows não deixa sobrescrever um arquivo aberto.
2. Coloque o script na mesma pasta da planilha. O nome padrão da planilha é `tabela atualizacao.xlsx` e pode ser mudado em `ARQUIVO`.
3. Dê dois cliques no script ou rode:

```bash
python gerar_atualizada.py
```

Antes de salvar, o script faz uma cópia `.backup` da planilha.

## Abas da planilha

| Aba | Para que serve |
|---|---|
| `prefeitura` | Cadastro atual. O cabeçalho fica na linha 2 e cada linha é uma unidade de avaliação (cadastro + unidade). |
| `bruta` | Dados coletados em campo. |
| `depara` | Diz qual coluna da bruta atualiza qual coluna da prefeitura. As colunas são `planilha_bruta__unidades` (origem), `modelo` (destino) e `montar` (marque `agora` para usar o par). |
| `cancelados` | (opcional) Cadastros cancelados por unificação e o cadastro que fica. |
| `atualizada` | Saída, recriada a cada rodada. |
| `comparacao`, `unidades sem coleta`, `cadastros no mesmo lote` | Relatórios de conferência gerados pelo script. |

Para mapear uma coluna nova não é preciso mexer no código: basta editar a aba `depara`.

## Regras principais

- Célula da bruta preenchida e diferente da prefeitura: o valor é atualizado.
- Célula da bruta vazia ou com texto nulo (`nulo`, `null`, `-`...): fica o dado da prefeitura.
- Registro que só existe na bruta: entra como linha nova.
- Cadastros desativados (`17-Situação Cadastro` diferente de 1) são ignorados.
- **Edificação compartilhada** (mesmo `uuid_edificacao`):
  - 1 para 1: atualiza normalmente;
  - vários registros e algum tem matrícula: a área da prefeitura é mantida;
  - vários registros e nenhum tem matrícula: é criada uma linha **PRINCIPAL** com a área da edificação, e os registros de origem viram **COMPLEMENTAR**.
- **Um cadastro por lote**: se o lote tem mais de um cadastro, só um recebe o dado de campo. A escolha vai primeiro para o cadastro com matrícula e depois para o de maior área construída.
- **Lote baldio**: se a coleta diz que o lote é baldio e o cadastro ainda tem edificação, é criada uma linha de baldio e as unidades existentes não são tocadas.
- **Condomínio**: uma edificação com cadastros de 3 ou mais matrículas diferentes vira uma linha de condomínio, e os cadastros membros ficam intocados.
- Cadastros listados em `CADASTROS_BAIXADOS` ou na aba `cancelados` saem marcados para baixar.

Todas essas regras podem ser ligadas, desligadas ou ajustadas no bloco **CONFIGURAÇÃO**, no topo do script.

## Cores na aba `atualizada`

| Cor | Significado |
|---|---|
| Amarelo | Valor trocado pelo dado de campo |
| Cinza | A coleta confirmou o valor da prefeitura |
| Verde | Linha nova (registro que só existe na coleta) |
| Azul | Linha principal da edificação compartilhada |
| Laranja | Área mantida por causa da regra da edificação |
| Lilás | Linha nova de lote baldio |
| Verde-água | Condomínio (forte) e cadastros membros (claro) |
| Rosa | Unidade sem coleta em cadastro que teve outra unidade atualizada |
| Vermelho | Cadastro mal preenchido, que não recebe o dado de campo ou que deve ser baixado |

As colunas `_situacao` e `_status`, no fim da aba, permitem filtrar as linhas por tipo de alteração (ATUALIZADO, UNIDADE NOVA, CADASTRO NOVO, CADASTRO BAIXADO etc.).
