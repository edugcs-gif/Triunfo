# -*- coding: utf-8 -*-
"""
gerar_atualizada.py
-------------------
Cruza a aba "prefeitura" com a aba "bruta" e grava a aba "atualizada" no mesmo
arquivo, com as celulas alteradas pintadas.

Quais colunas da bruta atualizam quais colunas da prefeitura vem da aba "depara"
(colunas: planilha_bruta__unidades = origem, modelo = destino, montar = agora).
Nao e preciso mexer no codigo para mapear coluna nova: basta editar o depara.

Regras basicas:
  - celula da bruta preenchida e diferente da prefeitura -> atualiza e PINTA de amarelo
  - celula da bruta vazia                                -> mantem o dado da prefeitura
  - registro que so existe na bruta                      -> entra como linha nova verde

Regra da edificacao compartilhada (uuid_edificacao):
  Conta quantos REGISTROS (cadastro + unidade) caem na mesma edificacao.
  - 1 para 1 ................................ atualiza normalmente
  - varios registros e ALGUM tem matricula .. mantem a area da prefeitura (celula
                                              laranja: o dado de campo foi barrado)
  - varios registros e NENHUM tem matricula . cria uma linha PRINCIPAL (azul) com a
                                              area da edificacao; os registros de
                                              origem viram COMPLEMENTAR e ficam com
                                              a area que ja tinham na prefeitura

Como rodar:
  1) feche o .xlsx no Excel (o Windows nao deixa sobrescrever arquivo aberto)
  2) coloque este script na mesma pasta do arquivo
  3) de dois cliques nele, ou rode:  python gerar_atualizada.py
"""

import math
import os
import re
import shutil
import sys
import unicodedata
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Falta a biblioteca openpyxl. Instale com:  pip install openpyxl")
    input("\nEnter para fechar...")
    sys.exit(1)


# ===========================================================================
#  CONFIGURACAO  -- so isso aqui costuma precisar de ajuste
# ===========================================================================

ARQUIVO = "tabela atualizacao.xlsx"   # "" = procura o unico .xlsx da pasta

ABA_PREFEITURA = "prefeitura"
ABA_BRUTA      = "bruta"
ABA_DEPARA     = "depara"
ABA_SAIDA      = "atualizada"         # e recriada a cada rodada
# A aba de saida traz so quem teve interacao com a coleta; o resto da prefeitura
# fica de fora. "cadastro" = mantem todas as unidades do cadastro que teve
# alguma interacao; "linha" = so as linhas que interagiram; None = tudo.
SO_COM_INTERACAO = "cadastro"
# Lista manual de cadastros que a prefeitura vai BAIXAR: todas as linhas deles
# entram na saida em vermelho, com status de baixado, e NAO recebem o dado de
# campo (a coleta que apontava para eles e listada no relatorio).
# Coleta com VARIOS numeros de cadastro na mesma celula ("54400 / 598800"):
# fica so um deles (mesmo criterio do lote: matricula, depois maior area
# construida) e os outros que existem na prefeitura sao marcados para BAIXAR.
# Cadastro desativado: linhas da prefeitura com 17-Situacao Cadastro diferente
# de 1 sao cadastros antigos/desativados. Elas sao IGNORADAS: nao entram na
# saida, nao recebem dado de campo e nao podem ser escolhidas em nenhuma regra.
COLUNA_SITUACAO_CADASTRO = "17-Situação Cadastro"
SITUACAO_CADASTRO_ATIVA  = ["1"]
# Edificacao dividida por mais de um cadastro: nao da para dizer quanto da area
# e de cada um, entao a area da unidade (140) NAO e atualizada em nenhum deles,
# venha isso de uma celula que lista varios cadastro.unidade ou de varias linhas
# da coleta com o mesmo uuid_edificacao.
EDIFICACAO_MULTICADASTRO_TRAVA_AREA = True
# True = alem de travar a area, cria uma linha PRINCIPAL (azul) para essa
# edificacao, com todos os numeros de cadastro juntos no lugar do numero.
PRINCIPAL_EDIF_MULTICADASTRO = True
MULTI_CADASTRO_ESCOLHE = True
# O que fazer com os outros cadastros da celula:
#   "principal"  -> o escolhido vira o cadastro PRINCIPAL e os outros ficam como
#                   COMPLEMENTARES (nao sao atualizados nem baixados); na aba de
#                   comparacao a area da prefeitura do principal e a SOMA das
#                   areas de todos os cadastros do grupo
#   "baixar"     -> os outros sao marcados para BAIXAR
MULTI_CADASTRO_MODO = "principal"
# Colunas da prefeitura que, no cadastro PRINCIPAL, sao comparadas pela soma de
# todos os cadastros do grupo. So faz sentido para area construida: a area do
# lote e a da matricula costumam vir repetidas em todos os cadastros.
MULTI_SOMA_COMPARACAO = ["149-*Área Constr - TOTAL"]
# (598800 saiu da lista em 10/11/2026: na aba cancelados ele e o cadastro que
#  FICA na unificacao do 54400)
CADASTROS_BAIXADOS = [
    "115100", "148100", "620200", "324100", "132600",
    "475700", "161500", "68801", "300800",
    "661400", "661500", "675300",
]
# Coleta que aponta para uma unidade que NAO existe na prefeitura, em cadastro
# baixado ou cancelado: True = e descartada e listada no relatorio, igual a
# coleta das unidades que existem; False = entra como unidade nova (verde)
# dentro do cadastro baixado.
DESCARTAR_COLETA_NOVA_EM_BAIXADO = True
ABA_COMPARACAO = "comparacao"         # resumo + comparacao de areas; "" desliga
# Unidades que NAO receberam dado de campo, em cadastros onde alguma outra
# unidade recebeu. Serve para achar erro de preenchimento na bruta. "" desliga.
ABA_PENDENCIAS = "unidades sem coleta"
# Cadastros diferentes no MESMO lote (uuid_lote) — util para achar cadastro
# duplicado ou matricula trocada. "" desliga.
ABA_LOTES = "cadastros no mesmo lote"
# True = lista so os grupos com matricula diferente; False = lista todos os
# lotes que tem mais de um cadastro, com a situacao de cada um.
LOTES_SO_MATRICULA_DIFERENTE = False
# Quando uma linha da coleta aponta para cadastros DIFERENTES (dois cadastros no
# mesmo lote), o dado de campo vai para um so: o que tem matricula; se todos
# tiverem (ou nenhum), o de maior area construida. Os outros ficam intocados.
# Lote (uuid_lote) com mais de um cadastro: so um recebe o dado de campo — o que
# tem matricula; se todos tiverem (ou nenhum), o de maior area construida. Os
# outros cadastros do lote ficam INTOCADOS e em vermelho, e as linhas da coleta
# que apontavam para eles viram unidade nova do cadastro escolhido.
UM_CADASTRO_POR_LOTE = True
COLUNA_UUID_LOTE = "uuid_lote"
COLUNA_AREA_CONSTRUIDA_CADASTRO = "149-*Área Constr - TOTAL"
# CONDOMINIO: edificacao com 2+ cadastros de matriculas DIFERENTES. Os
# cadastros que formam o condominio ficam intocados e marcados; entra uma linha
# de condominio com todos os dados da coleta e o historico dos cadastros. Outra
# edificacao do mesmo lote vira unidade nova do condominio.
REGRA_CONDOMINIO = True
# Varias linhas da coleta na MESMA edificacao do condominio: False = fica so a
# linha do condominio (uma por condominio); True = cria tambem uma "unidade do
# condominio" para cada linha da coleta.
CONDOMINIO_UNIDADES_MESMA_EDIF = False
# Minimo de matriculas DIFERENTES (da prefeitura) para virar condominio. Com
# menos, a edificacao segue as regras normais (lote / edificacao compartilhada).
CONDOMINIO_MIN_MATRICULAS = 3
TEXTO_CONDOMINIO = "CONDOMINIO"
# Os cadastros membros recebem os dados de lote e as classificacoes da
# edificacao que vieram da coleta, mas estas colunas de area ficam com o valor
# da prefeitura.
COLUNAS_AREA_TRAVADAS_CONDOMINIO = ["140-*Área Constr Unidade"]
# Nos cadastros membros de condominio, o proprietario tambem fica o da
# prefeitura: a coleta registra um nome so para o predio todo, que nao e o dono
# de cada cadastro. O documento vai junto para nao misturar nome e CPF.
COLUNAS_PROPRIETARIO_TRAVADAS_CONDOMINIO = ["6-Proprietário", "15-Código do CNPJ/CPF"]
# Numero da unidade nas linhas transferidas: "proxima" (continua a numeracao do
# cadastro escolhido), "coleta" (o numero que veio da coleta) ou "" (em branco).
UNIDADE_TRANSFERIDA = "proxima"

# Pares (coluna da bruta, coluna da prefeitura) comparados na aba de comparacao.
# Sao independentes do depara: comparar nao atualiza nada.
# Terceiro item (opcional) = limite de alerta desse par. Sem ele vale
# DESTACAR_DIF_ACIMA. O alerta pinta a celula e escreve "ACIMA DE x%" numa
# coluna propria, que da para filtrar.
# (coluna da bruta, coluna da prefeitura, limite de alerta, nome amigavel do
#  lado prefeitura, nome amigavel do lado coleta). Os tres ultimos sao opcionais.
COMPARACOES = [
    ("Lotes_area_lote",  "28-*Área do Cadastro",      0.02,
     "area do cadastro",      "area do lote coletada"),
    ("Lotes_area_lote",  "118-Área do Imóvel(Matr)",  0.10,
     "area da matricula",     "area do lote coletada"),
    ("Lotes_area_total", "149-*Área Constr - TOTAL",  0.02,
     "area construida atual", "area construida coletada"),
    # "soma:<coluna>" = soma dessa coluna em todas as edificacoes do cadastro
    # (cada uuid_edificacao conta uma vez)
    ("soma:Edificacoes_sem_beiral", "149-*Área Constr - TOTAL", 0.02,
     "area construida atual", "soma edificação sem beiral"),
]
# Cor da faixa de cada par no cabecalho da aba de comparacao (cicla).
CORES_COMPARACAO = ["9FC5E8", "B6D7A8", "FFE599", "D9D2E9", "F9CB9C"]
NOME_DIF_M2    = "diferença em metros"
NOME_DIF_PC    = "diferença em %"
NOME_ALERTA    = "alerta"
# "cadastro" = uma linha por cadastro (informacao de lote inteiro, que se repete
# em todas as unidades). "unidade" = uma linha por cadastro+unidade.
COMPARACAO_POR = "cadastro"
DESTACAR_DIF_ACIMA = 0.10             # pinta a diferenca acima de 10%

# Em que linha esta o cabecalho de cada aba
LINHA_CAB_PREFEITURA = 2              # linha 1 tem as anotacoes, dados a partir da 3
LINHA_CAB_BRUTA      = 1
LINHA_CAB_DEPARA     = 1

# --- chave de ligacao -------------------------------------------------------
COLUNA_CHAVE_PREFEITURA = "1-Número do Cadastro"
COLUNA_CHAVE_BRUTA      = "Lotes_num_cadastr"

# Cascata de vinculo: primeiro a coluna abaixo (formato cadastro.unidade, podendo
# listar varios separados por " / "), depois COLUNA_CHAVE_BRUTA + unidade, e por
# ultimo a linha entra como cadastro novo.
USAR_CADASTRO_EDIFICACAO   = True
COLUNA_CADASTRO_EDIFICACAO = "Edificacoes_num_cadastro"
TEXTO_CADASTRO_NOVO        = "CADASTRO NOVO"
COLUNA_CHAVE_COLETA        = "_chave_coleta"   # de onde veio o vinculo
# Colunas que identificam a linha: o depara NUNCA sobrescreve isso numa linha que
# ja existe na prefeitura (senao o vinculo por cadastro.unidade acabaria trocando
# o proprio numero do cadastro ou da unidade pelo que veio da coleta).
PROTEGER_COLUNAS_CHAVE = True

# Quando a coluna de origem do depara vem vazia, tenta estas outras colunas da
# bruta antes de desistir (as duas areas abaixo sao identicas quando as duas vem
# preenchidas; em 70 linhas so a segunda tem valor).
FALLBACK_COLUNAS_BRUTA = {
    "area_edificacao": ["Edificacoes_area_total"],
    # nome do proprietario: se o do lote vier limpo, tenta o contribuinte
    # responsavel; se os dois vierem limpos, fica o da prefeitura
    "Lotes_proprietario_lote": ["contri_resp"],
    # inscricao: a da edificacao; se vier em branco, a do lote
    "Edificacoes_inscricao_edificacao": ["Lotes_inscricao_imobiliaria"],
    # area da unidade: o sem beiral; se nao vier, a area total da edificacao
    "Edificacoes_sem_beiral": ["Edificacoes_area_total"],
}

# Arredondar os numeros que vem da coleta (as areas do QGIS chegam com 13 casas:
# 489.4181294608861). None = grava como veio; 2 = duas casas decimais.
# So mexe em valor numerico nao inteiro, entao codigo e numero de cadastro nao
# sao afetados.
ARREDONDAR_CASAS = None

# O valor que vem da coleta e sempre texto. Com True, quando a coluna de destino
# e numerica na prefeitura, o texto e convertido em numero antes de gravar. Sem
# isso o Excel ordena o cadastro novo depois de todos os numeros, porque esta
# comparando texto com numero. Nao converte quem tem zero a esquerda ("053").
AJUSTAR_TIPO_PELA_PREFEITURA = True

# --- lote baldio ------------------------------------------------------------
# Quando a coleta diz que hoje o lote e baldio, o cadastro da prefeitura pode
# ainda ter unidades com edificacao (o predio foi demolido). Nesses casos as
# unidades que ja existem NAO sao tocadas e o script cria uma linha nova com o
# mesmo numero de cadastro, com os dados de lote da coleta e os campos de
# edificacao zerados.
REGRA_LOTE_BALDIO      = True
COLUNA_OCUPACAO_BRUTA  = "Lotes_cd_ocupacao"
VALOR_OCUPACAO_BALDIO  = ["4"]
UNIDADE_LOTE_BALDIO    = 1          # numero da unidade da linha nova
# Cadastro que JA e terreno na prefeitura (nenhuma unidade de avaliacao): nao
# adianta criar linha nova, entao a coleta atualiza a linha que ja existe.
BALDIO_ATUALIZA_TERRENO = True
# Cadastro que na prefeitura so tem a linha de terreno (unidade vazia/NULL) e a
# coleta achou edificacao (ex.: 95200.1): a coleta atualiza essa linha e ela
# ganha o numero da unidade da coleta (era baldio e foi construido).
TERRENO_RECEBE_UNIDADE_DA_COLETA = True
# Colunas que a linha de baldio herda do cadastro que ja existe na prefeitura,
# alem das de COLUNAS_HERDADAS_DO_CADASTRO. Sao dados do lote/cadastro que a
# coleta nao traz e que nao mudam por o terreno estar vazio.
COLUNAS_HERDADAS_NO_BALDIO = ["118-Área do Imóvel(Matr)"]
COLUNAS_NULAS_BALDIO = [
    "140-*Área Constr Unidade", "141-*Ano da Construção",
    "142-*Espécie da Construção", "143-*Tipo de Edificação",
    "139-Tipo Construtivo", "146-VlrVenal ConstUnidad", "147-Vlr M² ConsUnidade",
    "150-Carta Habite-se nº", "151-Alvará de Demolição",
    "152-Data Alvará Demolição", "181-## <= 2024 ##", "182-Tipo Construtivo",
    "183-Padrão Imóvel", "144-Área Demolida Unida", "149-*Área Constr - TOTAL",
]

# --- limpeza dos valores que vem da coleta ----------------------------------
# Textos que a coleta usa no lugar de "vazio". Valem para qualquer coluna: onde
# aparecem, o script trata como nao coletado e mantem o dado da prefeitura.
TEXTOS_NULOS = ["nulo", "nula", "null", "n/a", "-", "--"]
# Textos que so contam como vazio nestas colunas de destino.
TEXTOS_NULOS_POR_COLUNA = {
    "6-Proprietário": ["sem informacao", "0"],
    "156-Responsável": [".", "..", "...", "0", "00", "n", "x", "?",
                        "sem informacao", "nao informado"],
}
# Colunas de documento: a coleta traz "123.456.789-00" e a prefeitura guarda so
# os numeros. Nestas colunas o valor da coleta fica so com os digitos; sem
# nenhum digito conta como vazio.
COLUNAS_SO_DIGITOS = ["157-CPF/CNPJ Responsável"]
# Numa LINHA NOVA de um cadastro que JA EXISTE na prefeitura, estas colunas, se
# ficarem vazias, sao herdadas do proprio cadastro (primeiro valor preenchido
# entre as unidades que ja estao na prefeitura). E o que faz a unidade nova sair
# com o mesmo proprietario das outras unidades do cadastro.
COLUNAS_HERDADAS_DO_CADASTRO = ["6-Proprietário"]
# Texto a gravar nas LINHAS NOVAS (verdes) quando a coluna ficaria vazia mesmo
# depois da coleta e da heranca acima.
TEXTO_PADRAO_LINHA_NOVA = {
    "6-Proprietário": "PROPRIETARIO NAO LOCALIZADO",
}

# Cada linha da prefeitura e uma UNIDADE DE AVALIACAO: o mesmo cadastro aparece
# varias vezes. Com True a ligacao e (cadastro + unidade), que e o par unico.
CHAVE_COM_UNIDADE = True
COLUNA_UNIDADE_PREFEITURA = "145-Nº Unidade Avaliação"
COLUNA_UNIDADE_BRUTA      = "unidade"

# --- regra da edificacao compartilhada --------------------------------------
REGRA_EDIFICACAO = True
COLUNA_UUID_EDIFICACAO = "uuid_edificacao"   # coluna da bruta que agrupa
COLUNA_MATRICULA       = "502-Nº Matrícula"  # coluna da PREFEITURA que decide
# Coluna de area que a regra controla (nao e atualizada nos casos travados)
COLUNAS_AREA_CONTROLADAS = ["140-*Área Constr Unidade"]
# De onde sai a area da linha PRINCIPAL (coluna da bruta) e onde ela e gravada
ORIGEM_AREA_PRINCIPAL  = "Edificacoes_sem_beiral"
DESTINO_AREA_PRINCIPAL = "140-*Área Constr Unidade"
# Grupo misto (um cadastro com matricula e outro sem):
#   "com_matricula" = trata o grupo todo como se tivesse matricula (nao cria principal)
#   "sem_matricula" = cria a linha principal assim mesmo
GRUPO_MISTO = "com_matricula"
# True = a linha principal recebe todos os campos do depara (do 1o registro do
# grupo), com a mesma limpeza e cascata das outras linhas, mais a heranca do
# cadastro. False = so a area, o uuid da edificacao e as marcacoes.
COPIAR_DADOS_NO_PRINCIPAL = True
# Quando o grupo e todo do MESMO cadastro (varias unidades na mesma edificacao),
# a linha principal ja nasce com esse numero de cadastro. Com varios cadastros
# diferentes ela fica sem numero, para a prefeitura atribuir.
PRINCIPAL_HERDA_CADASTRO_UNICO = True
# Na linha COMPLEMENTAR a coluna de origem mostra so o codigo da propria linha
# (cadastro/unidade). A lista com todos os codigos do grupo fica so na PRINCIPAL.
COMPLEMENTAR_MOSTRA_CODIGO = True
COLUNA_VINCULO = "_vinculo_edificacao"   # principal / complementar
COLUNA_ORIGEM  = "_cadastros_origem"     # registros (cadastro/unidade) do grupo
# O grupo e formado por REGISTROS (cadastro+unidade) na mesma edificacao. Duas
# unidades do MESMO cadastro na mesma edificacao ja contam como grupo.

# --- aba depara -------------------------------------------------------------
DEPARA_COL_ORIGEM  = ["planilha_bruta__unidades", "bruta", "origem", "de"]
DEPARA_COL_DESTINO = ["modelo", "prefeitura", "destino", "para"]
DEPARA_COL_ATIVO   = ["montar", "ativo", "usar", "processar"]      # opcional
DEPARA_VALOR_ATIVO = ["agora", "sim", "x", "s", "1", "true", "ok"]
DEPARA_SO_MARCADOS = True     # True = so processa as linhas marcadas 'agora'

# Coluna de destino que existe no depara mas nao na prefeitura
# (uuid, uuid_lote, uuid_edificacao): cria no fim. "inicio" cria antes da
# primeira coluna; None ignora essas linhas do depara.
POSICAO_COLUNAS_NOVAS = "fim"          # "fim" | "inicio" | None
# Coluna vazia que marca onde acabam as colunas da prefeitura e comecam as
# nossas (do uuid em diante). "" desliga.
COLUNA_SEPARADOR = "COLUNAS DA SOLO"
# Pares extras (coluna da bruta, nome da coluna na saida) que entram como se
# estivessem no depara marcados "agora". Se o mesmo par ja estiver no depara,
# nao duplica.
# Troca a coluna de origem do depara, sem precisar mexer na aba: destino -> nova
# origem na bruta. Vale tambem para as colunas da Solo criadas pelo script.
# (vazio: a troca do 140 ja esta feita na aba depara da base)
DEPARA_TROCAR_ORIGEM = {}
# Renomeia a coluna de destino na saida (nome antigo -> nome novo).
RENOMEAR_DESTINO = {
    "Edificacoes_area_total": "Área total da edificação",
}
# True = os pares de COLUNAS_SOLO_EXTRAS que ainda nao estao na aba depara sao
# acrescentados nela ao salvar, para a aba mostrar tudo o que o script usou.
ESCREVER_EXTRAS_NO_DEPARA = True
# Nomes alternativos de colunas da bruta: quando a coluna do depara nao existe,
# o script tenta esses apelidos (entregas mudam a grafia de um setor para outro)
APELIDOS_COLUNAS_BRUTA = {
    "Edificacoes_pavimentos": ["Edificacoes_pavimetos"],
    "Edificacoes_pavimetos":  ["Edificacoes_pavimentos"],
}
# Documento do responsavel: quando a coleta traz o nome do responsavel mas nao
# o documento, e esse nome e o mesmo do proprietario da linha, o documento do
# responsavel passa a ser o do proprietario. (resp. na bruta, doc na bruta,
# responsavel, doc do responsavel, proprietario, doc do proprietario)
DOC_RESPONSAVEL_DO_PROPRIETARIO = ("contri_resp", "doc_propri_unid",
                                   "156-Responsável", "157-CPF/CNPJ Responsável",
                                   "6-Proprietário", "15-Código do CNPJ/CPF")
# Campos preenchidos SO nos cadastros novos (numero que nao existe na
# prefeitura, ou linha sem numero): (coluna da prefeitura, coluna da bruta ou
# None, valor fixo quando nao vem da bruta).
DEPARA_CADASTRO_NOVO = [
    ("117-Cód. do Bairro Corr", "Lotes_cd_bairro",  None),
    ("116-Cód. Logradouro",     "Lotes_cod_lograd", None),
    ("13-Número do imóvel",     "Lotes_num_imovel", None),
    ("19-Cidade",               None,               "Triunfo"),
    ("21-Sigla UF",             None,               "RS"),
    ("INNATUREZAJURIDICARESP",  "cd_tip_contrib",   None),
    ("12-Endereço",             "Lotes_logradouro", None),
    ("18-Descrição do Bairro",  "Lotes_bairro",     None),
    ("20-Código do CEP",        None,               95840000),
]
# Aba de cadastros cancelados por unificacao (opcional): a 1a coluna tem o
# cadastro cancelado e COLUNA_CANCELADO_DESTINO o cadastro que fica. O cancelado
# e baixado; o que fica recebe as matriculas dele (concatenadas) e a soma da
# area da matricula, contando cada cadastro cancelado uma vez so.
ABA_CANCELADOS = "cancelados"
COLUNA_CANCELADO_DESTINO = "Observação"
# Textos em COLUNA_CANCELADO_DESTINO que querem dizer "cancelado sem unificacao":
# o cadastro e baixado e nao passa matricula nem area para nenhum outro.
CANCELADO_SEM_DESTINO = ["baixado", "baixada", "cancelado", "cancelada"]
COLUNA_AREA_MATRICULA = "118-Área do Imóvel(Matr)"
COLUNAS_SOLO_EXTRAS = [
    ("Edificacoes_inscricao_edificacao", "Inscrição da unidade"),
    # 1 = Matricula em Condominio, 2 = Pro diviso, 3 = Indiviso
    ("Lotes_cd_tip_condominio", "Tipo de condomínio"),
    ("Lotes_num_setor",  "102-* Setor Fiscal"),
    ("Lotes_num_quadra", "103-* Quadra"),
    ("contri_resp",      "156-Responsável"),
    ("doc_propri_unid",  "157-CPF/CNPJ Responsável"),
]
COR_SEPARADOR    = "BFBFBF"

# --- cores ------------------------------------------------------------------
COR_ATUALIZADO = "FFF2CC"     # amarelo  - valor trocado pelo dado de campo
COR_LINHA_NOVA = "D9EAD3"     # verde    - registro que so existe na bruta
COR_PRINCIPAL  = "CFE2F3"     # azul     - linha principal criada pela regra
COR_MANTIDA    = "FCE5CD"     # laranja  - area mantida por causa da regra
COR_ERRO       = "F4CCCC"     # vermelho - cadastro da edificacao mal preenchido
COR_IGUAL      = "E8E8E8"     # cinza    - a coleta trouxe o mesmo valor da prefeitura
COR_BALDIO     = "D9D2E9"     # lilas    - linha nova de lote baldio
COR_BAIXAR     = "EA9999"     # vermelho forte - unidade que a prefeitura tem que
                              # baixar: hoje o lote e baldio
COR_SEM_COLETA = "FFC7CE"     # rosa - unidade sem dado de campo em cadastro que
                              # teve outra unidade atualizada
COR_NAO_ATUALIZA = "E06666"   # vermelho escuro - cadastro do mesmo lote que nao
                              # recebe o dado de campo
COR_CONDOMINIO = "76A5AF"     # verde-agua forte - linha do condominio e unidades
COR_CONDOMINIO_MEMBRO = "D0E0E3"  # verde-agua claro - cadastro que forma o
                                  # condominio (intocado)
COR_CABECALHO  = "D9D9D9"

# --- comportamento ----------------------------------------------------------
MANTER_SE_BRUTA_VAZIA = True  # True = vazio na bruta nao apaga o dado da prefeitura
ADICIONAR_NOVOS       = True  # True = registro so da bruta vira linha nova
COLUNA_STATUS         = "_status"      # coluna extra no fim; "" para nao criar
# Coluna curta para filtrar a saida: ATUALIZADO, CONFIRMADO, UNIDADE NOVA,
# CADASTRO NOVO, PRINCIPAL, COMPLEMENTAR, CONDOMINIO, BAIXADO, etc. "" desliga.
COLUNA_CLASSIFICACAO  = "_situacao"
# Coluna ao lado, com a classificacao detalhada (valor original)
COLUNA_CLASSIF_HIST   = "_situacao_historico"
# Agrupamento: classificacao detalhada -> valor que aparece em _situacao
MAPA_SITUACAO = {
    "ATUALIZADO":                        "ATUALIZADO",
    "CONFIRMADO EM CAMPO":               "ATUALIZADO",
    "COMPLEMENTAR":                      "ATUALIZADO",
    "CONDOMINIO":                        "ATUALIZADO",
    "UNIDADE DO CONDOMINIO":             "ATUALIZADO",
    "MEMBRO DE CONDOMINIO":              "ATUALIZADO",
    "UNIDADE NOVA":                      "UNIDADE NOVA",
    "PRINCIPAL":                         "UNIDADE NOVA",
    "LOTE BALDIO":                       "UNIDADE NOVA",
    "CADASTRO NOVO":                     "CADASTRO NOVO",
    "CADASTRO NAO EXISTE NA PREFEITURA": "CADASTRO NOVO",
    "UNIDADE BAIXADA":                   "UNIDADE BAIXADA",
    "NAO ATUALIZAR":                     "CADASTRO BAIXADO",
    "CADASTRO CANCELADO":                "CADASTRO BAIXADO",
}
COPIAR_LINHAS_ACIMA   = True  # repete na saida as linhas acima do cabecalho
FAZER_BACKUP          = True  # copia .backup antes de salvar

# ===========================================================================


# ---------------------------------------------------------------- utilitarios

NULOS = set(["", "null", "none", "nan", "#n/d", "#nome?"])


def _sem_acento(v):
    s = unicodedata.normalize("NFKD", str(v))
    return "".join(c for c in s if not unicodedata.combining(c))


NULOS.update(_sem_acento(x).strip().lower() for x in TEXTOS_NULOS)


def norm_cab(v):
    """Nome de coluna sem acento, minusculo, sem espaco duplo."""
    if v is None:
        return ""
    s = str(v).replace("\n", " ").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).lower()


def norm_valor(v):
    """Forma canonica de um valor, para comparar sem falso positivo."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = re.sub(r"\s+", " ", str(v).strip())
    try:
        f = float(s.replace(",", "."))
        return str(int(f)) if f.is_integer() else ("%.10g" % f)
    except ValueError:
        return s.lower()


def num(v):
    """Converte para float o que der; o resto vira None."""
    if v is None or isinstance(v, datetime):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(" ", "")
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")   # 1.234,56
    else:
        t = t.replace(",", ".")
    try:
        f = float(t)
    except ValueError:
        return None
    return f if math.isfinite(f) else None


def arredonda(v, casas):
    """Arredonda numero nao inteiro; texto e inteiro passam intactos."""
    if v is None or isinstance(v, datetime):
        return v
    try:
        f = float(str(v).strip().replace(",", "."))
    except (ValueError, AttributeError):
        return v
    if f == int(f):
        return v
    return round(f, casas)


def vazio(v):
    """Vazio de verdade: None, espaco em branco e os textos de TEXTOS_NULOS."""
    return v is None or _sem_acento(v).strip().lower() in NULOS


def norm_chave(v):
    """Chave como texto: 123.0, '123' e ' 123 ' viram '123'."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def so_digitos(v):
    """'123.456.789-00' -> '12345678900'; sem nenhum digito -> ''."""
    return re.sub(r"\D", "", norm_chave(v))


assert so_digitos("123. 456.789-00") == "12345678900"
assert so_digitos(12345678900.0) == "12345678900"
assert so_digitos("-") == "" and so_digitos(None) == ""


def achar_aba(wb, nome):
    alvo = norm_cab(nome)
    for ws in wb.worksheets:
        if norm_cab(ws.title) == alvo:
            return ws
    raise SystemExit(
        "Nao achei a aba '%s'. Abas do arquivo: %s\n"
        "Ajuste os nomes no bloco CONFIGURACAO."
        % (nome, ", ".join(w.title for w in wb.worksheets))
    )


def ler_tabela(ws_val, linha_cab, ws_raw=None):
    """(cabecalhos normalizados, cabecalhos originais, linhas, linhas_acima).

    ws_raw socorre celulas que o Excel guardou como erro: textos que comecam com
    '=' (ex.: '=- LOC. TRIUNFO') viram #NOME? no valor calculado.
    """
    val = list(ws_val.iter_rows(min_row=1, values_only=True))
    raw = list(ws_raw.iter_rows(min_row=1, values_only=True)) if ws_raw else None
    if len(val) < linha_cab:
        return [], [], [], []

    cab_orig = list(val[linha_cab - 1])
    while cab_orig and (cab_orig[-1] is None or str(cab_orig[-1]).strip() == ""):
        cab_orig.pop()
    n = len(cab_orig)
    cab = [norm_cab(c) for c in cab_orig]

    acima = [list(r)[:n] + [None] * max(0, n - len(r)) for r in val[:linha_cab - 1]]

    dados = []
    for i in range(linha_cab, len(val)):
        row = list(val[i])[:n] + [None] * max(0, n - len(val[i]))
        if raw is not None and i < len(raw):
            rr = list(raw[i])
            for j in range(min(n, len(rr))):
                if (row[j] is None or (isinstance(row[j], str) and row[j].startswith("#"))) \
                        and isinstance(rr[j], str) and rr[j].startswith("="):
                    row[j] = rr[j][1:].strip()
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        dados.append(row)
    return cab, cab_orig, dados, acima


def indice_de(cab, nome, onde, obrigatoria=True):
    alvo = norm_cab(nome)
    if alvo in cab:
        return cab.index(alvo)
    for i, c in enumerate(cab):
        if alvo and alvo in c:
            return i
    if obrigatoria:
        raise SystemExit("Nao achei a coluna '%s' na aba %s." % (nome, onde))
    return None


def primeira_col(cab, apelidos, obrigatoria, onde):
    for a in apelidos:
        a = norm_cab(a)
        if a in cab:
            return cab.index(a)
    for a in apelidos:
        a = norm_cab(a)
        for i, c in enumerate(cab):
            if a and a in c:
                return i
    if obrigatoria:
        raise SystemExit(
            "Na aba %s nao achei nenhuma coluna chamada %s.\nColunas: %s"
            % (onde, " / ".join(apelidos), ", ".join(c for c in cab if c))
        )
    return None


# ---------------------------------------------------------------------- main

def main():
    os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

    caminho = ARQUIVO
    if not caminho:
        cands = [f for f in os.listdir(".")
                 if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
        if len(cands) != 1:
            raise SystemExit("Preencha ARQUIVO no topo do script (achei %d planilhas)."
                             % len(cands))
        caminho = cands[0]
    if not os.path.exists(caminho):
        raise SystemExit("Nao achei '%s' na pasta %s" % (caminho, os.getcwd()))

    print("Arquivo: %s" % caminho)
    print("Lendo...")
    wb_val = openpyxl.load_workbook(caminho, data_only=True)
    wb     = openpyxl.load_workbook(caminho)

    cab_pref, cab_pref_orig, dados_pref, acima_pref = ler_tabela(
        achar_aba(wb_val, ABA_PREFEITURA), LINHA_CAB_PREFEITURA,
        achar_aba(wb, ABA_PREFEITURA))
    cab_bru, _, dados_bru, _ = ler_tabela(achar_aba(wb_val, ABA_BRUTA), LINHA_CAB_BRUTA)
    cab_dep, _, dados_dep, _ = ler_tabela(achar_aba(wb_val, ABA_DEPARA), LINHA_CAB_DEPARA)

    # cancelados por unificacao: lidos antes do filtro de situacao, porque o
    # cadastro cancelado pode ja estar desativado na prefeitura
    cancelado_para, unificados, avisos_cancel = {}, {}, []
    cancel_mats, cancel_area = {}, {}
    cancelados_sem_destino = set()     # cancelado sem unificacao: so baixa
    _ws_can = next((w for w in wb_val.worksheets
                    if norm_cab(w.title) == norm_cab(ABA_CANCELADOS)), None) \
        if ABA_CANCELADOS else None
    if _ws_can is not None:
        _cab_can, _, _dados_can, _ = ler_tabela(_ws_can, 1)
        _i_dst = indice_de(_cab_can, COLUNA_CANCELADO_DESTINO, ABA_CANCELADOS)
        _i_c0 = indice_de(cab_pref, COLUNA_CHAVE_PREFEITURA, ABA_PREFEITURA)
        _i_m0 = indice_de(cab_pref, COLUNA_MATRICULA, ABA_PREFEITURA)
        _i_a0 = indice_de(cab_pref, COLUNA_AREA_MATRICULA, ABA_PREFEITURA)
        for _r in dados_pref:
            _c = norm_chave(_r[_i_c0])
            if not vazio(_r[_i_m0]):
                _m = str(_r[_i_m0]).strip()
                if _m not in cancel_mats.setdefault(_c, []):
                    cancel_mats[_c].append(_m)
            _a = num(_r[_i_a0])
            if _a is not None:                  # uma vez por cadastro, nao por unidade
                cancel_area[_c] = max(cancel_area.get(_c, 0.0), _a)
        _todos = {norm_chave(_r[_i_c0]) for _r in dados_pref}
        _sem_dst = {norm_cab(x) for x in CANCELADO_SEM_DESTINO}
        for _r in _dados_can:
            _c = norm_chave(_r[0])
            if not _c:
                continue
            if norm_cab(_r[_i_dst]) in _sem_dst:
                if _c not in _todos:
                    avisos_cancel.append("%s: cancelado nao existe na prefeitura" % _c)
                cancelados_sem_destino.add(_c)
                continue
            _nums = re.findall(r"\d+", str(_r[_i_dst] or ""))
            if len(_nums) != 1:
                avisos_cancel.append("%s: '%s' nao tem UM numero de cadastro"
                                     % (_c, _r[_i_dst]))
                continue
            if _c not in _todos:
                avisos_cancel.append("%s: cancelado nao existe na prefeitura" % _c)
            # shortcut: cadeia A->B e B->C nao e seguida, tratar se aparecer
            cancelado_para[_c] = _nums[0]
            unificados.setdefault(_nums[0], []).append(_c)

    # cadastros desativados (17-Situacao Cadastro <> 1) saem do jogo
    n_desativadas = n_cads_desativados = 0
    if COLUNA_SITUACAO_CADASTRO:
        _i_sit = indice_de(cab_pref, COLUNA_SITUACAO_CADASTRO, ABA_PREFEITURA,
                           obrigatoria=False)
        _i_cad0 = indice_de(cab_pref, COLUNA_CHAVE_PREFEITURA, ABA_PREFEITURA)
        if _i_sit is not None:
            _ativas = {norm_chave(x) for x in SITUACAO_CADASTRO_ATIVA}
            _fora = [r for r in dados_pref if norm_chave(r[_i_sit]) not in _ativas]
            n_desativadas = len(_fora)
            n_cads_desativados = len({norm_chave(r[_i_cad0]) for r in _fora})
            dados_pref = [r for r in dados_pref
                          if norm_chave(r[_i_sit]) in _ativas]
            acima_pref = acima_pref

    print("  prefeitura: %d colunas, %d linhas" % (len(cab_pref), len(dados_pref)))
    print("  bruta     : %d colunas, %d linhas" % (len(cab_bru), len(dados_bru)))
    if n_desativadas:
        print("  ignorei %d linhas de cadastro desativado (%s <> %s): %d cadastros"
              % (n_desativadas, COLUNA_SITUACAO_CADASTRO,
                 "/".join(SITUACAO_CADASTRO_ATIVA), n_cads_desativados))

    # ------------------------------------------------------------- depara
    i_org = primeira_col(cab_dep, DEPARA_COL_ORIGEM,  True,  ABA_DEPARA)
    i_dst = primeira_col(cab_dep, DEPARA_COL_DESTINO, True,  ABA_DEPARA)
    i_atv = primeira_col(cab_dep, DEPARA_COL_ATIVO,   False, ABA_DEPARA)
    ativos = [norm_cab(x) for x in DEPARA_VALOR_ATIVO]

    brutos_pares, ignorados, novas_cols = [], [], []
    _linhas_dep = list(dados_dep)
    _ja = {(norm_cab(r[i_org]), norm_cab(r[i_dst])) for r in _linhas_dep
           if not vazio(r[i_org]) and not vazio(r[i_dst])}
    _extras_novos = []
    for _org, _dst in COLUNAS_SOLO_EXTRAS:
        if (norm_cab(_org), norm_cab(_dst)) in _ja:
            continue
        _r = [None] * max(i_org, i_dst, i_atv if i_atv is not None else 0) 
        _r = _r + [None]
        _r[i_org], _r[i_dst] = _org, _dst
        if i_atv is not None:
            _r[i_atv] = DEPARA_VALOR_ATIVO[0]
        _linhas_dep.append(_r)
        _extras_novos.append((_org, _dst))
    _troca = {norm_cab(k): v for k, v in DEPARA_TROCAR_ORIGEM.items()}
    _renomeia = {norm_cab(k): v for k, v in RENOMEAR_DESTINO.items()}
    for row in _linhas_dep:
        org, dst = row[i_org], row[i_dst]
        if vazio(org) or vazio(dst):
            continue
        if norm_cab(dst) in _troca:
            org = _troca[norm_cab(dst)]
        if norm_cab(dst) in _renomeia:
            dst = _renomeia[norm_cab(dst)]
        if i_atv is not None:
            flag = norm_cab(row[i_atv])
            marcado = flag in ativos
            if DEPARA_SO_MARCADOS and not marcado:
                continue
            if not DEPARA_SO_MARCADOS and flag and not marcado:
                continue
        n_org, n_dst = norm_cab(org), norm_cab(dst)
        if n_org not in cab_bru:
            for _ap in APELIDOS_COLUNAS_BRUTA.get(str(org).strip(), ()):
                if norm_cab(_ap) in cab_bru:
                    print("    [apelido] '%s' nao existe na bruta; usei '%s'"
                          % (org, _ap))
                    n_org = norm_cab(_ap)
                    break
        if n_org not in cab_bru:
            ignorados.append("a bruta nao tem a coluna '%s'" % org)
            continue
        if n_dst not in cab_pref:
            if POSICAO_COLUNAS_NOVAS in ("fim", "inicio"):
                novas_cols.append((n_dst, str(dst).strip()))
            else:
                ignorados.append("a prefeitura nao tem a coluna '%s'" % dst)
                continue
        brutos_pares.append((n_org, n_dst, str(dst).strip()))

    cab_saida      = list(cab_pref)
    cab_saida_orig = list(cab_pref_orig)
    vistos = set()
    novas = [(n, o) for n, o in novas_cols if not (n in vistos or vistos.add(n))]
    if COLUNA_SEPARADOR and novas:
        novas = [(norm_cab(COLUNA_SEPARADOR), COLUNA_SEPARADOR)] + novas
    if novas and POSICAO_COLUNAS_NOVAS == "inicio":
        cab_saida      = [n for n, _ in novas] + cab_saida
        cab_saida_orig = [o for _, o in novas] + cab_saida_orig
    elif novas:
        cab_saida      += [n for n, _ in novas]
        cab_saida_orig += [o for _, o in novas]

    pares = [(cab_bru.index(o), cab_saida.index(d), nome) for o, d, nome in brutos_pares]
    if not pares:
        raise SystemExit("O depara nao produziu nenhum par de colunas valido.")

    print("  depara    : %d pares de colunas" % len(pares))
    for _n, o in novas:
        print("    [coluna criada na saida] %s" % o)
    for msg in sorted(set(ignorados)):
        print("    [ignorado] %s" % msg)

    # ------------------------------------------------------------- chaves
    k_pref = indice_de(cab_saida, COLUNA_CHAVE_PREFEITURA, ABA_PREFEITURA)
    k_bru  = indice_de(cab_bru,   COLUNA_CHAVE_BRUTA,      ABA_BRUTA)
    u_pref = u_bru = None
    if CHAVE_COM_UNIDADE:
        u_pref = indice_de(cab_saida, COLUNA_UNIDADE_PREFEITURA, ABA_PREFEITURA)
        u_bru  = indice_de(cab_bru,   COLUNA_UNIDADE_BRUTA,      ABA_BRUTA)
    ec_bru = None
    if USAR_CADASTRO_EDIFICACAO:
        ec_bru = indice_de(cab_bru, COLUNA_CADASTRO_EDIFICACAO, ABA_BRUTA,
                           obrigatoria=False)
        if ec_bru is None:
            print("    [aviso] a bruta nao tem a coluna '%s'; vinculo so pelo %s"
                  % (COLUNA_CADASTRO_EDIFICACAO, COLUNA_CHAVE_BRUTA))

    def chave_de(row, ik, iu):
        k = norm_chave(row[ik])
        if not k:
            return None
        return (k, norm_chave(row[iu])) if CHAVE_COM_UNIDADE else (k,)

    i_cad_p = indice_de(cab_pref, COLUNA_CHAVE_PREFEITURA, ABA_PREFEITURA)
    i_uni_p = indice_de(cab_pref, COLUNA_UNIDADE_PREFEITURA, ABA_PREFEITURA) \
        if CHAVE_COM_UNIDADE else None
    chaves_pref = set()
    for row in dados_pref:
        c = chave_de(row, i_cad_p, i_uni_p)
        if c:
            chaves_pref.add(c)

    cads_pref = {c[0] for c in chaves_pref}

    # perfil de cada cadastro da prefeitura: tem unidade de avaliacao? tem
    # matricula? qual a area construida? qual a chave da linha de terreno?
    i_mat_pp = indice_de(cab_pref, COLUNA_MATRICULA, ABA_PREFEITURA,
                         obrigatoria=False)
    i_area_pp = indice_de(cab_pref, COLUNA_AREA_CONSTRUIDA_CADASTRO, ABA_PREFEITURA,
                          obrigatoria=False)
    i_uni_pp = indice_de(cab_pref, COLUNA_UNIDADE_PREFEITURA, ABA_PREFEITURA,
                         obrigatoria=False)
    cads_com_unidade, cad_tem_matricula = set(), set()
    cad_area, chave_terreno, cad_max_unid = {}, {}, {}
    cad_matriculas = {}
    for row in dados_pref:
        _c = norm_chave(row[i_cad_p])
        if not _c:
            continue
        if i_uni_pp is not None and not vazio(row[i_uni_pp]):
            cads_com_unidade.add(_c)
        elif _c not in chave_terreno:
            chave_terreno[_c] = chave_de(row, i_cad_p, i_uni_p)
        if i_mat_pp is not None and not vazio(row[i_mat_pp]):
            cad_tem_matricula.add(_c)
            cad_matriculas.setdefault(_c, set()).add(norm_chave(row[i_mat_pp]))
        if i_area_pp is not None:
            _a = num(row[i_area_pp])
            if _a is not None:
                cad_area[_c] = max(cad_area.get(_c, 0.0), _a)
        if i_uni_pp is not None:
            _u = num(row[i_uni_pp])
            if _u is not None:
                cad_max_unid[_c] = max(cad_max_unid.get(_c, 0), int(_u))

    def escolhe_cadastro(cads):
        """Entre cadastros do mesmo lote, quem recebe o dado de campo."""
        com = [c for c in cads if c in cad_tem_matricula]
        candidatos = com if com else list(cads)
        return max(sorted(candidatos), key=lambda c: cad_area.get(c, 0.0))

    unidade_terreno = {}     # chave da linha de terreno -> unidade da coleta

    def parse_cadastro_edificacao(v):
        """'603800.2 / 603800.3' -> ([('603800','2'), ('603800','3')], erros)."""
        ok, erros = [], []
        for parte in re.split(r"[;/,\n]", str(v)):
            parte = parte.strip()
            if not parte:
                continue
            m = re.match(r"^(\d+)\s*\.\s*(\d+)$", parte)
            if not m:
                erros.append(parte)
                continue
            par = (m.group(1), m.group(2)) if CHAVE_COM_UNIDADE else (m.group(1),)
            if (par not in chaves_pref and CHAVE_COM_UNIDADE
                    and TERRENO_RECEBE_UNIDADE_DA_COLETA
                    and par[0] in chave_terreno and par[0] not in cads_com_unidade):
                _cht = chave_terreno[par[0]]
                unidade_terreno.setdefault(_cht, par[1])
                ok.append(_cht)
                continue
            if par not in chaves_pref:
                erros.append(parte + " (nao existe na prefeitura)")
                continue
            ok.append(par)
        return ok, erros

    # ----------------------- cascata de vinculo, linha a linha da bruta -----
    #  1) COLUNA_CADASTRO_EDIFICACAO, no formato cadastro.unidade (pode listar
    #     varios separados por / ; ,) -> vincula a TODOS eles
    #  2) COLUNA_CHAVE_BRUTA + unidade
    #  3) nada -> linha nova, com TEXTO_CADASTRO_NOVO no numero do cadastro
    #  valor torto na coluna 1 nao vincula: a linha e destacada para conferencia
    oc_bru = None
    if REGRA_LOTE_BALDIO:
        oc_bru = indice_de(cab_bru, COLUNA_OCUPACAO_BRUTA, ABA_BRUTA,
                           obrigatoria=False)
        if oc_bru is None:
            print("    [aviso] a bruta nao tem a coluna '%s'; regra do lote baldio"
                  " desligada" % COLUNA_OCUPACAO_BRUTA)
    baldios = [norm_chave(x) for x in VALOR_OCUPACAO_BALDIO]

    idx_bruta, duplicados = {}, []
    alvos_da_linha, via_da_linha = {}, {}
    linhas_erro, linhas_novas, linhas_baldio = [], [], []
    coleta_em_baixado = []    # coleta descartada por apontar para cadastro baixado
    n_via_edif = n_via_lote = n_um_por_lote = 0
    n_transferidas = 0

    n_baldio_atualiza = 0
    ul_bru = indice_de(cab_bru, COLUNA_UUID_LOTE, ABA_BRUTA, obrigatoria=False) \
        if UM_CADASTRO_POR_LOTE else None

    # ------------------ coleta com varios numeros de cadastro na mesma celula
    # "54400 / 598800": fica so um (matricula, depois maior area construida) e
    # os outros que existem na prefeitura sao marcados para BAIXAR.
    escolha_multi, cads_baixar_multi, multi_origem = {}, {}, {}
    multi_grupo = {}          # cadastro principal -> cadastros do grupo
    cads_baixados = ({norm_chave(x) for x in CADASTROS_BAIXADOS}
                     | set(cancelado_para) | cancelados_sem_destino)
    multi_usados = set()

    def _cads_do_texto(v):
        return [norm_chave(x) for x in re.split(r"[;/,\n]", str(v)) if x.strip()]

    if MULTI_CADASTRO_ESCOLHE:
        for _row in dados_bru:
            if vazio(_row[k_bru]):
                continue
            _txt = str(_row[k_bru]).strip()
            if _txt in escolha_multi or _txt in multi_origem:
                continue
            _cs = _cads_do_texto(_txt)
            if len(_cs) < 2:
                continue
            _ex = [c for c in _cs if c in cads_pref]
            _liv = [c for c in _ex if c not in cads_baixados]
            _esc = escolhe_cadastro(_liv or _ex) if _ex else _cs[0]
            escolha_multi[_txt] = _esc
            multi_origem[_txt] = _cs
            for _c in _ex:
                if _c != _esc:
                    cads_baixar_multi[_c] = _esc
            if len(_ex) > 1:
                multi_grupo.setdefault(_esc, set()).update(_ex)

    # linhas da coleta cuja edificacao e dividida por mais de um cadastro
    edif_varios_cads = set()
    edif_grupo_cads, edif_grupo_linha = {}, {}
    if EDIFICACAO_MULTICADASTRO_TRAVA_AREA:
        _ue_b0 = indice_de(cab_bru, COLUNA_UUID_EDIFICACAO, ABA_BRUTA,
                           obrigatoria=False)
        _cads_por_ue, _cads_por_linha = {}, {}
        for _row in dados_bru:
            _cs = set()
            if ec_bru is not None and not vazio(_row[ec_bru]):
                for _p in re.split(r"[;/,\n]", str(_row[ec_bru])):
                    _m = re.match(r"^\s*(\d+)\s*\.\s*(\d+)\s*$", _p)
                    if _m:
                        _cs.add(_m.group(1))
            _cs -= cads_baixados          # cadastro baixado nao conta no grupo
            _cads_por_linha[id(_row)] = _cs
            _ue = norm_chave(_row[_ue_b0]) if _ue_b0 is not None else ""
            if _ue:
                _cads_por_ue.setdefault(_ue, set()).update(_cs)
        for _row in dados_bru:
            _ue = norm_chave(_row[_ue_b0]) if _ue_b0 is not None else ""
            _cs = _cads_por_linha[id(_row)] | (_cads_por_ue.get(_ue) or set())
            if len(_cs) > 1:
                edif_varios_cads.add(id(_row))
                _k = _ue or ("linha:%d" % id(_row))
                if _k not in edif_grupo_cads:
                    edif_grupo_cads[_k] = _cs
                    edif_grupo_linha[_k] = _row

    def cad_res(row):
        """Numero do cadastro da coleta, ja resolvido quando vem varios."""
        if vazio(row[k_bru]):
            return ""
        _txt = str(row[k_bru]).strip()
        _esc = escolha_multi.get(_txt)
        if _esc:
            multi_usados.add(_txt)
            return _esc
        return norm_chave(row[k_bru])

    # passo 1: o que cada linha da coleta pretende atualizar
    tentativas = []
    n_baldio_atualiza = 0
    for row in dados_bru:
        if oc_bru is not None and norm_chave(row[oc_bru]) in baldios:
            _cad = cad_res(row)
            if not _cad and ec_bru is not None and not vazio(row[ec_bru]):
                _ok, _e = parse_cadastro_edificacao(row[ec_bru])
                if _ok:
                    _cad = _ok[0][0]
            _ch = chave_terreno.get(_cad) if BALDIO_ATUALIZA_TERRENO else None
            if _ch and _cad not in cads_com_unidade:
                if _ch not in idx_bruta:
                    idx_bruta[_ch] = row
                alvos_da_linha[id(row)] = [_ch]
                via_da_linha[id(row)] = ("%s: %s (cadastro ja e terreno)"
                                         % (COLUNA_OCUPACAO_BRUTA,
                                            norm_chave(row[oc_bru])))
                n_baldio_atualiza += 1
                continue
            linhas_baldio.append(row)
            continue
        v = row[ec_bru] if ec_bru is not None else None
        if not vazio(v):
            ok, erros = parse_cadastro_edificacao(v)
            if erros or not ok:
                linhas_erro.append((row, str(v).strip(), erros))
                continue
            tentativas.append((row, ok,
                               "%s: %s" % (COLUNA_CADASTRO_EDIFICACAO,
                                           str(v).strip()), "edif"))
        else:
            _cd = cad_res(row)
            c = None
            if _cd:
                c = (_cd, norm_chave(row[u_bru])) if CHAVE_COM_UNIDADE else (_cd,)
            if (c is not None and c not in chaves_pref and CHAVE_COM_UNIDADE
                    and TERRENO_RECEBE_UNIDADE_DA_COLETA and c[1]
                    and c[0] in chave_terreno and c[0] not in cads_com_unidade):
                unidade_terreno.setdefault(chave_terreno[c[0]], c[1])
                _c_via = c
                c = chave_terreno[c[0]]
            else:
                _c_via = c
            if c is not None and c in chaves_pref:
                tentativas.append((row, [c],
                                   "%s: %s" % (COLUNA_CHAVE_BRUTA,
                                               "/".join(x for x in _c_via if x)),
                                   "lote"))
            elif (DESCARTAR_COLETA_NOVA_EM_BAIXADO and c is not None
                    and c[0] in cads_baixados):
                # unidade que nao existe na prefeitura, em cadastro baixado ou
                # cancelado: descartada, como a coleta das unidades que existem
                coleta_em_baixado.append(
                    (c[0], c[1] if CHAVE_COM_UNIDADE else "",
                     "%s: %s - unidade nao existe na prefeitura"
                     % (COLUNA_CHAVE_BRUTA, "/".join(x for x in c if x))))
            else:
                linhas_novas.append((row, c))

    # passo 1b: CONDOMINIO -> edificacao com cadastros de matriculas diferentes
    cond_info, cads_condominio, linhas_condominio = {}, {}, []
    coleta_absorvida_cond = set()   # linhas da coleta resumidas na linha do cond.
    if REGRA_CONDOMINIO:
        _ue_b = indice_de(cab_bru, COLUNA_UUID_EDIFICACAO, ABA_BRUTA,
                          obrigatoria=False)
        if _ue_b is not None:
            _cads_ue = {}
            for row, alvos, _via, _t in tentativas:
                _ue = norm_chave(row[_ue_b])
                if _ue:
                    _cads_ue.setdefault(_ue, set()).update(a[0] for a in alvos)
            cond_de_ue = {}
            for _ue, _cs in sorted(_cads_ue.items()):
                if len(_cs) < 2:
                    continue
                _ms = set()
                for _c in _cs:
                    _ms |= cad_matriculas.get(_c, set())
                if len(_ms) >= CONDOMINIO_MIN_MATRICULAS:
                    _cid = "COND %d" % (len(cond_info) + 1)
                    cond_de_ue[_ue] = _cid
                    cond_info[_cid] = {"cads": set(_cs), "lotes": set()}
                    for _c in _cs:
                        cads_condominio[_c] = _cid
            if cond_info:
                for row, alvos, _via, _t in tentativas:
                    _ue = norm_chave(row[_ue_b])
                    _ul = norm_chave(row[ul_bru]) if ul_bru is not None else ""
                    if _ue in cond_de_ue and _ul:
                        cond_info[cond_de_ue[_ue]]["lotes"].add(_ul)
                lote_cond = {_ul: _cid for _cid, _inf in cond_info.items()
                             for _ul in _inf["lotes"]}
                _resto, _por_cid = [], {}
                for row, alvos, via, t in tentativas:
                    _ue = norm_chave(row[_ue_b])
                    _ul = norm_chave(row[ul_bru]) if ul_bru is not None else ""
                    _cid = cond_de_ue.get(_ue) or lote_cond.get(_ul)
                    if _cid is None:
                        _resto.append((row, alvos, via, t))
                        continue
                    _por_cid.setdefault(_cid, []).append(
                        (row, alvos, via, cond_de_ue.get(_ue) == _cid))
                for _cid, _lst in _por_cid.items():
                    _mesmas = [x for x in _lst if x[3]]
                    _outras = [x for x in _lst if not x[3]]
                    if _mesmas:
                        _r0, _a0, _v0, _x = _mesmas[0]
                        linhas_condominio.append((_r0, _cid, "condominio", _v0))
                        # os membros recebem lote + classificacoes dessa linha
                        for _a in _a0:
                            if _a not in idx_bruta:
                                idx_bruta[_a] = _r0
                        via_da_linha[id(_r0)] = _v0 + "  [%s]" % _cid
                        # varias linhas da coleta na mesma edificacao = unidades
                        # do condominio (ex.: apartamentos)
                        for _m in _mesmas[1:]:
                            coleta_absorvida_cond.add(id(_m[0]))
                        if len(_mesmas) > 1 and CONDOMINIO_UNIDADES_MESMA_EDIF:
                            for row, alvos, via, _x in _mesmas:
                                linhas_condominio.append(
                                    (row, _cid, "unidade_mesma", via))
                    for row, alvos, via, _x in _outras:
                        linhas_condominio.append((row, _cid, "unidade", via))
                tentativas = _resto
                # linha nova (sem vinculo) no mesmo lote tambem entra no condominio
                _resto_n = []
                for row, c in linhas_novas:
                    _ul = norm_chave(row[ul_bru]) if ul_bru is not None else ""
                    if _ul in lote_cond:
                        linhas_condominio.append((row, lote_cond[_ul], "unidade",
                                                  "sem vinculo, mesmo lote"))
                    else:
                        _resto_n.append((row, c))
                linhas_novas[:] = _resto_n

    # membro de condominio nunca e baixado pela regra dos numeros multiplos
    for _c in list(cads_baixar_multi):
        if _c in cads_condominio:
            del cads_baixar_multi[_c]

    # passo 2: lote com mais de um cadastro -> escolhe quem recebe o dado
    escolhido_do_lote, cads_nao_atualizar = {}, set()
    if UM_CADASTRO_POR_LOTE and ul_bru is not None:
        lote_cads = {}
        for row, alvos, _via, _t in tentativas:
            ul = norm_chave(row[ul_bru])
            if ul:
                lote_cads.setdefault(ul, set()).update(a[0] for a in alvos)
        for ul, cs in lote_cads.items():
            if len(cs) > 1:
                esc = escolhe_cadastro([c for c in cs if c not in cads_baixados] or cs)
                escolhido_do_lote[ul] = esc
                cads_nao_atualizar |= (cs - {esc})

    # passo 3: registra o vinculo
    linhas_transferidas = []
    for row, alvos, via, tipo_via in tentativas:
        ul = norm_chave(row[ul_bru]) if ul_bru is not None else ""
        esc = escolhido_do_lote.get(ul)
        if esc:
            cads_no_alvo = {a[0] for a in alvos}
            fora = sorted(cads_no_alvo - {esc})
            motivo = ("tem matricula" if esc in cad_tem_matricula
                      else "maior area construida")
            if fora:
                manter = [a for a in alvos if a[0] == esc]
                if manter:
                    alvos = manter
                    via += ("  [mesmo lote: atualizado o %s (%s); nao atualizados:"
                            " %s]" % (esc, motivo, ", ".join(fora)))
                    n_um_por_lote += 1
                else:
                    linhas_transferidas.append((row, esc, fora, motivo, via))
                    continue
        for alvo in alvos:
            if alvo in idx_bruta:
                duplicados.append("/".join(x for x in alvo if x))
                continue
            idx_bruta[alvo] = row
        alvos_da_linha[id(row)] = alvos
        via_da_linha[id(row)] = via
        if tipo_via == "edif":
            n_via_edif += 1
        else:
            n_via_lote += 1

    # cadastros que ja tem unidade na prefeitura e hoje sao lote baldio: a
    # prefeitura vai ter que baixar essas unidades
    cads_para_baixar = set()
    if REGRA_LOTE_BALDIO:
        for _row in linhas_baldio:
            _cad = norm_chave(_row[k_bru])
            if not _cad and ec_bru is not None and not vazio(_row[ec_bru]):
                _ok, _e = parse_cadastro_edificacao(_row[ec_bru])
                if _ok:
                    _cad = _ok[0][0]
            if _cad and _cad in cads_pref:
                cads_para_baixar.add(_cad)
        # so vale para cadastro que tem unidade de avaliacao de verdade: linha de
        # terreno (unidade vazia) nao tem o que baixar
        if CHAVE_COM_UNIDADE:
            cads_para_baixar &= cads_com_unidade

    # ------------------------------------- regra da edificacao compartilhada
    grupos = {}          # uuid_edif -> {registros:set, cadastros:set, area, linha}
    edif_do_registro = {}
    e_bru = a_bru = i_area_dest = i_matric = None
    cad_com_matricula = set()
    grupos_multi, grupos_principal, grupos_mistos = {}, {}, []

    if REGRA_EDIFICACAO:
        e_bru = indice_de(cab_bru, COLUNA_UUID_EDIFICACAO, ABA_BRUTA)
        a_bru = indice_de(cab_bru, ORIGEM_AREA_PRINCIPAL, ABA_BRUTA)
        i_area_dest = indice_de(cab_saida, DESTINO_AREA_PRINCIPAL, ABA_PREFEITURA)
        idx_area_ctrl = set()
        for nome in COLUNAS_AREA_CONTROLADAS:
            i = indice_de(cab_saida, nome, ABA_PREFEITURA, obrigatoria=False)
            if i is not None:
                idx_area_ctrl.add(i)

        # cadastro que tem matricula na prefeitura (basta uma unidade ter)
        i_mat_p = indice_de(cab_pref, COLUNA_MATRICULA, ABA_PREFEITURA)
        for row in dados_pref:
            if not vazio(row[i_mat_p]):
                cad_com_matricula.add(norm_chave(row[i_cad_p]))

        # o grupo junta os registros da mesma edificacao, venham eles de linhas
        # diferentes da bruta com o mesmo uuid_edificacao ou de uma unica celula
        # que lista varios cadastro.unidade
        for row in dados_bru:
            alvos = alvos_da_linha.get(id(row))
            if not alvos:
                continue
            ue = norm_chave(row[e_bru]) or ("celula:%d" % id(row))
            g = grupos.setdefault(ue, {"registros": set(), "cadastros": set(),
                                       "area": None, "linha": row})
            g["registros"].update(alvos)
            g["cadastros"].update(a[0] for a in alvos)
            if g["area"] is None and not vazio(row[a_bru]):
                g["area"] = row[a_bru]

        for ue, g in grupos.items():
            if len(g["registros"]) < 2:
                continue
            grupos_multi[ue] = g
            com = {c for c in g["cadastros"] if c in cad_com_matricula}
            if com and com != g["cadastros"]:
                grupos_mistos.append(ue)
            tem_matricula = bool(com) if GRUPO_MISTO == "com_matricula" \
                else (com == g["cadastros"])
            g["tem_matricula"] = tem_matricula
            if not tem_matricula:
                grupos_principal[ue] = g

        for ue, g in grupos_multi.items():
            g["origem"] = ", ".join(sorted("/".join(x for x in reg if x)
                                           for reg in g["registros"]))
            for reg in g["registros"]:
                edif_do_registro.setdefault(reg, ue)
    else:
        idx_area_ctrl = set()

    # ------------------------------------------------------------- montagem
    n_col_dados = len(cab_saida)
    n_pref_col  = len(cab_pref)
    desloca = n_col_dados - n_pref_col if POSICAO_COLUNAS_NOVAS == "inicio" and novas else 0

    comp_pares = []
    if ABA_COMPARACAO:
        for _item in COMPARACOES:
            _cb, _cp = _item[0], _item[1]
            _lim = _item[2] if len(_item) > 2 else DESTACAR_DIF_ACIMA
            _col_b = _cb[5:] if _cb.lower().startswith("soma:") else _cb
            _ib = indice_de(cab_bru,  _col_b, ABA_BRUTA,     obrigatoria=False)
            if _ib is not None and _col_b != _cb:
                _ib = ("soma", _ib)
            _ip = indice_de(cab_pref, _cp, ABA_PREFEITURA, obrigatoria=False)
            if _ib is None or _ip is None:
                print("    [comparacao ignorada] %s x %s" % (_cb, _cp))
                continue
            _np = _item[3] if len(_item) > 3 else _cp
            _nc = _item[4] if len(_item) > 4 else _cb
            comp_pares.append((_ib, _ip, _cb, _cp, _lim, _np, _nc))
    # soma por cadastro para os pares "soma:" (uma vez por uuid_edificacao)
    soma_por_cad = {}
    _pares_soma = [p_[0][1] for p_ in comp_pares if isinstance(p_[0], tuple)]
    if _pares_soma:
        _ue_b2 = indice_de(cab_bru, COLUNA_UUID_EDIFICACAO, ABA_BRUTA,
                           obrigatoria=False)
        _linhas_do_cad = {}
        for _ch, _row in idx_bruta.items():
            _linhas_do_cad.setdefault(_ch[0], []).append(_row)
        for _row, _c in linhas_novas:
            if _c is not None:
                _linhas_do_cad.setdefault(_c[0], []).append(_row)
        for _row, _esc, _f, _m, _v in linhas_transferidas:
            _linhas_do_cad.setdefault(_esc, []).append(_row)
        for _ib_s in _pares_soma:
            _m = {}
            for _cad, _rows in _linhas_do_cad.items():
                _vistos, _tot, _achou = set(), 0.0, False
                for _row in _rows:
                    _ue = norm_chave(_row[_ue_b2]) if _ue_b2 is not None else id(_row)
                    if _ue in _vistos:
                        continue
                    _vistos.add(_ue)
                    _v = num(_row[_ib_s])
                    if _v is not None:
                        _tot += _v
                        _achou = True
                if _achou:
                    _m[_cad] = round(_tot, 3)
            soma_por_cad[_ib_s] = _m

    # area construida da prefeitura comparada pela SOMA do grupo de cadastros:
    # numeros multiplos na coleta, condominio e edificacao compartilhada - em
    # todos eles varios cadastros dividem a mesma edificacao coletada.
    soma_pref_multi = {}
    if comp_pares:
        _ips_soma = set()
        for _nome in MULTI_SOMA_COMPARACAO:
            _i = indice_de(cab_pref, _nome, ABA_PREFEITURA, obrigatoria=False)
            if _i is not None:
                _ips_soma.add(_i)
        _grupos_soma = [set(_g) for _g in multi_grupo.values()]
        for _inf in cond_info.values():
            _grupos_soma.append(set(_inf["cads"]))
        for _g in grupos_multi.values():
            if len(_g["cadastros"]) > 1:
                _grupos_soma.append(set(_g["cadastros"]))
        # junta grupos que compartilham cadastro
        _merge = []
        for _g in _grupos_soma:
            _g = {norm_chave(x) for x in _g if x}
            if len(_g) < 2:
                continue
            _juntos = [x for x in _merge if x & _g]
            for _x in _juntos:
                _g |= _x
                _merge.remove(_x)
            _merge.append(_g)
        _val_cad = {}
        for _ip in {p_[1] for p_ in comp_pares} & _ips_soma:
            _m = {}
            for _r in dados_pref:
                _c = norm_chave(_r[i_cad_p])
                _v = num(_r[_ip])
                if _c and _v is not None:
                    _m[_c] = max(_m.get(_c, 0.0), _v)
            _val_cad[_ip] = _m
        for _g in _merge:
            for _ip, _m in _val_cad.items():
                _tot = sum(_m.get(_c, 0.0) for _c in _g)
                if not _tot:
                    continue
                for _c in _g:
                    soma_pref_multi.setdefault(_c, {})[_ip] = round(_tot, 3)
        n_soma_grupo = len(_merge)

    def valor_pref_comp(row, ip, cad):
        """Valor da prefeitura na comparacao (somado quando o cadastro divide a
        edificacao com outros)."""
        _s = soma_pref_multi.get(norm_chave(cad), {})
        return _s.get(ip, row[ip])

    def valor_coleta(rb, ib, cad):
        if isinstance(ib, tuple):
            return soma_por_cad.get(ib[1], {}).get(norm_chave(cad))
        return rb[ib]

    dados_comp = []
    cads_atualizados, cads_novos, cads_unid_nova = set(), set(), set()
    cads_vinculados, cads_confirmados = set(), set()
    cads_baldio_exist, cads_baldio_novo, cads_principal = set(), set(), set()

    # unidade a unidade: quem vinculou e quem nao vinculou
    i_prop_p = indice_de(cab_pref, COLUNAS_HERDADAS_DO_CADASTRO[0], ABA_PREFEITURA,
                         obrigatoria=False) if COLUNAS_HERDADAS_DO_CADASTRO else None
    i_area_p = indice_de(cab_pref, DESTINO_AREA_PRINCIPAL, ABA_PREFEITURA,
                         obrigatoria=False)
    unidades_do_cad = {}          # cadastro -> lista de (unidade, vinculou, via, prop, area)
    coleta_do_cad = {}            # cadastro -> o que a coleta traz nesse cadastro
    if ABA_PENDENCIAS:
        for _row in dados_bru:
            _k = norm_chave(_row[k_bru])
            if not _k:
                continue
            _d = coleta_do_cad.setdefault(_k, {"linhas": 0, "unidades": [], "edif": []})
            _d["linhas"] += 1
            _u = norm_chave(_row[u_bru]) if CHAVE_COM_UNIDADE else ""
            if _u and _u not in _d["unidades"]:
                _d["unidades"].append(_u)
            _e = norm_chave(_row[ec_bru]) if ec_bru is not None else ""
            if _e and _e not in _d["edif"]:
                _d["edif"].append(_e)

    saida, usados = [], set()
    n_atualizadas = n_celulas = n_travadas = n_iguais = 0
    n_linhas_mudaram = n_celulas_uuid = n_baixar = n_nao_atualiza = 0
    n_codigos = n_membros_cond = n_linhas_cond = 0
    n_coleta_cond = len({id(_x[0]) for _x in linhas_condominio}
                        | coleta_absorvida_cond)
    contagem_por_col = {}

    # colunas que na prefeitura sao numero de verdade (nao texto com cara de numero)
    col_numerica = set()
    if AJUSTAR_TIPO_PELA_PREFEITURA:
        for _j in range(len(cab_pref)):
            _n = _t = 0
            for _row in dados_pref:
                _v = _row[_j]
                if vazio(_v):
                    continue
                _t += 1
                if isinstance(_v, (int, float)) and not isinstance(_v, bool):
                    _n += 1
                if _t >= 400:
                    break
            if _t and _n / float(_t) >= 0.8:
                col_numerica.add(_j + desloca)

    def casa_tipo(v, i_p):
        """Texto numerico vira numero quando a coluna e numerica na prefeitura."""
        if i_p not in col_numerica or not isinstance(v, str):
            return v
        t = v.strip()
        if not t or (t.startswith("0") and len(t) > 1 and not t.startswith("0.")):
            return v
        f = num(t)
        if f is None:
            return v
        return int(f) if f == int(f) and "." not in t and "," not in t else f

    # colunas que nao existiam na prefeitura (uuid...): mudam sempre, entao nao
    # contam como "valor alterado" na hora de dizer se o cadastro mudou
    idx_colunas_novas = set(range(n_pref_col, n_col_dados)) if not desloca \
        else set(range(0, n_col_dados - n_pref_col))

    novo_cad = []
    for _dst, _org, _fixo in DEPARA_CADASTRO_NOVO:
        _i_d = indice_de(cab_saida, _dst, ABA_PREFEITURA, obrigatoria=False)
        _i_o = indice_de(cab_bru, _org, ABA_BRUTA, obrigatoria=False) if _org else None
        if _i_d is None or (_org and _i_o is None):
            print("    [cadastro novo: ignorado] %s <- %s" % (_dst, _org or _fixo))
            continue
        novo_cad.append((_i_d, _i_o, _fixo))

    def preenche_cadastro_novo(valores, rb):
        for _i_d, _i_o, _fixo in novo_cad:
            _v = _fixo if _i_o is None else rb[_i_o]
            if not vazio(_v):
                valores[_i_d] = casa_tipo(_v, _i_d)

    def codigo_linha(valores):
        """Codigo da propria linha, no mesmo formato da lista do grupo."""
        _c = norm_chave(valores[k_pref])
        _u = norm_chave(valores[u_pref]) if CHAVE_COM_UNIDADE else ""
        return "%s/%s" % (_c, _u) if _u else _c

    idx_trava_cond = set()
    for _nome in (COLUNAS_AREA_TRAVADAS_CONDOMINIO
                  + COLUNAS_PROPRIETARIO_TRAVADAS_CONDOMINIO):
        _i = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
        if _i is not None:
            idx_trava_cond.add(_i)

    idx_protegidas = set()
    if PROTEGER_COLUNAS_CHAVE:
        idx_protegidas.add(k_pref)
        if CHAVE_COM_UNIDADE:
            idx_protegidas.add(u_pref)

    nulos_por_col = {}
    for _nome, _lista in TEXTOS_NULOS_POR_COLUNA.items():
        _i = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
        if _i is not None:
            nulos_por_col[_i] = set(_sem_acento(x).strip().lower() for x in _lista)

    idx_so_digitos = set()
    for _nome in COLUNAS_SO_DIGITOS:
        _i = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
        if _i is not None:
            idx_so_digitos.add(_i)

    herdadas = []
    for _nome in COLUNAS_HERDADAS_DO_CADASTRO:
        _is = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
        _ip = indice_de(cab_pref, _nome, ABA_PREFEITURA, obrigatoria=False)
        if _is is not None and _ip is not None:
            herdadas.append((_is, _ip))

    linhas_do_cadastro = {}
    if herdadas:
        for _row in dados_pref:
            linhas_do_cadastro.setdefault(norm_chave(_row[i_cad_p]), []).append(_row)

    padrao_linha_nova = {}
    for _nome, _txt in TEXTO_PADRAO_LINHA_NOVA.items():
        _i = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
        if _i is not None:
            padrao_linha_nova[_i] = _txt

    def eh_nulo(v, i_p):
        return vazio(v) or _sem_acento(v).strip().lower() in nulos_por_col.get(i_p, ())

    def _mesmo_nome(a, b):
        f = lambda v: " ".join(_sem_acento(v).upper().split())
        return bool(f(a)) and f(a) == f(b)
    assert _mesmo_nome(" Fábio  Mayer de Souza", "FABIO MAYER DE SOUZA")
    assert not _mesmo_nome("Adão Leite", "ADÃO DE LIMA LEITE")
    assert not _mesmo_nome("", "")

    doc_resp = None
    if DOC_RESPONSAVEL_DO_PROPRIETARIO:
        _n = DOC_RESPONSAVEL_DO_PROPRIETARIO
        doc_resp = [indice_de(cab_bru, x, ABA_BRUTA, obrigatoria=False) for x in _n[:2]] \
            + [indice_de(cab_saida, x, ABA_PREFEITURA, obrigatoria=False) for x in _n[2:]]
        if None in doc_resp:
            doc_resp = None
    n_doc_resp = [0, 0]      # [pegou o doc do proprietario, nome novo sem doc]

    fallback_idx = {}
    for _orig, _alts in FALLBACK_COLUNAS_BRUTA.items():
        _i = indice_de(cab_bru, _orig, ABA_BRUTA, obrigatoria=False)
        if _i is None:
            continue
        _lst = [indice_de(cab_bru, a, ABA_BRUTA, obrigatoria=False) for a in _alts]
        _lst = [x for x in _lst if x is not None]
        if _lst:
            fallback_idx[_i] = _lst

    def aplica(valores, rb, pintar, mantidas, iguais, travar_area, contar=True,
               idx_trava=None):
        alterou = False
        for i_b, i_p, nome in pares:
            if i_p in idx_protegidas:
                continue
            novo = rb[i_b]
            if eh_nulo(novo, i_p):
                for i_alt in fallback_idx.get(i_b, ()):
                    if not eh_nulo(rb[i_alt], i_p):
                        novo = rb[i_alt]
                        break
            if i_p in idx_so_digitos:
                novo = so_digitos(novo)
            if eh_nulo(novo, i_p):
                if MANTER_SE_BRUTA_VAZIA:
                    continue
                novo = None
            if ARREDONDAR_CASAS is not None:
                novo = arredonda(novo, ARREDONDAR_CASAS)
            novo = casa_tipo(novo, i_p)
            if norm_valor(novo) == norm_valor(valores[i_p]):
                iguais.add(i_p)          # coletado e confirmado: fica cinza
                continue
            if travar_area and i_p in (idx_area_ctrl if idx_trava is None
                                        else idx_trava):
                mantidas.add(i_p)
                continue
            valores[i_p] = novo
            pintar.add(i_p)
            if contar:
                contagem_por_col[nome] = contagem_por_col.get(nome, 0) + 1
            alterou = True
        # responsavel "." / "0" na coleta nao conta como responsavel informado
        if (doc_resp and not eh_nulo(rb[doc_resp[0]], doc_resp[2])
                and vazio(rb[doc_resp[1]])):
            _ir, _id, _ip, _idp = doc_resp[2:]
            _doc = valores[_idp]
            if _id in idx_so_digitos:      # o 15 pode ter vindo da coleta com pontos
                _doc = casa_tipo(so_digitos(_doc), _id)
            if _mesmo_nome(valores[_ir], valores[_ip]) and not vazio(_doc):
                if norm_valor(valores[_id]) != norm_valor(_doc):
                    valores[_id] = _doc
                    pintar.add(_id)
                    alterou = True
                if contar:
                    n_doc_resp[0] += 1
            elif contar and _ir in pintar:
                n_doc_resp[1] += 1
        return alterou

    n_terreno_unidade = 0
    cads_unificados_ok = set()
    i_mat_u = indice_de(cab_saida, COLUNA_MATRICULA, ABA_PREFEITURA, obrigatoria=False)
    i_area_u = indice_de(cab_saida, COLUNA_AREA_MATRICULA, ABA_PREFEITURA,
                         obrigatoria=False)
    n_baixados_lista = 0
    n_baixar_multi = 0
    for row in dados_pref:
        valores = [None] * n_col_dados
        for j in range(n_pref_col):
            valores[j + desloca] = row[j]
        chave = chave_de(valores, k_pref, u_pref)
        pintar, mantidas, iguais = set(), set(), set()
        status, vinculo, origem, coleta = "", "", "", ""

        # marcacao da edificacao compartilhada (por registro cadastro+unidade)
        g = grupos_multi.get(edif_do_registro.get(chave, "")) if chave else None
        if g is not None:
            vinculo = ("area mantida (matricula)" if g["tem_matricula"]
                       else "complementar")
            origem = g["origem"]

        _baixado = norm_chave(valores[k_pref]) in cads_baixados
        if _baixado and chave and chave in idx_bruta:
            usados.add(chave)
            coleta_em_baixado.append(
                (norm_chave(valores[k_pref]),
                 valores[u_pref] if CHAVE_COM_UNIDADE else "",
                 via_da_linha.get(id(idx_bruta[chave]), "")))
        if chave and chave in idx_bruta and not _baixado:
            usados.add(chave)
            rb = idx_bruta[chave]
            coleta = via_da_linha.get(id(rb), "")
            travar = g is not None                 # nos dois casos a area e mantida
            if id(rb) in edif_varios_cads:
                travar = True
                if not vinculo:
                    vinculo = "area mantida (edificacao de varios cadastros)"
            _membro = norm_chave(valores[k_pref]) in cads_condominio
            if _membro:
                _ok_ap = aplica(valores, rb, pintar, mantidas, iguais, True,
                                idx_trava=idx_trava_cond)
            else:
                _ok_ap = aplica(valores, rb, pintar, mantidas, iguais, travar)
            if (chave in unidade_terreno and CHAVE_COM_UNIDADE
                    and not (oc_bru is not None
                             and norm_chave(rb[oc_bru]) in baldios)):
                valores[u_pref] = casa_tipo(unidade_terreno[chave], u_pref)
                pintar.add(u_pref)
                _ok_ap = True
                n_terreno_unidade += 1
                coleta += "  [era terreno: ganhou a unidade %s]" % unidade_terreno[chave]
            if _ok_ap:
                n_atualizadas += 1
                n_celulas += len(pintar)
                status = "atualizado em campo"
            else:
                status = "campo igual a prefeitura"
            if vinculo == "complementar" and COMPLEMENTAR_MOSTRA_CODIGO:
                origem = codigo_linha(valores)   # so o codigo da propria linha
                n_codigos += 1
            n_travadas += len(mantidas)
            n_iguais += len(iguais)
            if comp_pares:
                dados_comp.append((valores[k_pref],
                                   valores[u_pref] if CHAVE_COM_UNIDADE else "",
                                   coleta,
                                   [(valor_pref_comp(row, ip, valores[k_pref]),
                                     valor_coleta(rb, ib, valores[k_pref]))
                                    for ib, ip, _a, _b, _l, _n1, _n2
                                    in comp_pares]))
            cads_vinculados.add(norm_chave(valores[k_pref]))
            if pintar - idx_colunas_novas:
                cads_atualizados.add(norm_chave(valores[k_pref]))
                n_linhas_mudaram += 1
            n_celulas_uuid += len(pintar & idx_colunas_novas)
        if ABA_PENDENCIAS:
            unidades_do_cad.setdefault(norm_chave(valores[k_pref]), []).append(
                (valores[u_pref] if CHAVE_COM_UNIDADE else "",
                 bool(status), coleta,
                 row[i_prop_p] if i_prop_p is not None else "",
                 row[i_area_p] if i_area_p is not None else ""))
        tipo_linha = "normal"
        _cad_l = norm_chave(valores[k_pref])
        if _cad_l in cads_baixar_multi and not _baixado:
            if norm_cab(MULTI_CADASTRO_MODO) == "baixar":
                tipo_linha = "baixar"
                status = ("BAIXAR - a coleta traz este cadastro junto com o %s,"
                          " que ficou com o lote" % cads_baixar_multi[_cad_l])
            else:
                tipo_linha = "complementar"
                status = ("COMPLEMENTAR - a coleta traz este cadastro junto com"
                          " o %s, que e o principal" % cads_baixar_multi[_cad_l])
            vinculo = "complementar do principal"
            origem = cads_baixar_multi[_cad_l]
            n_baixar_multi += 1
        elif _baixado:
            tipo_linha = "baixar"
            if _cad_l in cancelado_para:
                status = ("BAIXAR - cadastro cancelado, unificado ao %s"
                          % cancelado_para[_cad_l])
            elif _cad_l in cancelados_sem_destino:
                status = "BAIXAR - cadastro cancelado, sem unificacao"
            else:
                status = "BAIXAR - cadastro baixado"
            vinculo = origem = ""
            n_baixados_lista += 1
        elif (_cad_l in cads_para_baixar
                and (not CHAVE_COM_UNIDADE or not vazio(valores[u_pref]))):
            tipo_linha = "baixar"
            status = "BAIXAR - o lote hoje e baldio"
            n_baixar += 1
        elif _cad_l in cads_condominio:
            tipo_linha = "condominio_membro"
            status = ("CONDOMINIO %s - lote e classificacoes atualizados, area"
                      " mantida" % cads_condominio[_cad_l]
                      if chave and chave in idx_bruta else
                      "CONDOMINIO - ver a linha do %s" % cads_condominio[_cad_l])
            vinculo = "membro do condominio"
            origem = cads_condominio[_cad_l]
            n_membros_cond += 1
        elif _cad_l in cads_nao_atualizar:
            tipo_linha = "nao_atualizar"
            status = ("NAO ATUALIZAR - outro cadastro do mesmo lote recebeu o dado"
                      " de campo")
            n_nao_atualiza += 1
        if _cad_l in unificados and not _baixado:
            _mats = [m for c in unificados[_cad_l] for m in cancel_mats.get(c, [])]
            _tem = {norm_chave(x) for x in re.split(r"[/;,]", str(valores[i_mat_u] or ""))}
            _novas = [m for m in dict.fromkeys(_mats) if norm_chave(m) not in _tem]
            if _novas:
                valores[i_mat_u] = " / ".join(
                    ([str(valores[i_mat_u]).strip()] if not vazio(valores[i_mat_u])
                     else []) + _novas)
                pintar.add(i_mat_u)
            _soma = sum(cancel_area.get(c, 0.0) for c in unificados[_cad_l])
            if _soma:
                valores[i_area_u] = round((num(valores[i_area_u]) or 0.0) + _soma, 2)
                pintar.add(i_area_u)
            status = (status + "; " if status else "") + (
                "unificacao: recebeu matricula e area do(s) cadastro(s) %s"
                % ", ".join(unificados[_cad_l]))
            cads_unificados_ok.add(_cad_l)
        saida.append((valores, pintar, mantidas, iguais, status, vinculo, origem,
                      coleta, tipo_linha))

    # linhas da bruta que nao acharam vinculo
    #   tem numero de cadastro  -> mantem o numero, e unidade nova
    #   nao tem numero nenhum   -> TEXTO_CADASTRO_NOVO
    novos = novos_sem_numero = unid_novas = cad_inexistente = herdados = 0
    if ADICIONAR_NOVOS:
        for row, c in linhas_novas:
            valores = [None] * n_col_dados
            mantidas = set()
            aplica(valores, row, set(), mantidas, set(), False, contar=False)
            if CHAVE_COM_UNIDADE:
                valores[u_pref] = casa_tipo(row[u_bru], u_pref)
            if c is None:
                valores[k_pref] = TEXTO_CADASTRO_NOVO
                novos_sem_numero += 1
                preenche_cadastro_novo(valores, row)
                coleta = "sem numero de cadastro na coleta"
                status = "cadastro novo (sem numero na coleta)"
            else:
                _txt_cad = str(row[k_bru]).strip()
                valores[k_pref] = casa_tipo(c[0], k_pref)
                coleta = "%s: %s" % (COLUNA_CHAVE_BRUTA,
                                     "/".join(x for x in c if x))
                if _txt_cad in escolha_multi:
                    coleta += ("  [a coleta trouxe %s; ficou o %s]"
                               % (_txt_cad, c[0]))
                if c[0] in cads_pref:
                    unid_novas += 1
                    cads_unid_nova.add(c[0])
                    status = "unidade nova de cadastro existente"
                else:
                    cad_inexistente += 1
                    cads_novos.add(c[0])
                    preenche_cadastro_novo(valores, row)
                    status = "cadastro nao existe na prefeitura"
            if c is not None:
                for _is, _ip in herdadas:          # herda do proprio cadastro
                    if not vazio(valores[_is]):
                        continue
                    for _pr in linhas_do_cadastro.get(c[0], ()):
                        if not vazio(_pr[_ip]):
                            valores[_is] = _pr[_ip]
                            herdados += 1
                            break
            for _i, _txt in padrao_linha_nova.items():
                if vazio(valores[_i]):
                    valores[_i] = _txt
            saida.append((valores, set(range(n_col_dados)), mantidas, set(),
                          status, "", "", coleta, "nova"))
            novos += 1

    # linha PRINCIPAL de cada edificacao compartilhada sem matricula
    principais = 0
    i_ue = indice_de(cab_saida, COLUNA_UUID_EDIFICACAO, ABA_PREFEITURA,
                     obrigatoria=False) if REGRA_EDIFICACAO else None
    for ue, g in grupos_principal.items():
        valores = [None] * n_col_dados
        rb = g["linha"]
        if COPIAR_DADOS_NO_PRINCIPAL:
            aplica(valores, rb, set(), set(), set(), False, contar=False)
        if i_ue is not None:
            valores[i_ue] = rb[e_bru]
        cad_unico = list(g["cadastros"])[0] if len(g["cadastros"]) == 1 else None
        for _c in g["cadastros"]:
            cads_principal.add(_c)
        if PRINCIPAL_HERDA_CADASTRO_UNICO and cad_unico:
            valores[k_pref] = casa_tipo(cad_unico, k_pref)
        elif len(g["cadastros"]) > 1:
            valores[k_pref] = " / ".join(sorted(g["cadastros"],
                                                key=lambda x: (num(x) or 0, x)))
        valores[i_area_dest] = casa_tipo(g["area"], i_area_dest)
        if COPIAR_DADOS_NO_PRINCIPAL:
            if cad_unico:                       # herda o que faltou do cadastro
                for _is, _ip in herdadas:
                    if vazio(valores[_is]):
                        for _pr in linhas_do_cadastro.get(cad_unico, ()):
                            if not vazio(_pr[_ip]):
                                valores[_is] = _pr[_ip]
                                break
            for _i, _txt in padrao_linha_nova.items():
                if vazio(valores[_i]):
                    valores[_i] = _txt
        saida.append((valores, set(range(n_col_dados)), set(), set(),
                      "principal (edificacao compartilhada)", "principal",
                      g["origem"], via_da_linha.get(id(rb), ""), "principal"))
        principais += 1

    # edificacao dividida por varios cadastros: uma linha PRINCIPAL com todos os
    # numeros, ja que nao da para dizer quanto da area e de cada cadastro
    principais_multi = 0
    if EDIFICACAO_MULTICADASTRO_TRAVA_AREA and PRINCIPAL_EDIF_MULTICADASTRO:
        _ues_feitos = set(grupos_principal)
        for _k, _cs in sorted(edif_grupo_cads.items()):
            if _k in _ues_feitos:
                continue
            if any(_c in cads_condominio for _c in _cs):
                continue
            rb = edif_grupo_linha[_k]
            valores = [None] * n_col_dados
            aplica(valores, rb, set(), set(), set(), False, contar=False)
            if i_ue is not None and e_bru is not None:
                valores[i_ue] = rb[e_bru]
            valores[k_pref] = " / ".join(sorted(_cs, key=lambda x: (num(x) or 0, x)))
            if CHAVE_COM_UNIDADE:
                valores[u_pref] = None
            if i_area_dest is not None and a_bru is not None \
                    and not vazio(rb[a_bru]):
                valores[i_area_dest] = casa_tipo(rb[a_bru], i_area_dest)
            for _i, _txt in padrao_linha_nova.items():
                if vazio(valores[_i]):
                    valores[_i] = _txt
            saida.append((valores, set(range(n_col_dados)), set(), set(),
                          "principal (edificacao de %d cadastros)" % len(_cs),
                          "principal", ", ".join(sorted(_cs)),
                          via_da_linha.get(id(rb), ""), "principal"))
            principais_multi += 1

    # linhas da coleta que apontavam para outro cadastro do mesmo lote:
    # viram unidade nova do cadastro escolhido
    if linhas_transferidas:
        _prox = dict(cad_max_unid)
        for row, esc, fora, motivo, via in linhas_transferidas:
            valores = [None] * n_col_dados
            mantidas = set()
            aplica(valores, row, set(), mantidas, set(), False, contar=False)
            valores[k_pref] = casa_tipo(esc, k_pref)
            if CHAVE_COM_UNIDADE:
                if UNIDADE_TRANSFERIDA == "proxima":
                    _prox[esc] = _prox.get(esc, 0) + 1
                    valores[u_pref] = _prox[esc]
                elif UNIDADE_TRANSFERIDA == "coleta":
                    valores[u_pref] = casa_tipo(row[u_bru], u_pref)
                else:
                    valores[u_pref] = None
            for _is, _ip in herdadas:
                if vazio(valores[_is]):
                    for _pr in linhas_do_cadastro.get(esc, ()):
                        if not vazio(_pr[_ip]):
                            valores[_is] = _pr[_ip]
                            break
            for _i, _txt in padrao_linha_nova.items():
                if vazio(valores[_i]):
                    valores[_i] = _txt
            saida.append((valores, set(range(n_col_dados)), mantidas, set(),
                          "unidade nova (veio do cadastro %s, mesmo lote)"
                          % ", ".join(fora), "", "",
                          via + "  [transferida para o %s: %s]" % (esc, motivo),
                          "nova"))
            n_transferidas += 1
            novos += 1

    # linhas do CONDOMINIO
    n_prop_cond = 0
    if linhas_condominio:
        _a_b = indice_de(cab_bru, ORIGEM_AREA_PRINCIPAL, ABA_BRUTA, obrigatoria=False)
        _i_ad = indice_de(cab_saida, DESTINO_AREA_PRINCIPAL, ABA_PREFEITURA,
                          obrigatoria=False)
        _i_mat_s = indice_de(cab_saida, COLUNA_MATRICULA, ABA_PREFEITURA,
                             obrigatoria=False)
        for row, cid, papel, via in linhas_condominio:
            inf = cond_info[cid]
            valores = [None] * n_col_dados
            aplica(valores, row, set(), set(), set(), False, contar=False)
            valores[k_pref] = TEXTO_CONDOMINIO
            # matricula do condominio = as da prefeitura dos membros; a da coleta
            # nao entra nessa regra
            if _i_mat_s is not None:
                _ms = set()
                for _c in inf["cads"]:
                    _ms |= cad_matriculas.get(_c, set())
                valores[_i_mat_s] = " / ".join(sorted(_ms)) if _ms else None
            if papel == "condominio":
                if CHAVE_COM_UNIDADE:
                    valores[u_pref] = None
                if _a_b is not None and _i_ad is not None and not vazio(row[_a_b]):
                    valores[_i_ad] = casa_tipo(row[_a_b], _i_ad)
                hist = "%s: %s" % (cid, ", ".join(
                    "%s (matr %s)" % (c, "/".join(sorted(cad_matriculas.get(c, [])))
                                      or "-")
                    for c in sorted(inf["cads"], key=lambda x: (num(x) or 0, x))))
                status = ("CONDOMINIO - %d cadastros com matriculas diferentes"
                          % len(inf["cads"]))
                vinc = "condominio"
            else:
                if CHAVE_COM_UNIDADE:
                    valores[u_pref] = casa_tipo(row[u_bru], u_pref)
                hist = cid
                if papel == "unidade_mesma":
                    # a area da edificacao fica so na linha do condominio
                    if _i_ad is not None:
                        valores[_i_ad] = None
                    status = "unidade do condominio"
                else:
                    status = ("unidade nova do condominio (outra edificacao do"
                              " lote)")
                vinc = "unidade do condominio"
            # coleta sem valor: herda o valor PREDOMINANTE entre os cadastros
            # membros (cada cadastro vota uma vez; empate -> mais linhas)
            for _is, _ip in herdadas:
                if not eh_nulo(valores[_is], _is):
                    continue
                _voto, _linhas, _texto = {}, {}, {}
                for _c in inf["cads"]:
                    _vistos = set()
                    for _pr in linhas_do_cadastro.get(_c, ()):
                        if eh_nulo(_pr[_ip], _is):
                            continue
                        _t = str(_pr[_ip]).strip()
                        _k = " ".join(_sem_acento(_t).upper().split())
                        _texto.setdefault(_k, _t)
                        _linhas[_k] = _linhas.get(_k, 0) + 1
                        if _k not in _vistos:
                            _vistos.add(_k)
                            _voto[_k] = _voto.get(_k, 0) + 1
                if _voto:
                    _k = max(sorted(_voto), key=lambda k: (_voto[k], _linhas[k]))
                    valores[_is] = _texto[_k]
                    n_prop_cond += 1
            for _i, _txt in padrao_linha_nova.items():
                if vazio(valores[_i]):
                    valores[_i] = _txt
            saida.append((valores, set(range(n_col_dados)), set(), set(),
                          status, vinc, hist, via, "condominio"))
            n_linhas_cond += 1

    # linhas de LOTE BALDIO: cadastro novo no mesmo numero, sem edificacao
    baldio_novos = baldio_repetidos = 0
    if REGRA_LOTE_BALDIO and linhas_baldio:
        herdadas_baldio = list(herdadas)
        for _nome in COLUNAS_HERDADAS_NO_BALDIO:
            _is = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
            _ip = indice_de(cab_pref,  _nome, ABA_PREFEITURA, obrigatoria=False)
            if _is is not None and _ip is not None:
                herdadas_baldio.append((_is, _ip))
            else:
                print("    [aviso] coluna herdada no baldio nao encontrada: %s" % _nome)
        idx_nulas = []
        for _nome in COLUNAS_NULAS_BALDIO:
            _i = indice_de(cab_saida, _nome, ABA_PREFEITURA, obrigatoria=False)
            if _i is None:
                print("    [aviso] coluna do baldio nao encontrada: %s" % _nome)
            else:
                idx_nulas.append(_i)
        cads_baldio_feitos = set()
        for row in linhas_baldio:
            cad = cad_res(row)
            if not cad and ec_bru is not None and not vazio(row[ec_bru]):
                _ok, _err = parse_cadastro_edificacao(row[ec_bru])
                if _ok:
                    cad = _ok[0][0]
            if cad:                        # uma linha de baldio por cadastro
                if cad in cads_baldio_feitos:
                    baldio_repetidos += 1
                    continue
                cads_baldio_feitos.add(cad)
            valores = [None] * n_col_dados
            aplica(valores, row, set(), set(), set(), False, contar=False)
            valores[k_pref] = casa_tipo(cad, k_pref) if cad else TEXTO_CADASTRO_NOVO
            if not cad or cad not in cads_pref:
                preenche_cadastro_novo(valores, row)
            if CHAVE_COM_UNIDADE:
                valores[u_pref] = UNIDADE_LOTE_BALDIO
            for _i in idx_nulas:
                valores[_i] = None
            if cad:
                for _is, _ip in herdadas_baldio:
                    if vazio(valores[_is]):
                        for _pr in linhas_do_cadastro.get(cad, ()):
                            if not vazio(_pr[_ip]):
                                valores[_is] = _pr[_ip]
                                break
            for _i, _txt in padrao_linha_nova.items():
                if vazio(valores[_i]):
                    valores[_i] = _txt
            ja_existe = cad in cads_pref
            (cads_baldio_exist if ja_existe else cads_baldio_novo).add(
                cad or ("s/n %d" % baldio_novos))
            saida.append((valores, set(range(n_col_dados)), set(), set(),
                          "lote baldio (%s)" % ("cadastro ja existe na prefeitura"
                                                if ja_existe else "cadastro novo"),
                          "lote baldio", "", "%s: %s" % (COLUNA_OCUPACAO_BRUTA,
                                                         norm_chave(row[oc_bru])),
                          "baldio"))
            baldio_novos += 1

    # linhas com o cadastro da edificacao mal preenchido -> so para conferir
    for row, valor, erros in linhas_erro:
        valores = [None] * n_col_dados
        aplica(valores, row, set(), set(), set(), False, contar=False)
        valores[k_pref] = None
        if CHAVE_COM_UNIDADE:
            valores[u_pref] = casa_tipo(row[u_bru], u_pref)
        saida.append((valores, set(), set(), set(),
                      "CONFERIR: %s invalido" % COLUNA_CADASTRO_EDIFICACAO, "",
                      "; ".join(erros), "%s: %s" % (COLUNA_CADASTRO_EDIFICACAO, valor),
                      "erro"))

    # unidades sem coleta em cadastro que teve outra unidade atualizada:
    # entram em rosa na aba, para conferir na bruta
    pendentes_prev = 0
    chaves_sem_coleta = set()
    for _cad, _lst in unidades_do_cad.items():
        _com = [1 for _u, ok, _v, _p, _a in _lst if ok]
        _sem = [_u for _u, ok, _v, _p, _a in _lst if not ok]
        if _com and _sem:
            pendentes_prev += len(_sem)
            for _u in _sem:
                chaves_sem_coleta.add((_cad, norm_chave(_u)))
    if chaves_sem_coleta:
        _nova = []
        for _e in saida:
            if _e[8] == "normal":
                _k = (norm_chave(_e[0][k_pref]),
                      norm_chave(_e[0][u_pref]) if CHAVE_COM_UNIDADE else "")
                if _k in chaves_sem_coleta:
                    _e = (_e[0], _e[1], _e[2], _e[3],
                          "BAIXAR - unidade sem dado de campo", _e[5], _e[6],
                          _e[7], "sem_coleta")
            _nova.append(_e)
        saida = _nova

    # so o que teve interacao com a coleta
    fora_da_saida = 0
    if SO_COM_INTERACAO:
        _com = set()
        for _e in saida:
            if _e[8] != "normal" or _e[4]:
                _com.add(norm_chave(_e[0][k_pref]))
        _nova = []
        for _e in saida:
            if _e[8] != "normal" or _e[4]:
                _nova.append(_e)
            elif (norm_cab(SO_COM_INTERACAO) == "cadastro"
                  and norm_chave(_e[0][k_pref]) in _com):
                _nova.append(_e)
            else:
                fora_da_saida += 1
        saida = _nova

    # ------------------------------------------------------------- escrita
    print("Gravando a aba '%s'..." % ABA_SAIDA)
    if ABA_SAIDA in wb.sheetnames:
        del wb[ABA_SAIDA]
    ws = wb.create_sheet(ABA_SAIDA)

    f_atu  = PatternFill("solid", fgColor=COR_ATUALIZADO)
    f_novo = PatternFill("solid", fgColor=COR_LINHA_NOVA)
    f_prin = PatternFill("solid", fgColor=COR_PRINCIPAL)
    f_mant = PatternFill("solid", fgColor=COR_MANTIDA)
    f_erro = PatternFill("solid", fgColor=COR_ERRO)
    f_bald = PatternFill("solid", fgColor=COR_BALDIO)
    f_baix = PatternFill("solid", fgColor=COR_BAIXAR)
    f_semc = PatternFill("solid", fgColor=COR_SEM_COLETA)
    f_naoat = PatternFill("solid", fgColor=COR_NAO_ATUALIZA)
    f_cond = PatternFill("solid", fgColor=COR_CONDOMINIO)
    f_condm = PatternFill("solid", fgColor=COR_CONDOMINIO_MEMBRO)
    f_igual = PatternFill("solid", fgColor=COR_IGUAL)
    f_cab  = PatternFill("solid", fgColor=COR_CABECALHO)

    lin_cab = 1
    if COPIAR_LINHAS_ACIMA and acima_pref:
        for r, linha in enumerate(acima_pref, start=1):
            for j in range(n_pref_col):
                if linha[j] is not None and str(linha[j]).strip() != "":
                    ws.cell(row=r, column=j + 1 + desloca, value=linha[j])
        lin_cab = len(acima_pref) + 1

    def classifica(tipo, status):
        _s = _sem_acento(status or "").lower()
        if tipo == "baixar":
            return "CADASTRO CANCELADO" if "cancelado" in _s else "UNIDADE BAIXADA"
        if tipo == "complementar":
            return "COMPLEMENTAR"
        if tipo == "nao_atualizar":
            return "NAO ATUALIZAR"
        if tipo == "condominio":
            return "UNIDADE DO CONDOMINIO" if "unidade" in _s else "CONDOMINIO"
        if tipo == "condominio_membro":
            return "MEMBRO DE CONDOMINIO"
        if tipo == "principal":
            return "PRINCIPAL"
        if tipo == "baldio":
            return "LOTE BALDIO"
        if tipo == "sem_coleta":
            return "UNIDADE BAIXADA"
        if tipo == "erro":
            return "CONFERIR"
        if tipo == "nova":
            if "cadastro novo" in _s:
                return "CADASTRO NOVO"
            if "nao existe na prefeitura" in _s:
                return "CADASTRO NAO EXISTE NA PREFEITURA"
            return "UNIDADE NOVA"
        if "atualizado em campo" in _s:
            return "ATUALIZADO"
        if "igual a prefeitura" in _s:
            return "CONFIRMADO EM CAMPO"
        if "unificacao" in _s:
            return "ATUALIZADO"
        return "UNIDADE BAIXADA"

    _mapa_sit = {norm_cab(k): v for k, v in MAPA_SITUACAO.items()}

    def agrupa_situacao(detalhe):
        return _mapa_sit.get(norm_cab(detalhe or ""), detalhe)

    extras = [c for c in (COLUNA_CLASSIFICACAO,
                          COLUNA_CLASSIF_HIST,
                          COLUNA_VINCULO if REGRA_EDIFICACAO else "",
                          COLUNA_ORIGEM if REGRA_EDIFICACAO else "",
                          COLUNA_CHAVE_COLETA,
                          COLUNA_STATUS) if c]
    cabecalho = list(cab_saida_orig) + extras
    f_sep = PatternFill("solid", fgColor=COR_SEPARADOR)
    for j, v in enumerate(cabecalho, start=1):
        c = ws.cell(row=lin_cab, column=j, value=v)
        c.font = Font(bold=True)
        c.fill = f_sep if v == COLUNA_SEPARADOR else f_cab
    ws.freeze_panes = ws.cell(row=lin_cab + 1, column=1).coordinate
    ws.auto_filter.ref = "A%d:%s%d" % (lin_cab, get_column_letter(len(cabecalho)),
                                       lin_cab + len(saida))

    r = lin_cab
    for (valores, pintar, mantidas, iguais, status, vinculo, origem, coleta,
         tipo) in saida:
        r += 1
        fundo = {"nova": f_novo, "principal": f_prin, "erro": f_erro,
                 "baldio": f_bald, "baixar": f_baix,
                 "sem_coleta": f_baix, "nao_atualizar": f_naoat,
                 "complementar": f_mant,
                 "condominio": f_cond}.get(tipo)
        for j in range(n_col_dados):
            v = valores[j]
            if (v is None and fundo is None and j not in pintar
                    and j not in mantidas and j not in iguais):
                continue                       # celula vazia: nem escreve
            c = ws.cell(row=r, column=j + 1, value=v)
            if fundo is not None:
                c.fill = fundo
            elif j in mantidas:
                c.fill = f_mant
            elif j in pintar:
                c.fill = f_atu
            elif j in iguais:
                c.fill = f_igual
            if tipo == "condominio_membro" and j == k_pref:
                c.fill = f_condm
        col = n_col_dados
        _det = classifica(tipo, status)
        for texto in ([agrupa_situacao(_det)] if COLUNA_CLASSIFICACAO else []) + \
                     ([_det] if COLUNA_CLASSIF_HIST else []) + \
                     ([vinculo, origem] if REGRA_EDIFICACAO else []) + \
                     ([coleta] if COLUNA_CHAVE_COLETA else []) + \
                     ([status] if COLUNA_STATUS else []):
            col += 1
            if texto or fundo is not None:
                c = ws.cell(row=r, column=col, value=texto or None)
                if fundo is not None:
                    c.fill = fundo

    base = r + 2
    ws.cell(row=base, column=1, value="LEGENDA").font = Font(bold=True)
    for i, (fill, txt) in enumerate([
            (f_atu,  "celula atualizada com o dado da bruta"),
            (f_igual, "a coleta trouxe o mesmo valor que ja estava na prefeitura"),
            (f_mant, "area mantida da prefeitura pela regra da edificacao"),
            (f_novo, "registro novo, existe so na bruta"),
            (f_prin, "cadastro PRINCIPAL criado para a edificacao compartilhada"),
            (f_baix, "BAIXAR: unidade que existe na prefeitura e hoje e lote baldio"),
            (f_cond, "CONDOMINIO: linha do condominio (e unidades novas dele)"),
            (f_condm, "cadastro que forma um condominio (na coluna do cadastro):"
                      " lote e classificacoes atualizados, area mantida"),
            (f_naoat, "NAO ATUALIZAR: outro cadastro do mesmo lote recebeu o dado"
                      " de campo"),
            (f_baix, "unidade sem dado de campo em cadastro que teve outra"
                     " atualizada: conferir na bruta"),
            (f_bald, "lote baldio hoje: linha nova sem os campos de edificacao"),
            (f_erro, "CONFERIR: %s mal preenchido, a linha nao foi vinculada"
                     % COLUNA_CADASTRO_EDIFICACAO)], start=1):
        ws.cell(row=base + i, column=1, value="  ").fill = fill
        ws.cell(row=base + i, column=2, value=txt)
    ws.cell(row=base + 13, column=2,
            value="gerado em %s" % datetime.now().strftime("%d/%m/%Y %H:%M"))

    # ------------------------------------------------- aba de comparacao
    if ABA_COMPARACAO and comp_pares:
        print("Gravando a aba '%s'..." % ABA_COMPARACAO)
        if ABA_COMPARACAO in wb.sheetnames:
            del wb[ABA_COMPARACAO]
        wc = wb.create_sheet(ABA_COMPARACAO)
        neg = Font(bold=True)
        f_dif = PatternFill("solid", fgColor=COR_MANTIDA)

        por_cadastro = norm_cab(COMPARACAO_POR) == "cadastro"
        if por_cadastro:                     # uma linha por cadastro
            _vistos, _uma_por_cad = set(), []
            for _reg in dados_comp:
                _k = norm_chave(_reg[0])
                if _k in _vistos:
                    continue
                _vistos.add(_k)
                _uma_por_cad.append(_reg)
            dados_comp = _uma_por_cad

        # ---- estatistica de cada par
        estat = []
        for k, (_ib, _ip, nb, np_, lim, _n1, _n2) in enumerate(comp_pares):
            difs = []
            for _c, _u, _v, vals in dados_comp:
                a, b = num(vals[k][0]), num(vals[k][1])
                if a is None or b is None or a == 0:
                    continue
                difs.append((b - a) / a)
            comparados = len(difs)
            fora = sum(1 for d in difs if abs(d) > lim)
            fora50 = sum(1 for d in difs if abs(d) > 0.50)
            media = sum(difs) / comparados if comparados else 0
            ord_ = sorted(difs)
            mediana = (ord_[comparados // 2] if comparados % 2
                       else (ord_[comparados // 2 - 1] + ord_[comparados // 2]) / 2) \
                if comparados else 0
            maior = max(difs, key=abs) if difs else 0
            estat.append((nb, np_, lim, comparados, fora, fora50, mediana, media,
                          maior))

        r = 1
        wc.cell(row=r, column=1, value="RESUMO DA RODADA").font = Font(bold=True, size=13)
        r += 1
        wc.cell(row=r, column=1,
                value="gerado em %s" % datetime.now().strftime("%d/%m/%Y %H:%M"))
        r += 2
        tot_cad_pref = len({norm_chave(_r[i_cad_p]) for _r in dados_pref})
        tocados = (cads_vinculados | cads_unid_nova | cads_novos
                   | cads_baldio_exist | cads_baldio_novo)
        so_confirmados = cads_vinculados - cads_atualizados
        linhas_resumo = [
            ("A) O QUE ACONTECEU COM CADA LINHA DA COLETA", "",
             "a bruta tem %d linhas" % len(dados_bru)),
            ("  vinculou a uma unidade que ja existia",
             n_via_edif + n_via_lote + n_baldio_atualiza,
             "achou o par cadastro+unidade na prefeitura e atualizou aquela linha"),
            ("      pelo %s" % COLUNA_CADASTRO_EDIFICACAO, n_via_edif,
             "a coleta informou cadastro.unidade"),
            ("      pelo %s + unidade" % COLUNA_CHAVE_BRUTA, n_via_lote,
             "caiu na segunda opcao da cascata"),
            ("      baldio em cadastro que ja era terreno", n_baldio_atualiza,
             "nao cria linha nova: atualiza a linha de terreno que ja existe"),
            ("      terreno que ganhou edificacao", n_terreno_unidade,
             "a prefeitura so tinha a linha de terreno (unidade vazia): a linha foi"
             " atualizada e recebeu o numero da unidade da coleta"),
            ("  virou linha nova", novos,
             "nao achou o par na prefeitura, entao entrou como linha nova (verde)"),
            ("  entrou num condominio", n_coleta_cond,
             "linha da coleta de edificacao com cadastros de matriculas"
             " diferentes, ou de outra edificacao do mesmo lote"),
            ("  virou linha de lote baldio", baldio_novos,
             "%s = %s: nao atualiza as unidades antigas, cria linha sem edificacao"
             % (COLUNA_OCUPACAO_BRUTA, "/".join(VALOR_OCUPACAO_BALDIO))),
            ("  ficou para conferir", len(linhas_erro),
             "%s mal preenchido: a linha nao foi usada" % COLUNA_CADASTRO_EDIFICACAO),
            ("  = total conferido", n_via_edif + n_via_lote + n_baldio_atualiza
             + novos + baldio_novos + len(linhas_erro) + baldio_repetidos
             + n_coleta_cond,
             "as %d linhas da bruta (inclui %d repetidas de baldio no mesmo cadastro)"
             % (len(dados_bru), baldio_repetidos)),
            ("", "", ""),
            ("B) CADASTROS ATUALIZADOS", "",
             "conta CADASTRO, nao linha: um cadastro com 5 unidades conta 1"),
            ("  cadastros que receberam dado de campo", len(cads_vinculados),
             "pelo menos uma unidade do cadastro vinculou com a coleta"),
            ("      com algum valor alterado", len(cads_atualizados),
             "a coleta trouxe algo diferente do que ja estava na prefeitura"),
            ("      so confirmados, nada mudou", len(so_confirmados),
             "a coleta bateu com a prefeitura em tudo (fora os uuid, que sao"
             " colunas novas e mudam sempre)"),
            ("  cadastros que ganharam unidade nova", len(cads_unid_nova),
             "o cadastro ja existia e a coleta trouxe uma unidade que nao estava la"),
            ("  cadastros de edificacao compartilhada", len(cads_principal),
             "entraram no grupo que gerou linha PRINCIPAL (azul)"),
            ("  cadastros que viraram lote baldio", len(cads_baldio_exist),
             "hoje o terreno esta vazio: as unidades antigas ficaram intocadas"),
            ("      unidades a BAIXAR nesses cadastros", n_baixar,
             "linhas em vermelho: a prefeitura precisa dar baixa nelas"),
            ("  condominios", len(cond_info),
             "edificacao com cadastros de matriculas diferentes: linha nova de"
             " condominio, cadastros membros intocados"),
            ("      cadastros membros de condominio", len(cads_condominio),
             "%d linhas da prefeitura marcadas; %d linhas de condominio criadas"
             % (n_membros_cond, n_linhas_cond)),
            ("  lotes com mais de um cadastro", len(escolhido_do_lote),
             "so um recebe o dado: o que tem matricula, ou o de maior area"
             " construida"),
            ("      cadastros nao atualizados (vermelho escuro)",
             len(cads_nao_atualizar), "%d linhas da prefeitura marcadas NAO"
             " ATUALIZAR" % n_nao_atualiza),
            ("      linhas da coleta transferidas", n_transferidas,
             "viraram unidade nova do cadastro escolhido"),
            ("  total de cadastros tocados", len(tocados - cads_novos
                                                 - cads_baldio_novo),
             "cadastros da prefeitura que aparecem de alguma forma nesta rodada"),
            ("  cadastros da prefeitura sem nenhum contato",
             tot_cad_pref - len(tocados - cads_novos - cads_baldio_novo),
             "de %d cadastros na prefeitura, esses nao foram a campo" % tot_cad_pref),
            ("", "", ""),
            ("C) CADASTROS NOVOS", "",
             "o que nao existe na prefeitura e precisa ser criado"),
            ("  cadastro com numero que nao existe na prefeitura", len(cads_novos),
             "a coleta trouxe um numero de cadastro que a base nao tem (%d linhas)"
             % cad_inexistente),
            ("  lote baldio sem numero de cadastro", len(cads_baldio_novo),
             "terreno vazio que a coleta achou e nao conseguiu numerar"),
            ("  linhas sem numero de cadastro na coleta", novos_sem_numero,
             "entram como '%s'; cada uma precisa receber um numero"
             % TEXTO_CADASTRO_NOVO),
            ("  linhas PRINCIPAL de edificacao com varios cadastros",
             principais_multi,
             "a edificacao e dividida por mais de um cadastro: a linha traz todos"
             " os numeros juntos e a area coletada"),
            ("  linhas PRINCIPAL criadas", principais,
             "uma por edificacao compartilhada sem matricula"),
            ("  linhas de condominio criadas", n_linhas_cond,
             "a linha do condominio e as unidades dele"),
            ("  = linhas novas na aba '%s'" % ABA_SAIDA,
             novos + baldio_novos + principais + len(linhas_erro) + n_linhas_cond,
             "tudo que nao veio da prefeitura"),
            ("", "", ""),
            ("D) VOLUME DE ALTERACAO", "", "contagem de celulas"),
            ("  linhas (unidades) que mudaram algum valor", n_linhas_mudaram,
             "sem contar as que so ganharam os uuid"),
            ("  linhas que receberam algo da coleta", n_atualizadas,
             "inclui as que so ganharam uuid/uuid_lote/uuid_edificacao"),
            ("  celulas atualizadas no total", n_celulas,
             "celulas amarelas na aba '%s'" % ABA_SAIDA),
            ("      dessas, so os uuid", n_celulas_uuid,
             "colunas que nao existem na prefeitura, entao contam sempre"),
            ("  celulas conferidas com o mesmo valor", n_iguais,
             "celulas cinza: a coleta confirmou o que ja estava la"),
            ("  celulas de area mantidas pela regra", n_travadas,
             "celulas laranja: edificacao compartilhada com matricula"),
            ("  complementares com o codigo da propria linha", n_codigos,
             "a lista com os codigos do grupo fica so na linha principal"),
            ("  unidades sem coleta a conferir", pendentes_prev,
             "unidade sem dado de campo em cadastro que teve outra atualizada"),
            ("", "", ""),
            ("  linhas na aba '%s'" % ABA_SAIDA, len(saida),
             "so quem teve interacao com a coleta"),
            ("  linhas da prefeitura fora da aba", fora_da_saida,
             "cadastros que nao foram a campo, ficaram so na aba prefeitura"),
        ]
        for rot, val, expl in linhas_resumo:
            if rot:
                c = wc.cell(row=r, column=1, value=rot)
                if not rot.startswith(" "):
                    c.font = neg
                    c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
                elif not rot.startswith("      "):
                    c.font = Font(bold=True)
                if val != "":
                    wc.cell(row=r, column=2, value=val).font = (
                        neg if not rot.startswith("  ") or rot.startswith("  =")
                        else Font())
                if expl:
                    wc.cell(row=r, column=3, value=expl).font = Font(italic=True,
                                                                     color="666666")
            r += 1

        r += 1
        wc.cell(row=r, column=1,
                value="COMPARACAO DE AREAS (por %s)" % ("cadastro" if por_cadastro
                                                        else "unidade")).font = neg
        r += 1
        for j, t in enumerate(["par comparado", "limite de alerta",
                               "registros comparados", "acima do limite",
                               "acima de 50%", "diferenca mediana",
                               "diferenca media", "maior diferenca"], start=1):
            c = wc.cell(row=r, column=j, value=t)
            c.font = neg
            c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        for nb, np_, lim, comparados, fora, fora50, mediana, media, maior in estat:
            r += 1
            wc.cell(row=r, column=1, value="%s  x  %s"
                    % (nb.replace("soma:", "soma de "), np_))
            wc.cell(row=r, column=2, value=lim).number_format = "0.0%"
            wc.cell(row=r, column=3, value=comparados)
            wc.cell(row=r, column=4, value=fora)
            wc.cell(row=r, column=5, value=fora50)
            for j, v in ((6, mediana), (7, media), (8, maior)):
                wc.cell(row=r, column=j, value=v).number_format = "0.0%"

        # ---- tabela linha a linha
        r += 3
        lin_faixa = r
        lin_cab_c = r + 1
        cab = [COLUNA_CHAVE_PREFEITURA]
        if not por_cadastro:
            cab.append(COLUNA_UNIDADE_PREFEITURA)
        cab.append(COLUNA_CHAVE_COLETA)
        cab_tec = list(cab)
        for _ib, _ip, nb, np_, lim, _n1, _n2 in comp_pares:
            cab_tec += ["%s (prefeitura)" % np_,
                        "%s (coleta)" % nb.replace("soma:", "soma de "),
                        "dif m2 [%s]" % np_, "dif %% [%s]" % np_,
                        "alerta [%s]" % np_]
        for _ib, _ip, nb, np_, lim, nome_p, nome_c in comp_pares:
            cab += [nome_p, nome_c, NOME_DIF_M2, NOME_DIF_PC, NOME_ALERTA]
        n_fixas = len(cab) - 5 * len(comp_pares)
        for k, (_ib, _ip, nb, np_, lim, nome_p, nome_c) in enumerate(comp_pares):
            cor = CORES_COMPARACAO[k % len(CORES_COMPARACAO)]
            f_faixa = PatternFill("solid", fgColor=cor)
            base_col = n_fixas + 5 * k
            for j, t in enumerate([nome_p, nome_c, NOME_DIF_M2, NOME_DIF_PC,
                                   NOME_ALERTA], start=1):
                c = wc.cell(row=lin_faixa, column=base_col + j, value=t)
                c.font = neg
                c.fill = f_faixa
        r = lin_cab_c
        for j, t in enumerate(cab_tec, start=1):
            c = wc.cell(row=r, column=j, value=t)
            c.font = neg
            c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        for cad, uni, via, vals in dados_comp:
            r += 1
            wc.cell(row=r, column=1, value=cad)
            col = 1
            if not por_cadastro:
                col += 1
                wc.cell(row=r, column=col, value=uni)
            col += 1
            wc.cell(row=r, column=col, value=via)
            for (a_raw, b_raw), (_ib, _ip, _nb, _np, lim, _n1, _n2) \
                    in zip(vals, comp_pares):
                a, b = num(a_raw), num(b_raw)
                wc.cell(row=r, column=col + 1, value=a if a is not None else a_raw)
                wc.cell(row=r, column=col + 2, value=b if b is not None else b_raw)
                if a is not None and b is not None:
                    d = b - a
                    cd = wc.cell(row=r, column=col + 3, value=round(d, 2))
                    cd.number_format = "0.00"
                    if a != 0:
                        pc = wc.cell(row=r, column=col + 4, value=d / a)
                        pc.number_format = "0.0%"
                        if abs(d / a) > lim:
                            pc.fill = f_dif
                            cd.fill = f_dif
                            al = wc.cell(row=r, column=col + 5,
                                         value="ACIMA DE %g%%" % (lim * 100))
                            al.fill = f_dif
                            al.font = Font(bold=True)
                col += 5
        wc.auto_filter.ref = "A%d:%s%d" % (lin_cab_c, get_column_letter(len(cab)), r)
        wc.column_dimensions["A"].width = 52
        for j in range(2, len(cab) + 1):
            wc.column_dimensions[get_column_letter(j)].width = 18
        wc.column_dimensions["C"].width = 70   # coluna da explicacao do resumo

    # ------------------------------------- aba de unidades sem coleta
    pendentes = 0
    if ABA_PENDENCIAS and unidades_do_cad:
        print("Gravando a aba '%s'..." % ABA_PENDENCIAS)
        if ABA_PENDENCIAS in wb.sheetnames:
            del wb[ABA_PENDENCIAS]
        wp = wb.create_sheet(ABA_PENDENCIAS)
        neg = Font(bold=True)
        cabp = [COLUNA_CHAVE_PREFEITURA, "unidade SEM coleta",
                "unidades do cadastro que vincularam",
                "unidades do cadastro sem coleta",
                "total de unidades na prefeitura",
                "linhas da coleta nesse cadastro",
                "unidades na coleta",
                COLUNA_CADASTRO_EDIFICACAO + " na coleta",
                COLUNAS_HERDADAS_DO_CADASTRO[0] if COLUNAS_HERDADAS_DO_CADASTRO else "",
                "%s (prefeitura)" % DESTINO_AREA_PRINCIPAL,
                "vinculo usado nas que atualizaram"]
        for j, t in enumerate(cabp, start=1):
            c = wp.cell(row=1, column=j, value=t)
            c.font = neg
            c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        r = 1
        for cad, lst in sorted(unidades_do_cad.items(),
                               key=lambda x: (num(x[0]) if num(x[0]) is not None
                                              else float("inf"), x[0])):
            com = [u for u, ok, _v, _p, _a in lst if ok]
            sem = [u for u, ok, _v, _p, _a in lst if not ok]
            if not com or not sem:
                continue                       # so interessa cadastro misto
            via = ", ".join(sorted(set(v for _u, ok, v, _p, _a in lst if ok and v)))
            col = coleta_do_cad.get(cad, {})
            for u, ok, _v, prop, area in lst:
                if ok:
                    continue
                r += 1
                pendentes += 1
                for j, v in enumerate([
                        casa_tipo(cad, k_pref), u,
                        ", ".join(str(x) for x in com),
                        ", ".join(str(x) for x in sem),
                        len(lst), col.get("linhas", 0),
                        ", ".join(col.get("unidades", [])),
                        " | ".join(col.get("edif", [])),
                        prop, area, via], start=1):
                    wp.cell(row=r, column=j, value=v)
        wp.freeze_panes = "C2"
        wp.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cabp)), max(r, 1))
        for j, w in enumerate([16, 18, 30, 26, 14, 14, 16, 34, 34, 16, 46], start=1):
            wp.column_dimensions[get_column_letter(j)].width = w

    # ------------------------------------- aba de cadastros no mesmo lote
    lotes_dif = 0
    if ABA_LOTES:
        i_ul = indice_de(cab_saida, "uuid_lote", ABA_PREFEITURA, obrigatoria=False)
        i_mt = indice_de(cab_saida, COLUNA_MATRICULA, ABA_PREFEITURA,
                         obrigatoria=False)
        i_pr = indice_de(cab_saida, COLUNAS_HERDADAS_DO_CADASTRO[0], ABA_PREFEITURA,
                         obrigatoria=False) if COLUNAS_HERDADAS_DO_CADASTRO else None
        if i_ul is None or i_mt is None:
            print("    [aviso] aba '%s' pulada: falta uuid_lote ou %s"
                  % (ABA_LOTES, COLUNA_MATRICULA))
        else:
            lotes = {}
            for valores, _p, _m, _i, _st, _v, _o, _c, _t in saida:
                ul = norm_chave(valores[i_ul])
                if not ul:
                    continue
                cad = norm_chave(valores[k_pref])
                if not cad:
                    continue          # linha sem numero de cadastro nao identifica
                d = lotes.setdefault(ul, {})
                reg = d.setdefault(cad, {"mat": set(), "prop": "", "unid": []})
                if not vazio(valores[i_mt]):
                    reg["mat"].add(norm_chave(valores[i_mt]))
                if i_pr is not None and not reg["prop"] and not vazio(valores[i_pr]):
                    reg["prop"] = valores[i_pr]
                if CHAVE_COM_UNIDADE and not vazio(valores[u_pref]):
                    reg["unid"].append(norm_chave(valores[u_pref]))

            print("Gravando a aba '%s'..." % ABA_LOTES)
            if ABA_LOTES in wb.sheetnames:
                del wb[ABA_LOTES]
            wl = wb.create_sheet(ABA_LOTES)
            neg = Font(bold=True)
            f_dif2 = PatternFill("solid", fgColor=COR_BAIXAR)
            cabl = ["uuid_lote", "situacao", COLUNA_CHAVE_PREFEITURA,
                    COLUNA_MATRICULA, "unidades desse cadastro",
                    COLUNAS_HERDADAS_DO_CADASTRO[0] if COLUNAS_HERDADAS_DO_CADASTRO
                    else "", "cadastros no lote", "matriculas no lote"]
            for j, t in enumerate(cabl, start=1):
                c = wl.cell(row=1, column=j, value=t)
                c.font = neg
                c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
            r = 1
            for ul, cads in sorted(lotes.items()):
                if len(cads) < 2:
                    continue
                mats = sorted({m for d in cads.values() for m in d["mat"]})
                com = sum(1 for d in cads.values() if d["mat"])
                if len(mats) > 1:
                    sit = "MATRICULAS DIFERENTES"
                    lotes_dif += 1
                elif len(mats) == 1 and com < len(cads):
                    sit = "so um cadastro tem matricula"
                elif len(mats) == 1:
                    sit = "mesma matricula nos dois"
                else:
                    sit = "nenhum tem matricula"
                if LOTES_SO_MATRICULA_DIFERENTE and not sit.startswith("MATRICULAS"):
                    continue
                for cad, d in sorted(cads.items()):
                    r += 1
                    vals = [ul, sit, casa_tipo(cad, k_pref),
                            ", ".join(sorted(d["mat"])), ", ".join(d["unid"]),
                            d["prop"], ", ".join(sorted(cads)), ", ".join(mats)]
                    for j, v in enumerate(vals, start=1):
                        c = wl.cell(row=r, column=j, value=v)
                        if sit.startswith("MATRICULAS"):
                            c.fill = f_dif2
            wl.freeze_panes = "C2"
            wl.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cabl)), max(r, 1))
            for j, w in enumerate([40, 26, 16, 18, 20, 34, 26, 26], start=1):
                wl.column_dimensions[get_column_letter(j)].width = w

    if ESCREVER_EXTRAS_NO_DEPARA and _extras_novos:
        _wsd = achar_aba(wb, ABA_DEPARA)
        for _org, _dst in _extras_novos:
            _lin = _wsd.max_row + 1
            _wsd.cell(row=_lin, column=i_org + 1, value=_org)
            _wsd.cell(row=_lin, column=i_dst + 1, value=_dst)
            if i_atv is not None:
                _wsd.cell(row=_lin, column=i_atv + 1, value=DEPARA_VALOR_ATIVO[0])
        print("Acrescentei %d pares na aba '%s': %s"
              % (len(_extras_novos), ABA_DEPARA,
                 ", ".join("%s -> %s" % x for x in _extras_novos)))

    if FAZER_BACKUP:
        shutil.copy2(caminho, caminho + ".backup")
        print("Backup: %s.backup" % os.path.basename(caminho))

    wb.save(caminho)

    # ------------------------------------------------------------- relatorio
    print("")
    print("-" * 60)
    print("%7d  linhas na aba '%s'" % (len(saida), ABA_SAIDA))
    if SO_COM_INTERACAO:
        print("%7d  linhas da prefeitura fora da aba (sem interacao)" % fora_da_saida)
    print("%7d  linhas atualizadas com dado de campo" % n_atualizadas)
    print("%7d  celulas alteradas (amarelo)" % n_celulas)
    print("%7d  celulas conferidas em campo, mesmo valor (cinza)" % n_iguais)
    print("%7d  cadastros distintos atualizados" % len(cads_atualizados))
    if ABA_PENDENCIAS:
        print("%7d  unidades SEM coleta em cadastro que teve outra unidade atualizada"
              % pendentes)
    if ABA_LOTES:
        print("%7d  lotes com cadastros de matricula DIFERENTE" % lotes_dif)
    if REGRA_CONDOMINIO:
        print("%7d  condominios (edificacao com cadastros de matriculas diferentes)"
              % len(cond_info))
        for _cid, _inf in cond_info.items():
            print("            %s: %s" % (_cid, ", ".join(sorted(_inf["cads"]))))
        print("%7d    ...linhas de condominio criadas (verde-agua)" % n_linhas_cond)
        print("%7d    ...linhas de condominio com proprietario predominante dos"
              " membros" % n_prop_cond)
        print("%7d    ...linhas da prefeitura marcadas como membro" % n_membros_cond)
    if UM_CADASTRO_POR_LOTE:
        print("%7d  lotes com 2+ cadastros -> so um recebe o dado de campo"
              % len(escolhido_do_lote))
        print("%7d    ...cadastros que ficaram sem atualizar (vermelho escuro)"
              % len(cads_nao_atualizar))
        print("%7d    ...linhas da prefeitura marcadas NAO ATUALIZAR" % n_nao_atualiza)
        print("%7d    ...linhas da coleta que viraram unidade nova do escolhido"
              % n_transferidas)
    if REGRA_LOTE_BALDIO:
        print("%7d  baldio em cadastro que ja era terreno -> atualizou o existente"
              % n_baldio_atualiza)
    if TERRENO_RECEBE_UNIDADE_DA_COLETA:
        print("%7d  terreno (unidade vazia) que ganhou edificacao -> recebeu a"
              " unidade da coleta" % n_terreno_unidade)
    if MULTI_CADASTRO_ESCOLHE:
        print("%7d  celulas da coleta com varios numeros de cadastro -> ficou"
              " um so" % len(multi_usados))
        print("%7d    ...cadastros %s por causa disso (%d linhas)"
              % (len(cads_baixar_multi),
                 "marcados BAIXAR" if norm_cab(MULTI_CADASTRO_MODO) == "baixar"
                 else "que viraram COMPLEMENTARES",
                 n_baixar_multi))
    if soma_pref_multi:
        print("%7d  grupos de cadastros que dividem a edificacao -> a area"
              " construida e comparada somada (%d cadastros)"
              % (n_soma_grupo, len(soma_pref_multi)))
        for _t in sorted(multi_usados):
            _e = escolha_multi[_t]
            print("            %-40s -> %s (%s: %s)"
                  % (_t, _e,
                     "baixar" if norm_cab(MULTI_CADASTRO_MODO) == "baixar"
                     else "complementares",
                     ", ".join(x for x in multi_origem[_t]
                               if x != _e and x in cads_pref) or "-"))
    if doc_resp:
        print("%7d  responsaveis com o mesmo nome do proprietario -> documento do"
              " proprietario" % n_doc_resp[0])
        print("%7d  responsaveis trocados SEM documento na coleta e com nome"
              " diferente do proprietario (ficou o documento antigo)" % n_doc_resp[1])
    if cancelado_para or cancelados_sem_destino or avisos_cancel:
        print("%7d  cadastros cancelados por unificacao (aba '%s') -> %d cadastros"
              " receberam matricula e area"
              % (len(cancelado_para), ABA_CANCELADOS, len(cads_unificados_ok)))
        for _d in sorted(set(unificados) - cads_unificados_ok):
            print("            destino %s nao esta ativo na prefeitura: nada foi"
                  " somado (cancelados: %s)" % (_d, ", ".join(unificados[_d])))
        print("%7d  cadastros cancelados SEM unificacao (aba '%s') -> so baixados:"
              " %s" % (len(cancelados_sem_destino), ABA_CANCELADOS,
                       ", ".join(sorted(cancelados_sem_destino)) or "-"))
        for _c in sorted(cancelados_sem_destino - cads_pref):
            print("            %s ja esta desativado na prefeitura: nao aparece"
                  " na aba" % _c)
        for _a in avisos_cancel:
            print("            aviso: %s" % _a)
    if cads_baixados:
        print("%7d  linhas de cadastro BAIXADO pela lista manual ou pela aba '%s'"
              " (%d cadastros)"
              % (n_baixados_lista, ABA_CANCELADOS, len(cads_baixados)))
        _sem = sorted(cads_baixados - {norm_chave(_e[0][k_pref]) for _e in saida}
                      - (cancelados_sem_destino - cads_pref))
        if _sem:
            print("            nao achei na prefeitura: %s" % ", ".join(_sem))
        for _c, _u, _v in coleta_em_baixado:
            print("            coleta ignorada no %s/%s  (%s)" % (_c, _u, _v))
    print("")
    print("vinculo das %d linhas da bruta:" % len(dados_bru))
    print("%7d  pelo %s" % (n_via_edif, COLUNA_CADASTRO_EDIFICACAO))
    print("%7d  pelo %s + unidade" % (n_via_lote, COLUNA_CHAVE_BRUTA))
    print("%7d  sem vinculo -> linha nova (verde)" % novos)
    print("%7d    ...unidade nova de cadastro que ja existe (mantem o numero)"
          % unid_novas)
    print("%7d    ...cadastro que nao existe na prefeitura (mantem o numero)"
          % cad_inexistente)
    print("%7d    ...sem numero na coleta -> '%s'"
          % (novos_sem_numero, TEXTO_CADASTRO_NOVO))
    if herdadas:
        print("%7d  celulas herdadas do proprio cadastro nas linhas novas (%s)"
              % (herdados, ", ".join(COLUNAS_HERDADAS_DO_CADASTRO)))
    print("%7d  com %s invalido -> CONFERIR (vermelho)"
          % (len(linhas_erro), COLUNA_CADASTRO_EDIFICACAO))
    if REGRA_LOTE_BALDIO:
        print("%7d  lote baldio -> linha nova sem edificacao (lilas)" % baldio_novos)
        print("%7d  unidades a BAIXAR (vermelho): cadastro que virou baldio"
              % n_baixar)
        if baldio_repetidos:
            print("%7d    ...linhas da coleta repetidas no mesmo cadastro baldio"
                  % baldio_repetidos)
    if REGRA_EDIFICACAO:
        print("")
        print("regra da edificacao (uuid_edificacao):")
        print("%7d  edificacoes com mais de um registro (cadastro+unidade)"
              % len(grupos_multi))
        print("%7d    ...com matricula -> area da prefeitura mantida"
              % (len(grupos_multi) - len(grupos_principal)))
        print("%7d    ...sem matricula -> gerou cadastro principal" % len(grupos_principal))
        if PRINCIPAL_EDIF_MULTICADASTRO:
            print("%7d  linhas PRINCIPAL de edificacao dividida por varios"
                  " cadastros" % principais_multi)
        print("%7d  linhas PRINCIPAL criadas (azul)" % principais)
        print("%7d  celulas de area mantidas pela regra (laranja)" % n_travadas)
        if COMPLEMENTAR_MOSTRA_CODIGO:
            print("%7d  complementares com o codigo da propria linha em %s"
                  % (n_codigos, COLUNA_ORIGEM))
        if grupos_mistos:
            print("%7d  grupos MISTOS (um cadastro com matricula e outro sem), tratados"
                  " como '%s': %s" % (len(grupos_mistos), GRUPO_MISTO,
                                      ", ".join(grupos_mistos[:3])))
    if linhas_erro:
        print("\nvalores a corrigir em %s:" % COLUNA_CADASTRO_EDIFICACAO)
        vistos_e = []
        for _row, valor, erros in linhas_erro:
            if valor not in vistos_e:
                vistos_e.append(valor)
                print("   %-52s -> %s" % (valor, "; ".join(erros)))
    print("")
    if duplicados:
        print("%7d  chaves repetidas na bruta (usei a 1a): %s"
              % (len(duplicados), ", ".join(duplicados[:8])
                 + ("..." if len(duplicados) > 8 else "")))
    if contagem_por_col:
        print("\nalteracoes por coluna:")
        for nome, q in sorted(contagem_por_col.items(), key=lambda x: -x[1]):
            print("   %5d  %s" % (q, nome))
    print("-" * 60)
    print("Pronto. Abra o arquivo e veja a aba '%s'." % ABA_SAIDA)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (0, None):
            print("\nERRO: %s" % e.code)
    except Exception:
        import traceback
        print("\nERRO inesperado:\n")
        traceback.print_exc()
    input("\nEnter para fechar...")
