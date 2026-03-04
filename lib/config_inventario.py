# -*- coding: utf-8 -*-
"""
config_inventario.py
====================
Arquivo de configuração centralizado para o Inventário por Pavimento.

Todas as constantes aqui podem ser editadas sem tocar na biblioteca
(extracao_lib.py) ou no script do botão.

Sections
--------
  FILTROS DE FASE          -- defaults de fase_criacao / excluir_demolidos
  QUALIDADE / RELATÓRIO    -- limites do relatório de conformidade
  PARÂMETROS SAP           -- nomes aceitos para o campo Codigo SAP
  PARÂMETROS AMBIENTE      -- nomes aceitos para o campo Room/Space
  CATEGORIAS IGNORADAS     -- categorias excluídas do inventário
  FEATURES                 -- flags de ativação / desativação
"""

# ============================================================
# FILTROS DE FASE
# ============================================================

# Nome exato da fase de criação aceita.
# None → desativado (qualquer fase de criação é aceita)
DEFAULT_FASE_CRIACAO = None          # ex: u"Construção nova"

# True  → elementos com phase_demolished != "" são rejeitados
# False → elementos demolidos são incluídos
DEFAULT_EXCLUIR_DEMOLIDOS = False

# ============================================================
# QUALIDADE / RELATÓRIO DE CONFORMIDADE
# ============================================================

# Número máximo de amostras de problemas registradas no JSON
MAX_AMOSTRAS_PROBLEMAS = 200

# ============================================================
# PARÂMETROS SAP
# ============================================================
# Lista de nomes de parâmetro tentados (em ordem) para obter
# o Código SAP de uma instância ou tipo.  Adicione variantes
# conforme necessário para seu projeto.

SAP_PARAMETER_NAMES = [
    u"Codigo SAP",
    u"Código SAP",
    u"SAP Code",
    u"Cod SAP",
    u"Codigo_SAP",
    u"Código_SAP",
    u"SAP_Code",
    u"CodSAP",
]

# ============================================================
# PARÂMETROS AMBIENTE (Room / Space fallback)
# ============================================================
# Nomes tentados (em ordem) quando el.Room e el.Space
# não estão disponíveis.

ROOM_PARAMETER_NAMES = [
    u"Compartimento",
    u"Room Name",
    u"Nome do Compartimento",
    u"Nome Compartimento",
    u"Room",
    u"Ambiente",
    u"Space Name",
]

# ============================================================
# CATEGORIAS IGNORADAS
# ============================================================
# Elementos cujas categorias estejam nesta lista são
# excluídos do inventário (não contabilizados nem listados).
# Use os nomes exatos como aparecem no Revit (sem acentos,
# ou com acentos — a comparação é case-insensitive).

IGNORED_CATEGORIES = [
    u"Linhas",
    u"Cotas",
    u"Textos",
    u"Detalhes",
    u"Grades",
    u"Eixos",
    u"Regioes",
    u"Regiões",
    u"Simbolos de referencia",
    u"Símbolos de referência",
    u"Elementos ocultos",
]

# ============================================================
# FEATURES
# ============================================================

# True  → adiciona seção "flat_items" no JSON (lista plana de
#          todos os itens — útil para importação em BI/notebooks)
EXPORT_FLAT_ITEMS = True

# True  → adiciona colunas phase_created / phase_demolished no CSV
INCLUDE_PHASE_COLUMNS_IN_CSV = True

# True  → normaliza textos: strip() + colapsa espaços internos
NORMALIZE_TEXT = True
