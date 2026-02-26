# -*- coding: utf-8 -*-

# ============================================================
# CONTEXTO RÁPIDO PARA QUEM NÃO CONHECE REVIT
# ============================================================
# Revit é um software BIM (Building Information Modeling) da Autodesk.
# Um "modelo" Revit é um arquivo .rvt com todos os dados de um projeto:
# paredes, portas, janelas, tubulações, cabos, etc.
# Cada item dentro do modelo é chamado de "elemento" (Element).
# Elementos são agrupados em "categorias" (ex: "Walls", "Doors", "Ducts").
# O modelo também pode conter "folhas" (ViewSheet), que são as pranchas
# de impressão do projeto (equivalente às páginas de um desenho técnico).
#
# Este script roda DENTRO do Revit via pyRevit — um add-in que permite
# executar Python dentro do Revit. Ele lê o modelo aberto e gera
# arquivos CSV/JSON com estatísticas e resumo do projeto.
# ============================================================

# pyrevit é a biblioteca que faz a ponte entre Python e o Revit aberto.
from pyrevit import revit

# Autodesk.Revit.DB é a API oficial do Revit. Contém todas as classes
# para acessar elementos, vistas, folhas, parâmetros, etc.
from Autodesk.Revit.DB import *

from collections import Counter   # Para contar ocorrências de categorias
import os, csv, datetime          # Utilitários padrão do Python
import json                       # Para salvar dados estruturados em JSON

# "doc" representa o documento (modelo) atualmente aberto no Revit.
# É o ponto de entrada para acessar qualquer dado do projeto.
doc = revit.doc

# =========================
# PARAMS (AJUSTE AQUI)
# =========================

# Pasta raiz onde os arquivos serão gerados (padrão: Área de Trabalho do Windows).
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop")

# Nome da pasta principal criada dentro de BASE_DIR.
ROOT_FOLDER_NAME = "revit_dump"

# Quantas categorias exibir no ranking das mais frequentes.
TOP_N = 25

# Palavras-chave para classificar as folhas do projeto por disciplina.
# Cada tupla é: ("NOME_DA_CLASSE", ["lista", "de", "palavras-chave"]).
# O script verifica se alguma palavra-chave aparece no número ou nome da folha.
# Exemplo: uma folha "EL-01 Quadro Elétrico" será classificada como "ELETRICA".
SHEET_KEYWORDS = [
    ("ELETRICA", ["EL", "ELE", "ELÉTR", "E-"]),
    ("INFRA",    ["INF", "TI", "DADOS", "REDE", "RACK", "CABEAMENTO"]),
    ("SEGURANCA",["SEG", "CFTV", "ALARME", "ACESSO", "CAMERA", "CÂMERA"]),
]
# =========================

# Captura o momento exato da execução — usado para criar a pasta do histórico.
now = datetime.datetime.now()
ts        = now.strftime("%H%M%S")           # Hora no formato HHMMSS (para o JSON)
day       = now.strftime("%Y-%m-%d")         # Data no formato YYYY-MM-DD
run_stamp = now.strftime("%Y-%m-%d_%H%M%S") # Carimbo completo para nomear a pasta

# doc.Title retorna o nome do arquivo do modelo Revit (sem extensão).
model = doc.Title

# Remove caracteres que são inválidos em nomes de pasta no Windows.
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)

# Nome do plugin — usado como subpasta dentro de cada execução e do "latest",
# permitindo que vários scripts salvem arquivos no mesmo modelo sem conflito.
PLUGIN_NAME = "ResumoModelo"

# ---------------------------------------------------------------
# Pasta com carimbo de data/hora — preserva o histórico de cada
# execução sem sobrescrever a anterior. Estrutura gerada:
#   revit_dump/NomeModelo/2026-02-26_143021/ResumoModelo/
# ---------------------------------------------------------------
RUN_DIR = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)  # Cria a pasta e todos os diretórios intermediários

# ---------------------------------------------------------------
# Pasta "latest" — sempre contém os arquivos da ÚLTIMA execução.
# Útil para dashboards ou automações que precisam ler o resultado
# mais recente sem precisar descobrir qual pasta é a mais nova.
#   revit_dump/NomeModelo/latest/ResumoModelo/
# ---------------------------------------------------------------
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
if not os.path.exists(LATEST_DIR):
    os.makedirs(LATEST_DIR)

# Caminhos dos 3 arquivos de saída na pasta com carimbo (histórico permanente).
out_topcats = os.path.join(RUN_DIR, "top_categorias.csv")
out_sheets  = os.path.join(RUN_DIR, "folhas_resumo.csv")
out_json    = os.path.join(RUN_DIR, "model_profile.json")

# Caminhos dos mesmos 3 arquivos na pasta "latest" (sempre sobrescritos).
out_topcats_latest = os.path.join(LATEST_DIR, "top_categorias.csv")
out_sheets_latest  = os.path.join(LATEST_DIR, "folhas_resumo.csv")
out_json_latest    = os.path.join(LATEST_DIR, "model_profile.json")


# ---------------------------------------------------------------
# Converte qualquer valor para bytes UTF-8 de forma segura.
#
# Necessário porque o pyRevit usa IronPython 2, que tem DOIS tipos
# de string: str (sequência de bytes) e unicode (texto real).
# O módulo csv do Python 2 espera bytes, então precisamos garantir
# a conversão antes de gravar qualquer valor.
# ---------------------------------------------------------------
def u8(x):
    try:
        if x is None:
            return ""
        if isinstance(x, unicode):          # Tipo exclusivo do Python 2
            return x.encode("utf-8")
        return str(x)
    except:
        try:
            return unicode(x).encode("utf-8")
        except:
            return ""


# ---------------------------------------------------------------
# Abre um arquivo CSV para escrita e injeta o BOM UTF-8 no início.
#
# BOM (Byte Order Mark) é uma sequência de 3 bytes (\xEF\xBB\xBF)
# que sinaliza ao Excel que o arquivo está em UTF-8. Sem ele, o
# Excel costuma exibir caracteres especiais (acentos, ç) errados.
# ---------------------------------------------------------------
def write_bom_csv(path):
    f = open(path, "wb")                   # "wb" = escrita em modo binário
    f.write(u"\ufeff".encode("utf-8"))     # \ufeff é o caractere BOM
    return f                               # Retorna o arquivo aberto para uso externo


# ---------------------------------------------------------------
# Normaliza uma string para maiúsculas de forma segura.
#
# Lida com os dois tipos de string do IronPython 2 (str e unicode)
# e com possíveis erros de codificação em nomes de folhas do Revit.
# Usado para comparar palavras-chave sem distinção de maiúsculas/minúsculas.
# ---------------------------------------------------------------
def norm_upper(s):
    try:
        if isinstance(s, unicode):
            return s.upper()
        return unicode(s, errors="ignore").upper()  # "ignore" descarta bytes inválidos
    except:
        try:
            return str(s).upper()
        except:
            return ""


# ============================================================
# 1) COLETA DE ELEMENTOS E CATEGORIAS
# ============================================================
# FilteredElementCollector é a classe da API do Revit para consultar
# elementos dentro do modelo (funciona como um SELECT no banco de dados).
#
# .WhereElementIsNotElementType() filtra apenas INSTÂNCIAS reais, ou seja,
# objetos que existem fisicamente no projeto (uma parede específica, uma
# porta colocada em planta, etc.). Exclui os "tipos" (templates/definições
# que descrevem como um elemento deve ser, mas não estão no modelo).
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

# Para cada elemento, tenta ler o nome da sua categoria.
# O try/except é necessário porque alguns elementos internos do Revit
# (como anotações ou objetos de sistema) não têm categoria definida
# e lançariam AttributeError ao acessar .Category.Name.
cats = []
for el in elements:
    try:
        if el.Category:                      # Verifica se a categoria não é None
            cats.append(el.Category.Name)   # Ex: "Walls", "Doors", "Ducts", "Pipes"
    except:
        pass  # Ignora elementos problemáticos silenciosamente

# Counter recebe a lista e retorna um dicionário {categoria: quantidade}.
# Exemplo: {"Walls": 320, "Doors": 85, "Windows": 60, ...}
counter = Counter(cats)

# ============================================================
# 2) COLETA E CLASSIFICAÇÃO DE FOLHAS (PRANCHAS)
# ============================================================
# ViewSheet é a classe que representa uma prancha de impressão no Revit.
# Cada folha tem:
#   - SheetNumber: código alfanumérico (ex: "EL-01", "ARQ-05")
#   - Name: título descritivo (ex: "Planta Baixa - Pavimento Térreo")
# Aqui usamos o FilteredElementCollector com .OfClass() para buscar
# apenas objetos do tipo ViewSheet dentro do modelo.
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()


def classify_sheet(sh):
    """
    Classifica uma folha de projeto por disciplina usando heurística
    de palavras-chave definidas em SHEET_KEYWORDS.

    Concatena SheetNumber + Name e verifica se alguma palavra-chave
    aparece no texto (em maiúsculas, para comparação sem case-sensitivity).
    Retorna a disciplina detectada ou "OUTROS" se nenhuma bater.
    """
    # Junta número e nome em uma string única e normaliza para maiúsculas.
    # O "or ''" evita erros caso SheetNumber ou Name sejam None no Revit.
    name = norm_upper((sh.SheetNumber or "") + " " + (sh.Name or ""))

    hits = []  # Lista das disciplinas detectadas nesta folha
    for label, keys in SHEET_KEYWORDS:
        for k in keys:
            if k in name:           # Verifica se a palavra-chave está no texto
                hits.append(label)  # Registra a disciplina
                break               # Já confirmou essa categoria, vai para a próxima

    if not hits:
        return "OUTROS"  # Nenhuma palavra-chave reconhecida

    # Se uma folha bater em mais de uma disciplina, os nomes são concatenados.
    # Exemplo: folha que contém "EL" e "CFTV" → "ELETRICA+SEGURANCA"
    hits = sorted(list(set(hits)))  # Remove duplicatas e ordena alfabeticamente
    return "+".join(hits)


# Conta quantas folhas foram classificadas em cada disciplina.
sheet_classes = Counter()
for sh in sheets:
    try:
        sheet_classes[classify_sheet(sh)] += 1
    except:
        pass

# ============================================================
# 3) SALVAR CSV — TOP CATEGORIAS
# ============================================================
# Grava um CSV com as TOP_N categorias mais frequentes no modelo.
# Colunas: nome do modelo | total de elementos | categoria | quantidade.
def write_topcats(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "total_elementos", "categoria", "qtd"])  # Linha de cabeçalho
    # .most_common(TOP_N) retorna lista de tuplas [(cat, qty), ...] ordenada do maior para menor.
    for cat, qtd in counter.most_common(TOP_N):
        w.writerow([u8(model), u8(len(elements)), u8(cat), u8(qtd)])
    f.close()

write_topcats(out_topcats)         # Salva no histórico (pasta com carimbo data/hora)
write_topcats(out_topcats_latest)  # Salva no "latest" (sobrescreve a execução anterior)

# ============================================================
# 4) SALVAR CSV — RESUMO DE FOLHAS
# ============================================================
# Grava um CSV com TODAS as folhas do projeto e a disciplina detectada.
# Colunas: nome do modelo | número da folha | nome da folha | classificação.
def write_sheets(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "sheet_number", "sheet_name", "classificacao"])  # Cabeçalho
    for sh in sheets:
        try:
            w.writerow([u8(model), u8(sh.SheetNumber), u8(sh.Name), u8(classify_sheet(sh))])
        except:
            continue  # Pula folhas que causarem erro de leitura (corrompidas ou especiais)
    f.close()

write_sheets(out_sheets)         # Salva no histórico
write_sheets(out_sheets_latest)  # Salva no "latest"

# ============================================================
# 5) SALVAR JSON — PERFIL COMPLETO DO MODELO
# ============================================================
# Consolida todas as métricas em um único JSON estruturado.
# Útil para integrar com dashboards, APIs ou outros scripts Python.
profile = {
    "model":             model,           # Nome do arquivo Revit
    "run_dir":           RUN_DIR,         # Caminho da pasta do histórico desta execução
    "latest_dir":        LATEST_DIR,      # Caminho da pasta "latest"
    "day":               day,            # Data da execução (YYYY-MM-DD)
    "time":              ts,             # Hora da execução (HHMMSS)
    "run_stamp":         run_stamp,       # Carimbo completo (data + hora)
    "total_elements":    len(elements),   # Total de instâncias físicas no modelo
    "unique_categories": len(counter),    # Número de categorias distintas encontradas
    "top_categories": [                   # Lista das TOP_N categorias mais frequentes
        {"category": c, "count": n}
        for c, n in counter.most_common(TOP_N)
    ],
    "total_sheets":  len(sheets),         # Total de pranchas (folhas) no projeto
    "sheet_classes": dict(sheet_classes), # Dicionário {disciplina: qtd_de_folhas}
}


def write_json(path):
    # IronPython 2 não suporta o parâmetro "encoding" em json.dump,
    # por isso escrevemos em modo binário ("wb") com BOM UTF-8 manualmente,
    # da mesma forma que fazemos nos arquivos CSV.
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    f.write(u8(json.dumps(profile, ensure_ascii=False, indent=2)))
    f.close()

write_json(out_json)         # Salva no histórico
write_json(out_json_latest)  # Salva no "latest"

# ============================================================
# OUTPUT — Exibe o resumo no console do pyRevit
# ============================================================
print("=" * 55)
print("  RESUMO DO MODELO CONCLUIDO")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(RUN_DIR))
print("    - top_categorias.csv")
print("    - folhas_resumo.csv")
print("    - model_profile.json")
print("")
print("  [Ultima execucao]")
print("    {}".format(LATEST_DIR))
print("    - top_categorias.csv")
print("    - folhas_resumo.csv")
print("    - model_profile.json")
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Elementos    : {}".format(len(elements)))
print("  Categorias   : {}".format(len(counter)))
print("  Folhas       : {}".format(len(sheets)))
print("")
print("-" * 55)
print("  TOP {} CATEGORIAS".format(TOP_N))
print("-" * 55)
for cat, qtd in counter.most_common(10):
    print("  {:>6}  {}".format(qtd, cat))
print("")
print("-" * 55)
print("  FOLHAS POR DISCIPLINA")
print("-" * 55)
for k, v in sheet_classes.most_common():
    print("  {:>6}  {}".format(v, k))
print("=" * 55)
print("PROCESSO REALIZADO COM SUCESSO")