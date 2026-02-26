# -*- coding: utf-8 -*-

# ============================================================
# CONTEXTO RAPIDO PARA QUEM NAO CONHECE REVIT
# ============================================================
# Este script faz uma extracao RAPIDA e FILTRADA do modelo Revit.
# Diferente da extracao completa, ele foca apenas nas categorias
# relevantes para instalacoes (eletrica, infra, seguranca) e coleta
# somente os parametros mais importantes de cada elemento.
#
# Resultado: dois arquivos CSV menores e mais rapidos de processar,
# ideais para analises preliminares ou integracao com ferramentas
# que nao precisam de todos os dados do modelo.
#
# Arquivos gerados:
#   elements_rapido.csv  — metadados dos elementos filtrados
#   params_top.csv       — apenas os parametros-chave de cada elemento
#
# Os arquivos sao salvos em duas pastas:
#   - Pasta com carimbo de data/hora (historico permanente)
#   - Pasta "latest" (sobrescrita a cada execucao)
# ============================================================

# pyrevit: biblioteca que faz a ponte entre Python e o Revit aberto.
from pyrevit import revit

# API oficial do Revit — acesso a elementos, categorias, niveis, etc.
from Autodesk.Revit.DB import *

import csv, os, datetime  # Utilitarios padrao do Python

# "doc" = documento (modelo) atualmente aberto no Revit.
doc = revit.doc

# =========================
# PARAMS (AJUSTE AQUI)
# =========================

# Pasta raiz onde os arquivos serao salvos (padrao: Area de Trabalho do Windows).
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"

# Limite de caracteres para valores de parametros muito longos.
MAX_PARAM_LEN = 300

# Nome do plugin — usado como subpasta para organizar os arquivos.
# Estrutura: revit_dump/NomeModelo/ExtrairElementosRapido/timestamp/
PLUGIN_NAME = "ExtrairElementosRapido"

# ---------------------------------------------------------------
# Categorias do Revit que serao incluidas na extracao.
# Elementos de outras categorias sao ignorados para manter o CSV
# enxuto e focado nas disciplinas de instalacoes.
# ---------------------------------------------------------------
ALLOW_CATEGORIES = set([
    # Eletrica / MEP
    u"Conduites", u"Conexoes do conduite", u"Fiacao", u"Circuitos eletricos",
    u"Luminarias", u"Dispositivos eletricos", u"Equipamentos eletricos",
    u"Bandejas de cabos", u"Conexoes da bandeja de cabos",
    u"Eletrocalhas", u"Conexoes da eletrocalha",
    # Identificadores/Tags
    u"Identificadores de luminaria", u"Identificadores de dispositivos eletricos",
    # Espacos e referencias uteis
    u"Ambientes", u"Salas", u"Niveis", u"Linhas de centro",
])

# ---------------------------------------------------------------
# Parametros-chave coletados de cada elemento.
# Ao inves de coletar todos os parametros (como na extracao completa),
# aqui filtramos apenas os mais relevantes para analise de custo/projeto.
# ---------------------------------------------------------------
TOP_PARAMS = set([
    u"Mark", u"Marca", u"Comentarios", u"Comment",
    u"Nome do sistema", u"Classificacao do sistema", u"System Classification",
    u"Numero do circuito", u"Circuit Number",
    u"Painel", u"Panel",
    u"Tensao", u"Voltage",
    u"Carga", u"Load",
    u"Fabricante", u"Manufacturer",
    u"Modelo", u"Model",
])
# =========================

# Captura o momento exato da execucao para organizar as pastas de saida.
now       = datetime.datetime.now()
ts        = now.strftime("%H%M%S")           # Hora em formato HHMMSS
day       = now.strftime("%Y-%m-%d")         # Data em formato YYYY-MM-DD
run_stamp = now.strftime("%Y-%m-%d_%H%M%S") # Carimbo completo data+hora

# doc.Title retorna o nome do arquivo .rvt sem extensao.
model = doc.Title

# Remove caracteres invalidos para nomes de pasta no Windows.
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)

# ---------------------------------------------------------------
# Pasta com carimbo de data/hora — preserva historico permanente.
# Estrutura: revit_dump/NomeModelo/ExtrairElementosRapido/2026-02-26_143021/
# ---------------------------------------------------------------
RUN_DIR = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)  # Cria a pasta e todos os diretorios intermediarios

# ---------------------------------------------------------------
# Pasta "latest" — sempre sobrescrita pela ultima execucao.
# Estrutura: revit_dump/NomeModelo/latest/ExtrairElementosRapido/
# ---------------------------------------------------------------
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
if not os.path.exists(LATEST_DIR):
    os.makedirs(LATEST_DIR)

# Caminhos na pasta com carimbo (historico permanente).
elements_path = os.path.join(RUN_DIR, "elements_rapido.csv")
params_path   = os.path.join(RUN_DIR, "params_top.csv")

# Caminhos na pasta "latest" (sempre sobrescritos).
elements_path_latest = os.path.join(LATEST_DIR, "elements_rapido.csv")
params_path_latest   = os.path.join(LATEST_DIR, "params_top.csv")


# ============================================================
# FUNCOES AUXILIARES
# ============================================================

def u8(x):
    """
    Converte qualquer valor para bytes UTF-8 de forma segura.
    Necessario pelo IronPython 2 ter dois tipos de string (str/unicode).
    O modulo csv do Python 2 exige bytes ao gravar em modo binario.
    """
    try:
        if x is None:
            return ""
        if isinstance(x, unicode):
            return x.encode("utf-8")
        return str(x)
    except:
        try:
            return unicode(x).encode("utf-8")
        except:
            return ""


def write_bom_csv(path):
    """
    Abre arquivo CSV em modo binario e injeta BOM UTF-8 no inicio.
    O BOM faz o Excel reconhecer o encoding automaticamente,
    exibindo acentos sem precisar importar manualmente.
    """
    f = open(path, "wb")               # "wb" = escrita binaria
    f.write(u"\ufeff".encode("utf-8")) # \ufeff = Byte Order Mark UTF-8
    return f


def eid_int(eid):
    """
    Extrai o valor inteiro de um ElementId do Revit.
    Tenta IntegerValue (Revit < 2024) e Value (Revit >= 2024).
    """
    try: return eid.IntegerValue
    except: pass
    try: return eid.Value
    except: pass
    try: return int(eid.ToString())
    except: return None


def bbox_to_str(bb):
    """
    Converte a BoundingBox de um elemento em duas strings "X,Y,Z".
    BoundingBox e o cubo imaginario que envolve o elemento no espaco 3D.
    Retorna ('', '') se o elemento nao tiver bounding box.
    """
    if not bb:
        return ("", "")
    mn, mx = bb.Min, bb.Max
    return (
        "{:.4f},{:.4f},{:.4f}".format(mn.X, mn.Y, mn.Z),
        "{:.4f},{:.4f},{:.4f}".format(mx.X, mx.Y, mx.Z),
    )


def location_to_str(el):
    """
    Extrai a localizacao geometrica do elemento como string.
    - LocationPoint: posicao fixa (ex: luminaria) → 'X,Y,Z'
    - LocationCurve: elemento linear (ex: conduite) → 'X0,Y0,Z0 -> X1,Y1,Z1'
    """
    loc = el.Location
    if not loc:
        return ("", "")
    if isinstance(loc, LocationPoint):
        p = loc.Point
        return ("POINT", "{:.4f},{:.4f},{:.4f}".format(p.X, p.Y, p.Z))
    if isinstance(loc, LocationCurve):
        c  = loc.Curve
        p0 = c.GetEndPoint(0)
        p1 = c.GetEndPoint(1)
        return ("CURVE", "{:.4f},{:.4f},{:.4f} -> {:.4f},{:.4f},{:.4f}".format(
            p0.X, p0.Y, p0.Z, p1.X, p1.Y, p1.Z))
    return ("", "")


def get_level_name(el):
    """
    Retorna o nome do nivel (pavimento) ao qual o elemento pertence.
    Niveis no Revit representam os andares/pavimentos do projeto.
    """
    try:
        if el.LevelId and el.LevelId != ElementId.InvalidElementId:
            lv = doc.GetElement(el.LevelId)
            return lv.Name if lv else ""
    except:
        pass
    return ""


def get_family_type(el):
    """
    Retorna (FamilyName, TypeName) do elemento.
    Family = familia do componente (ex: 'Luminaria Embutida').
    Type   = variante especifica (ex: '60x60cm - 2x40W').
    """
    try:
        t = doc.GetElement(el.GetTypeId())
        if not t:
            return ("", "")
        fam = ""
        try: fam = t.FamilyName
        except: fam = ""
        return (fam, t.Name)
    except:
        return ("", "")


def get_workset_name(el):
    """
    Retorna o nome do Workset do elemento.
    Worksets sao conjuntos de trabalho em projetos colaborativos.
    Retorna '' em modelos de usuario unico.
    """
    try:
        ws_table = doc.GetWorksetTable()
        ws       = ws_table.GetWorkset(el.WorksetId)
        return ws.Name if ws else ""
    except:
        return ""


def param_to_str(p):
    """
    Converte o valor de um parametro Revit para string.
    Parametros podem ser: String, Integer, Double ou ElementId.
    Para ElementId, resolve o nome do elemento referenciado quando possivel.
    """
    st = p.StorageType
    try:
        if st == StorageType.String:
            return p.AsString() or ""
        if st == StorageType.Integer:
            return str(p.AsInteger())
        if st == StorageType.Double:
            return "{:.6f}".format(p.AsDouble())
        if st == StorageType.ElementId:
            eid = p.AsElementId()
            if eid == ElementId.InvalidElementId:
                return ""
            ref = doc.GetElement(eid)
            if ref:
                try: return "{}:{}".format(eid_int(eid), ref.Name)
                except: return str(eid_int(eid))
            return str(eid_int(eid))
    except:
        return ""
    return ""


def write_top_params(unique_id, element_id, el, writer):
    """
    Grava apenas os parametros definidos em TOP_PARAMS no CSV de params.
    Ignora parametros sem valor ou que nao estejam na lista de interesse.
    """
    for p in el.Parameters:
        try:
            name = p.Definition.Name
            if name not in TOP_PARAMS:
                continue       # Pula parametros fora da lista de interesse
            val = param_to_str(p)
            if not val:
                continue       # Pula parametros sem valor
            if len(val) > MAX_PARAM_LEN:
                val = val[:MAX_PARAM_LEN]
            writer.writerow([u8(model), u8(unique_id), u8(element_id),
                             u8(name), u8(val), u8(run_stamp)])
        except:
            continue


# ============================================================
# ABERTURA DOS ARQUIVOS CSV (historico + latest)
# ============================================================
fe     = write_bom_csv(elements_path)
fp     = write_bom_csv(params_path)
fe_lat = write_bom_csv(elements_path_latest)
fp_lat = write_bom_csv(params_path_latest)

ew     = csv.writer(fe)
pw     = csv.writer(fp)
ew_lat = csv.writer(fe_lat)
pw_lat = csv.writer(fp_lat)

# Cabecalho do CSV de elementos.
ELEM_HEADER = [
    "model",       # Nome do arquivo Revit
    "element_id",  # ID numerico do elemento
    "unique_id",   # GUID unico global do elemento
    "category",    # Categoria Revit (ex: "Luminarias", "Conduites")
    "family",      # Nome da familia
    "type",        # Nome do tipo dentro da familia
    "level",       # Nivel/pavimento
    "workset",     # Workset (ou '' se nao houver)
    "loc_kind",    # Tipo de localizacao: POINT, CURVE ou ''
    "loc_value",   # Coordenadas da localizacao
    "bbox_min",    # Vertice minimo da bounding box (X,Y,Z)
    "bbox_max",    # Vertice maximo da bounding box (X,Y,Z)
    "run_stamp",   # Carimbo de data/hora desta execucao
]
ew.writerow(ELEM_HEADER)
ew_lat.writerow(ELEM_HEADER)

# Cabecalho do CSV de parametros-chave.
PARAM_HEADER = [
    "model",       # Nome do modelo
    "unique_id",   # GUID do elemento dono deste parametro
    "element_id",  # ID numerico do elemento dono
    "param_name",  # Nome do parametro
    "value_str",   # Valor convertido para string
    "run_stamp",   # Carimbo de data/hora
]
pw.writerow(PARAM_HEADER)
pw_lat.writerow(PARAM_HEADER)


# ============================================================
# COLETA DE ELEMENTOS FILTRADOS
# ============================================================
# Varre todos os elementos do modelo, mas processa apenas os
# que pertencem às categorias definidas em ALLOW_CATEGORIES.
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

print("Iniciando extracao rapida de {} elementos no total...".format(len(elements)))

kept    = 0  # Elementos incluidos (dentro das categorias de interesse)
skipped = 0  # Elementos ignorados (categoria nao listada ou erro)

for el in elements:
    try:
        cat = el.Category.Name if el.Category else ""

        # Filtra por categoria — ignora tudo que nao e relevante.
        if cat and (cat not in ALLOW_CATEGORIES):
            continue

        el_id     = eid_int(el.Id)
        el_id_str = str(el_id) if el_id is not None else ""

        fam, typ  = get_family_type(el)
        lvl       = get_level_name(el)
        workset   = get_workset_name(el)
        loc_kind, loc_val = location_to_str(el)
        bbmin, bbmax = bbox_to_str(el.get_BoundingBox(None))

        # Escreve a linha de elemento em ambos os CSVs.
        row = [
            u8(model), u8(el_id_str), u8(el.UniqueId),
            u8(cat), u8(fam), u8(typ), u8(lvl), u8(workset),
            u8(loc_kind), u8(loc_val),
            u8(bbmin), u8(bbmax), u8(run_stamp),
        ]
        ew.writerow(row)
        ew_lat.writerow(row)

        # Escreve os parametros-chave em ambos os CSVs.
        write_top_params(el.UniqueId, el_id_str, el, pw)
        write_top_params(el.UniqueId, el_id_str, el, pw_lat)

        kept += 1
    except:
        skipped += 1
        continue

# Fecha todos os arquivos abertos.
fe.close()
fp.close()
fe_lat.close()
fp_lat.close()


# ============================================================
# OUTPUT — Resumo no console do pyRevit
# ============================================================
print("=" * 55)
print("  EXTRACAO RAPIDA DE ELEMENTOS CONCLUIDA")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(RUN_DIR))
print("    - elements_rapido.csv")
print("    - params_top.csv")
print("")
print("  [Ultima execucao]")
print("    {}".format(LATEST_DIR))
print("    - elements_rapido.csv")
print("    - params_top.csv")
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Total no modelo     : {}".format(len(elements)))
print("  Elementos extraidos : {}".format(kept))
print("  Elementos ignorados : {}".format(len(elements) - kept - skipped))
print("  Elementos com erro  : {}".format(skipped))
print("")
print("-" * 55)
print("  CATEGORIAS PERMITIDAS : {}".format(len(ALLOW_CATEGORIES)))
print("  PARAMETROS COLETADOS  : {}".format(len(TOP_PARAMS)))
print("=" * 55)
print("PROCESSO REALIZADO COM SUCESSO")