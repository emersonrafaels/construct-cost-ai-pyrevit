# -*- coding: utf-8 -*-

# ============================================================
# CONTEXTO RÁPIDO PARA QUEM NÃO CONHECE REVIT
# ============================================================
# Este script faz uma extração COMPLETA de todos os elementos
# físicos de um modelo Revit aberto, gerando três arquivos:
#
#   1. elements.csv  — uma linha por elemento com seus metadados
#                      principais (categoria, família, tipo, nível,
#                      localização, bounding box, etc.)
#
#   2. params.csv    — todos os parâmetros (properties) de cada
#                      elemento, tanto de instância quanto de tipo.
#                      Um parâmetro Revit é equivalente a um atributo
#                      do objeto (ex: "Potência", "Tensão", "Material").
#
#   3. model_hierarchy.json — estrutura hierárquica completa do modelo:
#                      Categoria → Família → Tipo → Instâncias
#                      Inclui todos os parâmetros embutidos no JSON.
#
# Os arquivos são salvos em duas pastas:
#   - Pasta com carimbo de data/hora (histórico permanente)
#   - Pasta "latest" (sobrescrita a cada execução, fácil de automações)
# ============================================================

# pyrevit: biblioteca que faz a ponte entre Python e o Revit aberto.
from pyrevit import revit

# API oficial do Revit — acesso a elementos, tipos, famílias, parâmetros, etc.
from Autodesk.Revit.DB import *

import csv, os, datetime  # Utilitários padrão do Python
import json               # Para salvar a hierarquia em formato JSON estruturado

# "doc" = documento (modelo) atualmente aberto no Revit.
doc = revit.doc

# =========================================
# PARAMS (AJUSTE AQUI)
# =========================================

# Pasta raiz onde os arquivos serão salvos (padrão: Área de Trabalho do Windows).
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"

# Limite de caracteres para valores de parâmetros muito longos.
# Evita que textos enormes (ex: descrições completas) inflem o CSV/JSON.
MAX_PARAM_LEN = 500

# Quando True, inclui os parâmetros do TIPO de cada elemento no JSON.
# Útil para análises detalhadas, mas aumenta o tamanho do arquivo.
INCLUDE_TYPE_PARAMS_IN_JSON = True
# =========================================

# Captura o momento exato da execução para organizar as pastas de saída.
now       = datetime.datetime.now()
ts        = now.strftime("%H%M%S")           # Hora em formato HHMMSS
day       = now.strftime("%Y-%m-%d")         # Data em formato YYYY-MM-DD
run_stamp = now.strftime("%Y-%m-%d_%H%M%S") # Carimbo completo data+hora

# doc.Title retorna o nome do arquivo .rvt sem extensão.
model = doc.Title

# Remove caracteres inválidos para nomes de pasta no Windows.
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)

# Nome do plugin — usado como subpasta dentro de cada execução e do "latest",
# permitindo que vários scripts salvem arquivos no mesmo modelo sem conflito.
PLUGIN_NAME = "ExtrairElementosCompleto"

# ---------------------------------------------------------------
# Pasta com carimbo de data/hora — preserva histórico permanente.
# Estrutura gerada:
#   revit_dump/NomeModelo/2026-02-26_143021/ExtrairElementosCompleto/
# ---------------------------------------------------------------
RUN_DIR = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)  # Cria a pasta e todos os diretórios intermediários

# ---------------------------------------------------------------
# Pasta "latest" — sempre sobrescrita pela última execução.
# Útil para dashboards ou automações que precisam ler o resultado
# mais recente sem precisar descobrir qual pasta é a mais nova.
# Estrutura gerada:
#   revit_dump/NomeModelo/latest/ExtrairElementosCompleto/
# ---------------------------------------------------------------
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
if not os.path.exists(LATEST_DIR):
    os.makedirs(LATEST_DIR)

# Caminhos na pasta com carimbo (histórico permanente — nunca sobrescritos).
elements_path  = os.path.join(RUN_DIR, "elements.csv")
params_path    = os.path.join(RUN_DIR, "params.csv")
json_path      = os.path.join(RUN_DIR, "model_hierarchy.json")

# Caminhos na pasta "latest" (sempre sobrescritos).
elements_path_latest = os.path.join(LATEST_DIR, "elements.csv")
params_path_latest   = os.path.join(LATEST_DIR, "params.csv")
json_path_latest     = os.path.join(LATEST_DIR, "model_hierarchy.json")



# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def u8(x):
    """
    Converte qualquer valor para bytes UTF-8 de forma segura.
    Necessário pelo IronPython 2 ter dois tipos de string (str/unicode).
    O módulo csv do Python 2 exige bytes ao gravar em modo binário.
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


def safe_str(x):
    """Retorna a string ou '' se None — evita 'None' literal nos CSVs."""
    return x if x else ""


def write_bom_csv(path):
    """
    Abre arquivo CSV em modo binário e injeta BOM UTF-8 no início.
    O BOM faz o Excel reconhecer o encoding automaticamente,
    exibindo acentos sem precisar importar manualmente.
    """
    f = open(path, "wb")               # "wb" = escrita binária
    f.write(u"\ufeff".encode("utf-8")) # \ufeff = Byte Order Mark UTF-8
    return f


def eid_int(eid):
    """
    Extrai o valor inteiro de um ElementId do Revit.
    A API mudou entre versões do Revit (IntegerValue → Value),
    por isso tentamos os dois métodos antes de converter via ToString().
    """
    try:
        return eid.IntegerValue  # Revit < 2024
    except:
        pass
    try:
        return eid.Value         # Revit >= 2024
    except:
        pass
    try:
        return int(eid.ToString())
    except:
        return None


def bbox_to_dict(bb):
    """
    Converte a BoundingBox de um elemento em duas strings "X,Y,Z".
    BoundingBox é o cubo imaginário que envolve o elemento no espaço 3D.
    Min = vértice inferior-esquerdo-frontal; Max = superior-direito-traseiro.
    Retorna dict com 'min' e 'max' (ou strings vazias se não houver bbox).
    """
    if not bb:
        return {"min": "", "max": ""}
    mn, mx = bb.Min, bb.Max
    return {
        "min": "{:.4f},{:.4f},{:.4f}".format(mn.X, mn.Y, mn.Z),
        "max": "{:.4f},{:.4f},{:.4f}".format(mx.X, mx.Y, mx.Z),
    }


def location_to_dict(el):
    """
    Extrai a localização geométrica de um elemento.

    No Revit existem dois tipos de localização:
      - LocationPoint: elementos com posição fixa (ex: uma luminária, uma porta).
        Retorna (kind='POINT', value='X,Y,Z').
      - LocationCurve: elementos lineares (ex: um duto, um conduíte, uma parede).
        Retorna (kind='CURVE', value='X0,Y0,Z0 -> X1,Y1,Z1') com os dois extremos.

    Retorna dict com 'kind' e 'value'.
    """
    loc = el.Location
    if not loc:
        return {"kind": "", "value": ""}
    if isinstance(loc, LocationPoint):
        p = loc.Point
        return {
            "kind":  "POINT",
            "value": "{:.4f},{:.4f},{:.4f}".format(p.X, p.Y, p.Z),
        }
    if isinstance(loc, LocationCurve):
        c  = loc.Curve
        p0 = c.GetEndPoint(0)
        p1 = c.GetEndPoint(1)
        return {
            "kind":  "CURVE",
            "value": "{:.4f},{:.4f},{:.4f} -> {:.4f},{:.4f},{:.4f}".format(
                p0.X, p0.Y, p0.Z, p1.X, p1.Y, p1.Z),
        }
    return {"kind": "", "value": ""}


def get_level_name(el):
    """
    Retorna o nome do nível (pavimento/andar) ao qual o elemento pertence.
    No Revit, elementos são associados a "Levels" (ex: "Térreo", "1º Pav.").
    ElementId.InvalidElementId indica que não há nível associado.
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

    No Revit:
      - Family (Família): define a geometria e o comportamento de um componente.
        Ex: "Luminária Embutida Quadrada".
      - Type (Tipo): variante específica dentro da família com parâmetros fixos.
        Ex: "60x60cm - 2x40W".
    Cada instância no modelo é criada a partir de um tipo de uma família.
    """
    try:
        t = doc.GetElement(el.GetTypeId())  # Obtém o tipo pelo ID
        if not t:
            return ("", "")
        fam = ""
        try:
            fam = t.FamilyName  # Nome da família pai
        except:
            fam = ""
        return (fam, t.Name)    # t.Name = nome do tipo específico
    except:
        return ("", "")


def get_workset_name(el):
    """
    Retorna o nome do Workset ao qual o elemento pertence.
    Worksets são conjuntos de trabalho usados em projetos colaborativos
    onde múltiplos usuários editam o mesmo modelo Revit simultaneamente.
    Retorna '' em modelos de usuário único (sem worksets).
    """
    try:
        ws_table = doc.GetWorksetTable()
        ws_id    = el.WorksetId
        ws       = ws_table.GetWorkset(ws_id)
        return ws.Name if ws else ""
    except:
        return ""


def get_phase_names(el):
    """
    Retorna (fase_criação, fase_demolição) do elemento.
    Fases no Revit representam etapas construtivas do projeto
    (ex: "Fase 1 - Estrutura", "Fase 2 - Acabamento").
    Elementos demolidos em uma fase são removidos do modelo naquela etapa.
    """
    try:
        pc = el.CreatedPhaseId
        pd = el.DemolishedPhaseId
        phase_c = doc.GetElement(pc).Name if pc != ElementId.InvalidElementId else ""
        phase_d = doc.GetElement(pd).Name if pd != ElementId.InvalidElementId else ""
        return (phase_c, phase_d)
    except:
        return ("", "")


def get_design_option(el):
    """
    Retorna o nome da Design Option do elemento, se houver.
    Design Options permitem modelar variantes do projeto (ex: duas versões
    de layout) dentro do mesmo arquivo Revit para comparação.
    """
    try:
        do_id = el.DesignOption
        if do_id is None:
            return ""
        return do_id.Name
    except:
        return ""


def param_to_str(p):
    """
    Converte o valor de um parâmetro Revit para string.

    Parâmetros podem ter 4 tipos de armazenamento (StorageType):
      - String:    texto livre
      - Integer:   número inteiro (também usado para booleans: 0/1)
      - Double:    número decimal em unidades internas do Revit (pés)
      - ElementId: referência a outro elemento do modelo (ex: material, tipo)

    Para ElementId, tenta resolver o nome do elemento referenciado.
    """
    st = p.StorageType
    try:
        if st == StorageType.String:
            return p.AsString() or ""
        if st == StorageType.Integer:
            return str(p.AsInteger())
        if st == StorageType.Double:
            # Valor bruto em unidades internas (pés). Para conversão, use UnitUtils.
            return "{:.6f}".format(p.AsDouble())
        if st == StorageType.ElementId:
            eid = p.AsElementId()
            if eid == ElementId.InvalidElementId:
                return ""
            ref = doc.GetElement(eid)
            if ref:
                try:
                    return "{}:{}".format(eid_int(eid), ref.Name)
                except:
                    return str(eid_int(eid))
            return str(eid_int(eid))
    except:
        return ""
    return ""


def collect_params(el, max_len):
    """
    Coleta todos os parâmetros de um elemento e retorna uma lista de dicts.
    Cada dict contém: name, storage_type, value, is_shared, guid, group.

    Parâmetros com valor vazio são ignorados para reduzir o volume de dados.
    'group' indica a aba onde o parâmetro aparece nas Properties do Revit
    (ex: "Identity Data", "Dimensions", "Electrical").
    """
    result = []
    for p in el.Parameters:
        try:
            name = p.Definition.Name
            st   = str(p.StorageType)
            val  = param_to_str(p)

            if not val:
                continue  # Ignora parâmetros sem valor

            if len(val) > max_len:
                val = val[:max_len]  # Trunca valores muito longos

            is_shared = False
            guid      = ""
            try:
                if p.IsShared:        # Parâmetros compartilhados têm GUID único
                    is_shared = True
                    guid = str(p.GUID)
            except:
                pass

            # BuiltInParameterGroup indica a aba/grupo do parâmetro na UI do Revit
            group = ""
            try:
                group = str(p.Definition.ParameterGroup)
            except:
                try:
                    group = str(p.Definition.GetGroupTypeId())
                except:
                    pass

            result.append({
                "name":         name,
                "storage_type": st,
                "value":        val,
                "is_shared":    is_shared,
                "guid":         guid,
                "group":        group,
            })
        except:
            continue
    return result


def write_params_to_csv(owner_unique_id, owner_id, scope, params_list, writer):
    """
    Grava a lista de parâmetros no CSV de params.
    'scope' indica se os parâmetros são de instância ("instance") ou de tipo ("type").
    """
    for p in params_list:
        try:
            writer.writerow([
                u8(model),
                u8(owner_unique_id),
                u8(owner_id),
                u8(scope),
                u8(p["name"]),
                u8(p["storage_type"]),
                u8(p["value"]),
                u8("1" if p["is_shared"] else "0"),
                u8(p["guid"]),
                u8(p["group"]),
            ])
        except:
            continue


# ============================================================
# COLETA DE TODOS OS ELEMENTOS DO MODELO
# ============================================================
# FilteredElementCollector varre o modelo procurando elementos.
# .WhereElementIsNotElementType() traz apenas INSTÂNCIAS físicas
# (objetos colocados no modelo), excluindo templates e definições de tipo.
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

print("Iniciando extracao de {} elementos...".format(len(elements)))

# ============================================================
# ABERTURA DOS ARQUIVOS CSV (histórico + latest)
# ============================================================
# Abrimos todos os 4 arquivos de uma vez para escrever em paralelo
# no mesmo loop, evitando percorrer a lista de elementos duas vezes.
fe       = write_bom_csv(elements_path)
fp       = write_bom_csv(params_path)
fe_lat   = write_bom_csv(elements_path_latest)
fp_lat   = write_bom_csv(params_path_latest)

ew     = csv.writer(fe)
pw     = csv.writer(fp)
ew_lat = csv.writer(fe_lat)
pw_lat = csv.writer(fp_lat)

# Cabeçalhos do CSV de elementos — uma coluna por metadado do elemento.
ELEM_HEADER = [
    "model",          # Nome do arquivo Revit
    "element_id",     # ID numérico único dentro do modelo (muda entre versões)
    "unique_id",      # GUID único global (estável entre versões do modelo)
    "category",       # Categoria Revit (ex: "Walls", "Ducts", "Lighting Fixtures")
    "family",         # Nome da família (ex: "Luminária Embutida Quadrada")
    "type",           # Nome do tipo dentro da família (ex: "60x60cm - 2x40W")
    "level",          # Nível/pavimento ao qual o elemento pertence
    "workset",        # Workset para projetos colaborativos (ou '' se não houver)
    "phase_created",  # Fase em que o elemento foi criado/construído
    "phase_demolished",# Fase em que o elemento foi demolido ('' = ainda existe)
    "design_option",  # Design Option (variante do projeto) ou ''
    "loc_kind",       # Tipo de localização: "POINT", "CURVE" ou ''
    "loc_value",      # Coordenadas da localização
    "bbox_min",       # Vértice mínimo da bounding box (X,Y,Z)
    "bbox_max",       # Vértice máximo da bounding box (X,Y,Z)
    "run_stamp",      # Carimbo de data/hora desta execução
]
ew.writerow(ELEM_HEADER)
ew_lat.writerow(ELEM_HEADER)

# Cabeçalhos do CSV de parâmetros — um parâmetro por linha.
PARAM_HEADER = [
    "model",        # Nome do modelo
    "unique_id",    # GUID do elemento dono deste parâmetro
    "element_id",   # ID numérico do elemento dono
    "scope",        # "instance" = parâmetro da instância | "type" = do tipo
    "param_name",   # Nome do parâmetro (ex: "Tensão nominal", "Potência")
    "storage_type", # Tipo de dado: String, Integer, Double ou ElementId
    "value_str",    # Valor convertido para string
    "is_shared",    # 1 = parâmetro compartilhado (tem GUID) | 0 = embutido
    "guid",         # GUID do parâmetro compartilhado (ou '' se não for)
    "group",        # Grupo/aba onde aparece na janela de Properties do Revit
]
pw.writerow(PARAM_HEADER)
pw_lat.writerow(PARAM_HEADER)


# ============================================================
# ESTRUTURA HIERÁRQUICA PARA O JSON
# ============================================================
# Construímos um dicionário aninhado com a hierarquia do Revit:
#   Categoria → Família → Tipo → Lista de instâncias
#
# Cada instância no JSON já contém seus parâmetros embutidos,
# eliminando a necessidade de juntar tabelas para análise.
#
# Estrutura:
# {
#   "CategoryName": {
#     "total": N,
#     "families": {
#       "FamilyName": {
#         "total": N,
#         "types": {
#           "TypeName": {
#             "total": N,
#             "type_params": {...},   ← parâmetros do tipo (compartilhados por instâncias)
#             "instances": [...]      ← lista de instâncias com seus parâmetros
#           }
#         }
#       }
#     }
#   }
# }
hierarchy = {}   # Dicionário principal da hierarquia
type_params_cache = {}  # Cache para não re-coletar parâmetros do mesmo tipo Id

ok      = 0  # Contador de elementos extraídos com sucesso
skipped = 0  # Contador de elementos que causaram erro e foram pulados


# ============================================================
# LOOP PRINCIPAL — PROCESSA CADA ELEMENTO
# ============================================================
for el in elements:
    try:
        # --- Metadados básicos do elemento ---
        cat        = safe_str(el.Category.Name) if el.Category else ""
        fam, typ   = get_family_type(el)
        lvl        = get_level_name(el)
        workset    = get_workset_name(el)
        phase_c, phase_d = get_phase_names(el)
        design_opt = get_design_option(el)
        loc        = location_to_dict(el)
        bb         = bbox_to_dict(el.get_BoundingBox(None))  # None = View=None → bbox do modelo

        el_id     = eid_int(el.Id)
        el_id_str = str(el_id) if el_id is not None else ""

        # --- Coleta de parâmetros de instância ---
        inst_params = collect_params(el, MAX_PARAM_LEN)

        # --- Coleta de parâmetros do TIPO (com cache para eficiência) ---
        # O mesmo tipo pode ser compartilhado por centenas de instâncias.
        # O cache evita re-processar os mesmos parâmetros múltiplas vezes.
        type_params = []
        type_id_str = ""
        try:
            t = doc.GetElement(el.GetTypeId())
            if t:
                type_id_str = str(eid_int(t.Id)) if eid_int(t.Id) is not None else ""
                if type_id_str not in type_params_cache:
                    type_params_cache[type_id_str] = collect_params(t, MAX_PARAM_LEN)
                type_params = type_params_cache[type_id_str]
        except:
            pass

        # -------------------------------------------------------
        # ESCRITA NO CSV DE ELEMENTOS (histórico + latest)
        # -------------------------------------------------------
        row = [
            u8(model),
            u8(el_id_str),
            u8(el.UniqueId),
            u8(cat),
            u8(fam),
            u8(typ),
            u8(lvl),
            u8(workset),
            u8(phase_c),
            u8(phase_d),
            u8(design_opt),
            u8(loc["kind"]),
            u8(loc["value"]),
            u8(bb["min"]),
            u8(bb["max"]),
            u8(run_stamp),
        ]
        ew.writerow(row)
        ew_lat.writerow(row)

        # -------------------------------------------------------
        # ESCRITA NO CSV DE PARÂMETROS (instância + tipo)
        # -------------------------------------------------------
        write_params_to_csv(el.UniqueId, el_id_str, "instance", inst_params, pw)
        write_params_to_csv(el.UniqueId, el_id_str, "instance", inst_params, pw_lat)
        write_params_to_csv(el.UniqueId, el_id_str, "type",     type_params, pw)
        write_params_to_csv(el.UniqueId, el_id_str, "type",     type_params, pw_lat)

        # -------------------------------------------------------
        # CONSTRUÇÃO DA HIERARQUIA JSON
        # -------------------------------------------------------
        # Garante que os níveis do dicionário existem antes de inserir.
        # setdefault() cria a chave com o valor padrão se ela não existir.
        cat_node = hierarchy.setdefault(cat, {"total": 0, "families": {}})
        cat_node["total"] += 1

        fam_node = cat_node["families"].setdefault(fam, {"total": 0, "types": {}})
        fam_node["total"] += 1

        # Parâmetros do tipo são adicionados apenas na primeira instância
        # desse tipo (evita repetição — todos os dados de tipo são iguais).
        if typ not in fam_node["types"]:
            type_params_for_json = {}
            if INCLUDE_TYPE_PARAMS_IN_JSON:
                for tp in type_params:
                    type_params_for_json[tp["name"]] = tp["value"]
            fam_node["types"][typ] = {
                "total":       0,
                "type_params": type_params_for_json,  # Parâmetros do tipo (compartilhados)
                "instances":   [],                     # Lista de instâncias
            }

        typ_node = fam_node["types"][typ]
        typ_node["total"] += 1

        # Constrói o dicionário da instância com todos os metadados.
        instance_params_for_json = {p["name"]: p["value"] for p in inst_params}

        typ_node["instances"].append({
            "element_id":       el_id_str,
            "unique_id":        el.UniqueId,
            "level":            lvl,
            "workset":          workset,
            "phase_created":    phase_c,
            "phase_demolished": phase_d,
            "design_option":    design_opt,
            "location":         loc,         # {"kind": "POINT"|"CURVE", "value": "X,Y,Z..."}
            "bbox":             bb,          # {"min": "X,Y,Z", "max": "X,Y,Z"}
            "instance_params":  instance_params_for_json,  # {nome: valor}
        })

        ok += 1

    except:
        skipped += 1
        continue

# ============================================================
# FECHAMENTO DOS ARQUIVOS CSV
# ============================================================
fe.close()
fp.close()
fe_lat.close()
fp_lat.close()


# ============================================================
# GERAÇÃO DO JSON HIERÁRQUICO
# ============================================================
# Envelope final: adiciona metadados globais ao redor da hierarquia.
output_json = {
    "model":            model,        # Nome do arquivo Revit
    "run_stamp":        run_stamp,    # Carimbo de data/hora desta execução
    "day":              day,          # Data (YYYY-MM-DD)
    "time":             ts,           # Hora (HHMMSS)
    "run_dir":          RUN_DIR,      # Pasta do histórico
    "latest_dir":       LATEST_DIR,   # Pasta "latest"
    "total_elements":   ok,           # Total de instâncias extraídas
    "total_skipped":    skipped,      # Total de elementos ignorados por erro
    "total_categories": len(hierarchy),
    # A hierarquia principal: Categoria → Família → Tipo → Instâncias
    "hierarchy":        hierarchy,
}


def write_json(path):
    """
    Salva o JSON em modo binário com BOM UTF-8.
    IronPython 2 não suporta o parâmetro 'encoding' no json.dump,
    por isso escrevemos manualmente como bytes.
    """
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))                          # BOM UTF-8
    f.write(u8(json.dumps(output_json, ensure_ascii=False, indent=2)))
    f.close()

# Salva em ambas as pastas
write_json(json_path)
write_json(json_path_latest)


# ============================================================
# OUTPUT — Resumo no console do pyRevit
# ============================================================
print("=" * 55)
print("  EXTRACAO DE ELEMENTOS COMPLETA")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(RUN_DIR))
print("    - elements.csv")
print("    - params.csv")
print("    - model_hierarchy.json")
print("")
print("  [Ultima execucao]")
print("    {}".format(LATEST_DIR))
print("    - elements.csv")
print("    - params.csv")
print("    - model_hierarchy.json")
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Elementos extraidos : {}".format(ok))
print("  Elementos pulados   : {}".format(skipped))
print("  Total no collector  : {}".format(len(elements)))
print("  Categorias unicas   : {}".format(len(hierarchy)))
print("")
print("-" * 55)
print("  HIERARQUIA (Top 15 categorias)")
print("-" * 55)
for cat_name, cat_data in sorted(hierarchy.items(), key=lambda x: -x[1]["total"])[:15]:
    n_fams  = len(cat_data["families"])
    n_types = sum(len(f["types"]) for f in cat_data["families"].values())
    print("  {:>5}  {}  ({} familias, {} tipos)".format(
        cat_data["total"], cat_name, n_fams, n_types))
print("=" * 55)