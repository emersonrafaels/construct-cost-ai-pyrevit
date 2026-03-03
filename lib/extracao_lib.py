# -*- coding: utf-8 -*-
"""
extracao_lib.py
===============
Biblioteca compartilhada para os scripts de extração do VerificadorInteligente.

Funções principais:
  extrair_completo(doc, model, model_safe, run_stamp, day, ts)
      → Extrai todos os elementos: elements.csv, params.csv, model_hierarchy.json

  extrair_metadados(doc, model, model_safe, run_stamp, day, ts)
      → Extrai folhas e vistas: sheets_views.csv

  extrair_inventario_por_folha(doc, model, model_safe, run_stamp, day, now)
      → Inventário de equipamentos por folha (Sheet):
        inventario_por_folha.csv, inventario_totais.csv, inventario.json

Compatível com Revit 2022-2026  |  IronPython 2 / CPython 3
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
    LocationPoint,
    LocationCurve,
    StorageType,
    ViewSheet,
)

from collections import defaultdict, Counter
import os, csv, json, sys

# ============================================================
# CONFIG GLOBAL
# ============================================================
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "RevitScan"
MAX_PARAM_LEN    = 500

# Quando True, emite parâmetros do TIPO de cada elemento no JSON completo.
INCLUDE_TYPE_PARAMS_IN_JSON = True

# Categorias monitoradas pelo inventário por folha.
CATEGORIAS_INCLUIR = set([
    u"Luminarias",
    u"Dispositivos eletricos",
    u"Equipamentos eletricos",
    u"Circuitos eletricos",
    u"Identificadores de luminaria",
    u"Conduites",            u"Conexoes do conduite",
    u"Bandejas de cabos",    u"Conexoes da bandeja de cabos",
    u"Eletrocalhas",         u"Conexoes da eletrocalha",
    u"Fiacao",
    u"Sprinklers",           u"Tubulacoes",
    u"Aparelhos sanitarios",
    u"Acessorios de tubulacao",
    u"Equipamentos mecanicos",
    u"Mobiliario",           u"Especialidades",
    u"Componentes",          u"Detalhes de modelo",
    u"Portas",               u"Janelas",
])

# ============================================================
# UTILITÁRIOS COMUNS — Python 2 / Python 3
# ============================================================
try:
    _unicode = unicode  # IronPython 2
except NameError:
    _unicode = str      # CPython 3


def u8(x):
    """Converte qualquer valor para string compatível com csv.writer."""
    try:
        if x is None:
            return ""
        if sys.version_info[0] >= 3:
            return str(x)
        if isinstance(x, _unicode):
            return x.encode("utf-8")
        return str(x)
    except:
        try:
            return _unicode(x).encode("utf-8") if sys.version_info[0] < 3 else str(x)
        except:
            return ""


def safe_str(x):
    """Retorna a string ou '' se None — evita 'None' literal nos CSVs."""
    return x if x else ""


def write_bom_csv(path):
    """
    Abre arquivo CSV com BOM UTF-8 para que o Excel reconheça o encoding.
    Compatível com IronPython 2 (modo binário) e CPython 3 (encoding param).
    """
    if sys.version_info[0] >= 3:
        f = open(path, "w", encoding="utf-8-sig", newline="")
    else:
        f = open(path, "wb")
        f.write(b"\xef\xbb\xbf")
    return f


def make_dirs(*paths):
    """Cria todas as pastas fornecidas (equivalente a mkdir -p)."""
    for p in paths:
        if not os.path.exists(p):
            os.makedirs(p)


def eid_int(eid):
    """Extrai o valor inteiro de um ElementId (compatível com Revit < e >= 2024)."""
    try:
        return int(eid.IntegerValue)
    except:
        try:
            return int(eid.Value)
        except:
            try:
                return int(eid.ToString())
            except:
                return None


def bbox_to_dict(bb):
    """
    Converte BoundingBox em dict {'min': 'X,Y,Z', 'max': 'X,Y,Z'}.
    Retorna strings vazias se o elemento não tiver bounding box.
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
    Extrai localização geométrica do elemento.
    - LocationPoint → kind='POINT', value='X,Y,Z'
    - LocationCurve → kind='CURVE', value='X0,Y0,Z0 -> X1,Y1,Z1'
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
    """Retorna o nome do nível (pavimento) ao qual o elemento pertence."""
    try:
        if el.LevelId and el.LevelId != ElementId.InvalidElementId:
            lv = el.Document.GetElement(el.LevelId)
            return lv.Name if lv else ""
    except:
        pass
    return ""


def get_family_type(el):
    """
    Retorna (FamilyName, TypeName) do elemento.
    FamilyName = geometria/comportamento (ex: "Luminária Embutida Quadrada")
    TypeName   = variante específica       (ex: "60x60cm - 2x40W")
    """
    fam = u""
    typ = u""
    try:
        t = el.Document.GetElement(el.GetTypeId())
        if t:
            try:
                fam = t.FamilyName or u""
            except:
                pass
            try:
                typ = t.Name or u""
            except:
                pass
    except:
        pass
    return fam, typ


def get_workset_name(el):
    """Retorna o nome do Workset do elemento ('' em projetos sem worksets)."""
    try:
        ws_table = el.Document.GetWorksetTable()
        ws_id    = el.WorksetId
        ws       = ws_table.GetWorkset(ws_id)
        return ws.Name if ws else ""
    except:
        return ""


def get_phase_names(el):
    """Retorna (fase_criação, fase_demolição) do elemento."""
    try:
        pc = el.CreatedPhaseId
        pd = el.DemolishedPhaseId
        phase_c = el.Document.GetElement(pc).Name if pc != ElementId.InvalidElementId else ""
        phase_d = el.Document.GetElement(pd).Name if pd != ElementId.InvalidElementId else ""
        return (phase_c, phase_d)
    except:
        return ("", "")


def get_design_option(el):
    """Retorna o nome da Design Option do elemento, se houver."""
    try:
        do_id = el.DesignOption
        if do_id is None:
            return ""
        return do_id.Name
    except:
        return ""


def param_to_str(p, doc=None):
    """
    Converte o valor de um parâmetro Revit para string.
    Tipos: String, Integer, Double, ElementId.
    Para ElementId tenta resolver o nome do elemento referenciado.
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
            if doc:
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


def collect_params(el, max_len=MAX_PARAM_LEN, doc=None):
    """
    Coleta todos os parâmetros de um elemento (instância ou tipo).
    Retorna lista de dicts: {name, storage_type, value, is_shared, guid, group}.
    Parâmetros com valor vazio são descartados.
    """
    result = []
    for p in el.Parameters:
        try:
            name = p.Definition.Name
            st   = str(p.StorageType)
            val  = param_to_str(p, doc)

            if not val:
                continue
            if len(val) > max_len:
                val = val[:max_len]

            is_shared = False
            guid      = ""
            try:
                if p.IsShared:
                    is_shared = True
                    guid = str(p.GUID)
            except:
                pass

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


def write_json_file(path, data):
    """Serializa dict como JSON com BOM UTF-8. Compatível com Py2/Py3."""
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    if sys.version_info[0] >= 3:
        with open(path, "w", encoding="utf-8-sig") as fj:
            fj.write(json_str)
    else:
        fj = open(path, "wb")
        fj.write(b"\xef\xbb\xbf")
        if isinstance(json_str, _unicode):
            fj.write(json_str.encode("utf-8"))
        else:
            fj.write(json_str)
        fj.close()


# ============================================================
# FUNÇÃO 1 — EXTRAIR COMPLETO
# ============================================================

def extrair_completo(doc, model, model_safe, run_stamp, day, ts):
    """
    Extrai todos os elementos físicos do modelo Revit.

    Arquivos gerados:
      elements.csv          — um elemento por linha com metadados
      params.csv            — todos os parâmetros (instância + tipo) por elemento
      model_hierarchy.json  — hierarquia Categoria > Família > Tipo > Instâncias

    Parâmetros:
      doc        — documento Revit aberto
      model      — nome do modelo (doc.Title)
      model_safe — nome sanitizado para uso em caminhos de pasta
      run_stamp  — carimbo data+hora "%Y-%m-%d_%H%M%S"
      day        — data "%Y-%m-%d"
      ts         — hora "%H%M%S"

    Retorna dict:
      ok, skipped, n_elements, run_dir, latest_dir, hierarchy
    """
    ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
    RUN_DIR    = os.path.join(ROOT_DIR, model_safe, "ExtrairElementosCompleto", run_stamp)
    LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", "ExtrairElementosCompleto")
    make_dirs(RUN_DIR, LATEST_DIR)

    elements_path        = os.path.join(RUN_DIR,    "elements.csv")
    params_path          = os.path.join(RUN_DIR,    "params.csv")
    json_path            = os.path.join(RUN_DIR,    "model_hierarchy.json")
    elements_path_latest = os.path.join(LATEST_DIR, "elements.csv")
    params_path_latest   = os.path.join(LATEST_DIR, "params.csv")
    json_path_latest     = os.path.join(LATEST_DIR, "model_hierarchy.json")

    elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

    fe     = write_bom_csv(elements_path)
    fp     = write_bom_csv(params_path)
    fe_lat = write_bom_csv(elements_path_latest)
    fp_lat = write_bom_csv(params_path_latest)

    ew     = csv.writer(fe)
    pw     = csv.writer(fp)
    ew_lat = csv.writer(fe_lat)
    pw_lat = csv.writer(fp_lat)

    ELEM_HEADER = [
        "model", "element_id", "unique_id", "category", "family", "type",
        "level", "workset", "phase_created", "phase_demolished",
        "design_option", "loc_kind", "loc_value", "bbox_min", "bbox_max", "run_stamp",
    ]
    PARAM_HEADER = [
        "model", "unique_id", "element_id", "scope",
        "param_name", "storage_type", "value_str", "is_shared", "guid", "group",
    ]
    for w in (ew, ew_lat):
        w.writerow(ELEM_HEADER)
    for w in (pw, pw_lat):
        w.writerow(PARAM_HEADER)

    hierarchy         = {}
    type_params_cache = {}
    ok      = 0
    skipped = 0

    def _write_params(owner_uid, owner_id_str, scope, params_list):
        for p in params_list:
            try:
                row = [
                    u8(model), u8(owner_uid), u8(owner_id_str), u8(scope),
                    u8(p["name"]), u8(p["storage_type"]), u8(p["value"]),
                    u8("1" if p["is_shared"] else "0"), u8(p["guid"]), u8(p["group"]),
                ]
                pw.writerow(row)
                pw_lat.writerow(row)
            except:
                continue

    for el in elements:
        try:
            cat        = safe_str(el.Category.Name) if el.Category else ""
            fam, typ   = get_family_type(el)
            lvl        = get_level_name(el)
            workset    = get_workset_name(el)
            phase_c, phase_d = get_phase_names(el)
            design_opt = get_design_option(el)
            loc        = location_to_dict(el)
            bb         = bbox_to_dict(el.get_BoundingBox(None))

            el_id     = eid_int(el.Id)
            el_id_str = str(el_id) if el_id is not None else ""

            inst_params = collect_params(el, MAX_PARAM_LEN, doc)

            type_params = []
            type_id_str = ""
            try:
                t = doc.GetElement(el.GetTypeId())
                if t:
                    type_id_str = str(eid_int(t.Id)) if eid_int(t.Id) is not None else ""
                    if type_id_str not in type_params_cache:
                        type_params_cache[type_id_str] = collect_params(t, MAX_PARAM_LEN, doc)
                    type_params = type_params_cache[type_id_str]
            except:
                pass

            row = [
                u8(model), u8(el_id_str), u8(el.UniqueId),
                u8(cat), u8(fam), u8(typ), u8(lvl), u8(workset),
                u8(phase_c), u8(phase_d), u8(design_opt),
                u8(loc["kind"]), u8(loc["value"]),
                u8(bb["min"]), u8(bb["max"]), u8(run_stamp),
            ]
            ew.writerow(row)
            ew_lat.writerow(row)

            _write_params(el.UniqueId, el_id_str, "instance", inst_params)
            _write_params(el.UniqueId, el_id_str, "type",     type_params)

            # Hierarquia JSON
            cat_node = hierarchy.setdefault(cat, {"total": 0, "families": {}})
            cat_node["total"] += 1

            fam_node = cat_node["families"].setdefault(fam, {"total": 0, "types": {}})
            fam_node["total"] += 1

            if typ not in fam_node["types"]:
                type_params_for_json = {}
                if INCLUDE_TYPE_PARAMS_IN_JSON:
                    for tp in type_params:
                        type_params_for_json[tp["name"]] = tp["value"]
                fam_node["types"][typ] = {
                    "total":       0,
                    "type_params": type_params_for_json,
                    "instances":   [],
                }

            typ_node = fam_node["types"][typ]
            typ_node["total"] += 1
            inst_params_json = {p["name"]: p["value"] for p in inst_params}
            typ_node["instances"].append({
                "element_id":       el_id_str,
                "unique_id":        el.UniqueId,
                "level":            lvl,
                "workset":          workset,
                "phase_created":    phase_c,
                "phase_demolished": phase_d,
                "design_option":    design_opt,
                "location":         loc,
                "bbox":             bb,
                "instance_params":  inst_params_json,
            })

            ok += 1
        except:
            skipped += 1
            continue

    fe.close()
    fp.close()
    fe_lat.close()
    fp_lat.close()

    output_json = {
        "model":            model,
        "run_stamp":        run_stamp,
        "day":              day,
        "time":             ts,
        "run_dir":          RUN_DIR,
        "latest_dir":       LATEST_DIR,
        "total_elements":   ok,
        "total_skipped":    skipped,
        "total_categories": len(hierarchy),
        "hierarchy":        hierarchy,
    }
    write_json_file(json_path, output_json)
    write_json_file(json_path_latest, output_json)

    return {
        "ok":         ok,
        "skipped":    skipped,
        "n_elements": len(elements),
        "run_dir":    RUN_DIR,
        "latest_dir": LATEST_DIR,
        "hierarchy":  hierarchy,
    }


# ============================================================
# FUNÇÃO 2 — EXTRAIR METADADOS (Folhas e Vistas)
# ============================================================

def extrair_metadados(doc, model, model_safe, run_stamp, day, ts):
    """
    Extrai metadados de folhas (ViewSheet) e vistas do modelo.
    Cada folha pode conter múltiplas vistas via Viewports.

    Arquivos gerados:
      sheets_views.csv — um par (folha, vista) por linha

    Parâmetros: doc, model, model_safe, run_stamp, day, ts

    Retorna dict:
      total_pairs, skipped, n_sheets, run_dir, latest_dir
    """
    ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
    RUN_DIR    = os.path.join(ROOT_DIR, model_safe, "ExtrairMetadados", run_stamp)
    LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", "ExtrairMetadados")
    make_dirs(RUN_DIR, LATEST_DIR)

    out        = os.path.join(RUN_DIR,    "sheets_views.csv")
    out_latest = os.path.join(LATEST_DIR, "sheets_views.csv")

    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

    f     = write_bom_csv(out)
    f_lat = write_bom_csv(out_latest)
    w     = csv.writer(f)
    w_lat = csv.writer(f_lat)

    HEADER = [
        "model",           # Nome do arquivo Revit
        "run_stamp",       # Carimbo de data/hora desta execução
        "sheet_number",    # Número da folha (ex: "EL-01")
        "sheet_name",      # Título da folha
        "sheet_id",        # ID interno da folha
        "view_id",         # ID interno da vista
        "view_name",       # Nome da vista
        "view_type",       # FloorPlan, Section, Detail, ThreeD, etc.
        "view_discipline", # Architectural, Mechanical, Electrical, etc.
        "view_scale",      # Escala (ex: 50 = 1:50)
        "is_template",     # 1 = template | 0 = vista normal
    ]
    w.writerow(HEADER)
    w_lat.writerow(HEADER)

    total_pairs = 0
    skipped     = 0

    for sh in sheets:
        try:
            vports = sh.GetAllViewports()
            for vpid in vports:
                try:
                    vp = doc.GetElement(vpid)
                    v  = doc.GetElement(vp.ViewId)

                    discipline = ""
                    try:
                        discipline = str(v.Discipline)
                    except:
                        pass

                    scale = ""
                    try:
                        scale = str(v.Scale)
                    except:
                        pass

                    is_template = ""
                    try:
                        is_template = "1" if v.IsTemplate else "0"
                    except:
                        pass

                    row = [
                        u8(model),
                        u8(run_stamp),
                        u8(safe_str(sh.SheetNumber)),
                        u8(safe_str(sh.Name)),
                        u8(sh.Id.ToString()),
                        u8(v.Id.ToString()),
                        u8(safe_str(v.Name)),
                        u8(str(v.ViewType)),
                        u8(discipline),
                        u8(scale),
                        u8(is_template),
                    ]
                    w.writerow(row)
                    w_lat.writerow(row)
                    total_pairs += 1
                except:
                    skipped += 1
                    continue
        except:
            skipped += 1
            continue

    f.close()
    f_lat.close()

    return {
        "total_pairs": total_pairs,
        "skipped":     skipped,
        "n_sheets":    len(sheets),
        "run_dir":     RUN_DIR,
        "latest_dir":  LATEST_DIR,
    }


# ============================================================
# FUNÇÃO 3 — INVENTÁRIO POR FOLHA
# ============================================================

def extrair_inventario_por_folha(doc, model, model_safe, run_stamp, day, now):
    """
    Gera inventário de equipamentos agrupado por Folha (Sheet).
    Cada elemento é contado apenas uma vez por folha, mesmo aparecendo
    em múltiplas vistas da mesma folha.

    Arquivos gerados:
      inventario_por_folha.csv — (folha x elemento x quantidade)
      inventario_totais.csv    — totais globais por tipo
      inventario.json          — dados estruturados completos

    Parâmetros: doc, model, model_safe, run_stamp, day, now (datetime)

    Retorna dict:
      sheets_data, inventario, totais, folhas_com_itens, folhas_sem_itens,
      total_seen, run_dir, latest_dir
    """
    ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
    RUN_DIR    = os.path.join(ROOT_DIR, model_safe, "InventarioPorFolha", run_stamp)
    LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", "InventarioPorFolha")
    make_dirs(RUN_DIR, LATEST_DIR)

    out_folha     = os.path.join(RUN_DIR,    "inventario_por_folha.csv")
    out_totais    = os.path.join(RUN_DIR,    "inventario_totais.csv")
    out_json      = os.path.join(RUN_DIR,    "inventario.json")
    out_folha_lat = os.path.join(LATEST_DIR, "inventario_por_folha.csv")
    out_tot_lat   = os.path.join(LATEST_DIR, "inventario_totais.csv")
    out_json_lat  = os.path.join(LATEST_DIR, "inventario.json")

    # ----------------------------------------------------------
    # 1) Coleta de folhas (ViewSheets)
    # ----------------------------------------------------------
    sheets_data = {}
    for vs in (FilteredElementCollector(doc)
               .OfClass(ViewSheet)
               .WhereElementIsNotElementType()):
        try:
            sid  = eid_int(vs.Id)
            num  = (vs.SheetNumber or u"").strip()
            nome = (vs.Name or u"Sem Nome").strip()
            sheets_data[sid] = {"id": sid, "numero": num, "nome": nome}
        except:
            continue

    # ----------------------------------------------------------
    # 2) Mapeamento folha → elementos via Viewports
    #    Elemento registrado apenas UMA vez por folha.
    # ----------------------------------------------------------
    sheet_el_info = defaultdict(dict)   # sheet_id → {el_id: (cat, fam, typ)}
    total_seen    = 0

    for vs in (FilteredElementCollector(doc)
               .OfClass(ViewSheet)
               .WhereElementIsNotElementType()):
        try:
            sid = eid_int(vs.Id)
            if sid not in sheets_data:
                continue
            vp_ids = vs.GetAllViewports()
            if not vp_ids:
                continue

            for vp_id in vp_ids:
                try:
                    vp = doc.GetElement(vp_id)
                    if not vp:
                        continue
                    view_id = vp.ViewId
                    view    = doc.GetElement(view_id)
                    if not view:
                        continue

                    for el in (FilteredElementCollector(doc, view_id)
                               .WhereElementIsNotElementType()):
                        try:
                            cat_obj = el.Category
                            if not cat_obj:
                                continue
                            cat = cat_obj.Name or u""
                            if cat not in CATEGORIAS_INCLUIR:
                                continue
                            el_id = eid_int(el.Id)
                            if el_id is None:
                                continue
                            if el_id not in sheet_el_info[sid]:
                                fam, typ = get_family_type(el)
                                sheet_el_info[sid][el_id] = (cat, fam, typ)
                                total_seen += 1
                        except:
                            continue
                except:
                    continue
        except:
            continue

    # ----------------------------------------------------------
    # 3) Agrega: (folha, fam, typ) → quantidade
    # ----------------------------------------------------------
    inventario = defaultdict(lambda: defaultdict(int))
    for sid, el_dict in sheet_el_info.items():
        for el_id, (cat, fam, typ) in el_dict.items():
            inventario[sid][(fam, typ)] += 1

    folhas_com_itens = len(inventario)
    folhas_sem_itens = len(sheets_data) - folhas_com_itens

    totais = Counter()
    for sid, itens in inventario.items():
        for chave, cnt in itens.items():
            totais[chave] += cnt

    # ----------------------------------------------------------
    # 4) CSV — inventário por folha
    # ----------------------------------------------------------
    def write_folha_csv(path):
        f = write_bom_csv(path)
        wr = csv.writer(f)
        wr.writerow(["model", "run_stamp", "folha", "folha_nome",
                     "elemento", "nome", "qtd"])
        for sid in sorted(inventario.keys(),
                          key=lambda x: sheets_data.get(x, {}).get("numero", "")):
            sh    = sheets_data.get(sid, {"numero": "", "nome": ""})
            itens = inventario[sid]
            for (fam, typ), cnt in sorted(itens.items(),
                                          key=lambda x: (x[0][0], x[0][1])):
                wr.writerow([
                    u8(model), u8(run_stamp),
                    u8(sh["numero"]), u8(sh["nome"]),
                    u8(fam), u8(typ), u8(cnt),
                ])
        f.close()

    write_folha_csv(out_folha)
    write_folha_csv(out_folha_lat)

    # ----------------------------------------------------------
    # 5) CSV — totais globais
    # ----------------------------------------------------------
    def write_totais_csv(path):
        f = write_bom_csv(path)
        wr = csv.writer(f)
        wr.writerow(["model", "run_stamp", "elemento", "nome", "qtd_total"])
        for (fam, typ), cnt in sorted(totais.items(), key=lambda x: -x[1]):
            wr.writerow([u8(model), u8(run_stamp), u8(fam), u8(typ), u8(cnt)])
        f.close()

    write_totais_csv(out_totais)
    write_totais_csv(out_tot_lat)

    # ----------------------------------------------------------
    # 6) JSON
    # ----------------------------------------------------------
    inv_serial = {}
    for sid, itens in inventario.items():
        sh  = sheets_data.get(sid, {"numero": "", "nome": ""})
        key = u"{} -- {}".format(sh["numero"], sh["nome"])
        inv_serial[key] = {
            "id":          sid,
            "folha":       sh["numero"],
            "folha_nome":  sh["nome"],
            "total_itens": sum(itens.values()),
            "itens": [
                {"elemento": fam, "nome": typ, "qtd": cnt}
                for (fam, typ), cnt in sorted(itens.items())
            ],
        }

    resultado = {
        "model":     model,
        "run_stamp": run_stamp,
        "resumo": {
            "total_folhas":     len(sheets_data),
            "folhas_com_itens": folhas_com_itens,
            "folhas_sem_itens": folhas_sem_itens,
            "total_elementos":  total_seen,
            "tipos_distintos":  len(totais),
        },
        "inventario_por_folha": inv_serial,
        "totais_por_tipo": [
            {"elemento": fam, "nome": typ, "qtd_total": cnt}
            for (fam, typ), cnt in sorted(totais.items(), key=lambda x: -x[1])
        ],
    }

    write_json_file(out_json, resultado)
    write_json_file(out_json_lat, resultado)

    return {
        "sheets_data":      sheets_data,
        "inventario":       dict(inventario),
        "totais":           dict(totais),
        "folhas_com_itens": folhas_com_itens,
        "folhas_sem_itens": folhas_sem_itens,
        "total_seen":       total_seen,
        "run_dir":          RUN_DIR,
        "latest_dir":       LATEST_DIR,
    }
