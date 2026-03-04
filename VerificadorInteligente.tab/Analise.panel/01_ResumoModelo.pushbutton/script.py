# -*- coding: utf-8 -*-
# ============================================================
# RESUMO DO MODELO REVIT — VISAO GERAL COMPLETA
# ============================================================
# Gera um RESUMO EXECUTIVO do modelo Revit aberto, exibindo as
# principais metricas e rankings em um painel HTML interativo
# e exportando os dados para CSV e JSON.
#
# Secoes produzidas:
#   1.  Informacoes do Projeto
#   2.  Versao do Revit
#   3.  Estatisticas Gerais
#   4.  Top Categorias de Elementos
#   5.  Top Familias (por instancias)
#   6.  Top Tipos de Familia
#   7.  Top Folhas (Sheets por numero de vistas)
#   8.  Vistas por Tipo
#   9.  Niveis / Pavimentos
#  10.  Worksets
#  11.  Top Alertas (Warnings)
#  12.  Modelos Linkados (RVT Links)
#  13.  Fases do Projeto
#
# Arquivos gerados (dual-save: historico + latest):
#   resumo_modelo.json   — snapshot completo do modelo
#   top_categorias.csv   — ranking de categorias
#   top_familias.csv     — ranking de familias
#   top_folhas.csv       — ranking de folhas por numero de vistas
#
# Compativel com Revit 2022-2026  |  IronPython 2 / CPython 3
# ============================================================

from pyrevit import revit, script

# Imports explicitos — evita poluicao de namespace e erros silenciosos
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FilteredWorksetCollector,
    WorksetKind,
    Level,
    RevitLinkInstance,
    DesignOption,
    ViewSheet,
    Viewport,
    View,
    Family,
)

# UnitUtils disponivel apenas no Revit 2022+
try:
    from Autodesk.Revit.DB import UnitUtils, UnitTypeId
    _HAS_UNIT = True
except ImportError:
    _HAS_UNIT = False

from collections import Counter, OrderedDict, defaultdict
import os, csv, datetime, json, sys

# ============================================================
# CONFIGURACOES (AJUSTE AQUI)
# ============================================================
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "RevitScan"
PLUGIN_NAME      = "ResumoModelo"
TOP_N            = 15   # Quantidade de itens em cada ranking
BAR_WIDTH        = 18   # Largura da barra ASCII de progresso
# ============================================================

doc = revit.doc
app = __revit__.Application  # objeto principal do Revit

now        = datetime.datetime.now()
ts         = now.strftime("%H%M%S")
day        = now.strftime("%Y-%m-%d")
run_stamp  = now.strftime("%Y-%m-%d_%H%M%S")
model      = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR    = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)

for _d in [RUN_DIR, LATEST_DIR]:
    if not os.path.exists(_d):
        os.makedirs(_d)


# ============================================================
# UTILITARIOS — Python 2 / Python 3 / IronPython
# ============================================================
try:
    _unicode = unicode   # Python 2 / IronPython
except NameError:
    _unicode = str       # Python 3


def u8(x):
    """Garante string unicode pura; trata bytes UTF-8 e Latin-1."""
    if x is None:
        return u""
    if isinstance(x, _unicode):
        return x
    if isinstance(x, (bytes,)):
        for enc in ("utf-8", "latin-1"):
            try:
                return x.decode(enc)
            except (UnicodeDecodeError, AttributeError):
                continue
        return u""
    try:
        return _unicode(x)
    except Exception:
        return u""


def eid_int(eid):
    """
    Retorna o valor inteiro de um ElementId ou WorksetId.
    Revit 2022-2023: usa .IntegerValue
    Revit 2024+    : usa .Value  (Int64; IntegerValue foi removido)
    """
    try:
        return eid.Value          # Revit 2024+
    except AttributeError:
        return eid.IntegerValue   # Revit 2022-2023


def to_meters(internal_value):
    """Converte pes internos do Revit para metros."""
    try:
        if _HAS_UNIT:
            return round(
                UnitUtils.ConvertFromInternalUnits(internal_value, UnitTypeId.Meters), 2
            )
        return round(internal_value * 0.3048, 2)
    except Exception:
        return 0.0


def normalize_json(obj):
    """Garante que todo objeto seja serializavel em JSON pelo IronPython 2."""
    if isinstance(obj, dict):
        return {normalize_json(k): normalize_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [normalize_json(v) for v in obj]
    if isinstance(obj, _unicode):
        return obj
    if isinstance(obj, (str, bytes)):
        for enc in ("utf-8", "latin-1"):
            try:
                return obj.decode(enc)
            except (UnicodeDecodeError, AttributeError):
                continue
        return u""
    return obj


def write_bom_csv(path):
    """Abre CSV com BOM UTF-8 (compativel com Excel)."""
    if sys.version_info[0] >= 3:
        return open(path, "w", encoding="utf-8-sig", newline="")
    f = open(path, "wb")
    f.write(b"\xef\xbb\xbf")
    return f


def csv_row(*cols):
    """Gera linha CSV compativel com Python 2 (bytes) e Python 3 (str)."""
    if sys.version_info[0] >= 3:
        return [u8(c) for c in cols]
    return [u8(c).encode("utf-8") for c in cols]


def bar(value, maximum, width=BAR_WIDTH):
    """Barra de progresso proporcional para representacao visual de rankings."""
    if not maximum or maximum == 0:
        return u"[" + u" " * width + u"]"
    filled = int(round(value * width / float(maximum)))
    filled = min(filled, width)
    return u"[" + u"\u2588" * filled + u"\u2591" * (width - filled) + u"]"


def save_json(obj, *paths):
    """Serializa dicionario para JSON UTF-8 em todos os caminhos fornecidos."""
    for path in paths:
        try:
            safe = normalize_json(obj)
            blob = json.dumps(safe, indent=2, ensure_ascii=True)
            if isinstance(blob, _unicode):
                blob = blob.encode("utf-8")
            with open(path, "wb") as jf:
                jf.write(blob)
        except Exception as ex:
            with open(path, "wb") as jf:
                jf.write((u'{"error": "' + u8(str(ex)) + u'"}').encode("utf-8"))


# ============================================================
# INICIALIZA PAINEL HTML DO PYREVIT
# ============================================================
# script.get_output() retorna um objeto que renderiza HTML no painel
# lateral do pyRevit — suporta tabelas, Markdown e links de elementos.
output = script.get_output()
output.set_height(920)
output.set_title(u"Resumo do Modelo — " + u8(model))

output.print_md(u"# Resumo do Modelo Revit")
output.print_md(
    u"**Modelo:** `{}`  \n**Execucao:** `{}`".format(u8(model), run_stamp)
)


# ============================================================
# COLETA GLOBAL (base unica para todos os rankings)
# ============================================================
# WhereElementIsNotElementType exclui tipos/familias e retorna apenas instancias.
all_elements = list(
    FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements()
)


# ============================================================
# 1. INFORMACOES DO PROJETO
# ============================================================
output.print_md(u"---\n## 1. Informacoes do Projeto")

proj = doc.ProjectInformation
proj_table = [
    [u"Nome do Projeto", u8(proj.Name)],
    [u"Numero",          u8(proj.Number)],
    [u"Cliente",         u8(proj.ClientName)],
    [u"Endereco",        u8(proj.Address)],
    [u"Status",          u8(proj.Status)],
    [u"Autor",           u8(proj.Author)],
    [u"Data de Emissao", u8(proj.IssueDate)],
]
output.print_table(proj_table, columns=[u"Campo", u"Valor"])


# ============================================================
# 2. VERSAO DO REVIT
# ============================================================
output.print_md(u"---\n## 2. Versao do Revit")

output.print_table([
    [u"Versao", u8(str(app.VersionNumber))],
    [u"Nome",   u8(str(app.VersionName))],
    [u"Build",  u8(str(app.VersionBuild))],
    [u"Idioma", u8(str(app.Language))],
], columns=[u"Campo", u"Valor"])


# ============================================================
# PRE-COLETA DE RECURSOS DO MODELO
# (necessaria antes das Estatisticas Gerais)
# ============================================================

# Todas as vistas (Views) — inclui ViewSheet, ViewPlan, View3D, etc.
all_views = list(
    FilteredElementCollector(doc).OfClass(View).ToElements()
)

# Apenas folhas (Sheets)
sheets = list(
    FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
)

# Niveis / Pavimentos
levels = list(
    FilteredElementCollector(doc).OfClass(Level).ToElements()
)

# Modelos linkados
links = list(
    FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
)

# Opcoes de projeto (Design Options)
d_opts = list(
    FilteredElementCollector(doc).OfClass(DesignOption).ToElements()
)

# Familias carregadas no modelo
families = list(
    FilteredElementCollector(doc).OfClass(Family).ToElements()
)

# Alertas (Warnings) do modelo
warnings = []
try:
    warnings = list(doc.GetWarnings())
except Exception:
    pass

# Worksets de usuario (apenas se modelo for workshared)
worksets_list  = []
worksets_count = 0
try:
    if doc.IsWorkshared:
        worksets_list = list(
            FilteredWorksetCollector(doc)
            .OfKind(WorksetKind.UserWorkset)
            .ToWorksets()
        )
        worksets_count = len(worksets_list)
except Exception:
    pass


# ============================================================
# 3. ESTATISTICAS GERAIS
# ============================================================
output.print_md(u"---\n## 3. Estatisticas Gerais")

phases_count = 0
try:
    phases_count = len(list(doc.Phases))
except Exception:
    pass

stats_table = [
    [u"Total de Elementos (instancias)", str(len(all_elements))],
    [u"Familias Carregadas",             str(len(families))],
    [u"Folhas (Sheets)",                 str(len(sheets))],
    [u"Vistas (total)",                  str(len(all_views))],
    [u"Niveis / Pavimentos",             str(len(levels))],
    [u"Modelos Linkados (RVT)",          str(len(links))],
    [u"Fases do Projeto",                str(phases_count)],
    [u"Design Options",                  str(len(d_opts))],
    [u"Alertas (Warnings)",              str(len(warnings))],
    [u"Workshared",                      u"Sim" if doc.IsWorkshared else u"Nao"],
    [u"Worksets de Usuario",             str(worksets_count)],
]
output.print_table(stats_table, columns=[u"Metrica", u"Valor"])


# ============================================================
# 4. TOP CATEGORIAS DE ELEMENTOS
# ============================================================
output.print_md(u"---\n## 4. Top {} Categorias de Elementos".format(TOP_N))

# Agrupa IDs de elementos por nome de categoria
cats_map = defaultdict(list)
for el in all_elements:
    try:
        if el.Category:
            cats_map[el.Category.Name].append(el.Id)
    except Exception:
        pass

cat_counter = Counter({k: len(v) for k, v in cats_map.items()})
total_elems = sum(cat_counter.values())
max_cat_qty = cat_counter.most_common(1)[0][1] if cat_counter else 1

cats_table = []
for cat, qty in cat_counter.most_common(TOP_N):
    pct       = round(qty * 100.0 / total_elems, 1) if total_elems else 0
    sample_id = cats_map[cat][0] if cats_map[cat] else None
    link      = output.linkify(sample_id) if sample_id else u"—"
    cats_table.append([
        u8(cat),
        str(qty),
        u"{}%".format(pct),
        bar(qty, max_cat_qty),
        link,
    ])

if cats_table:
    output.print_table(cats_table, columns=[u"Categoria", u"Qtd", u"%", u"Distribuicao", u"Exemplo"])
else:
    output.print_md(u"_Nenhuma categoria encontrada._")
output.print_md(
    u"_Total: {} categorias | {} instancias_".format(len(cat_counter), total_elems)
)


# ============================================================
# 5. TOP FAMILIAS (por numero de instancias)
# ============================================================
output.print_md(u"---\n## 5. Top {} Familias por Instancias".format(TOP_N))

# FamilyName esta disponivel no Symbol de instancias ou no tipo do elemento
family_map = defaultdict(list)
for el in all_elements:
    try:
        fam_name = None
        # FamilyInstance tem atributo .Symbol com .FamilyName direto
        if hasattr(el, "Symbol") and el.Symbol is not None:
            fam_name = u8(el.Symbol.FamilyName)
        else:
            # Fallback: pega o tipo e verifica FamilyName
            typ = doc.GetElement(el.GetTypeId())
            if typ is not None and hasattr(typ, "FamilyName"):
                fam_name = u8(typ.FamilyName)
        if fam_name:
            family_map[fam_name].append(el.Id)
    except Exception:
        pass

fam_counter  = Counter({k: len(v) for k, v in family_map.items()})
total_fam_qty = sum(fam_counter.values())
max_fam_qty  = fam_counter.most_common(1)[0][1] if fam_counter else 1

fams_table = []
for fam, qty in fam_counter.most_common(TOP_N):
    pct       = round(qty * 100.0 / total_fam_qty, 1) if total_fam_qty else 0
    sample_id = family_map[fam][0] if family_map[fam] else None
    link      = output.linkify(sample_id) if sample_id else u"—"
    fams_table.append([u8(fam), str(qty), u"{}%".format(pct), bar(qty, max_fam_qty), link])

if fams_table:
    output.print_table(fams_table, columns=[u"Familia", u"Instancias", u"%", u"Distribuicao", u"Exemplo"])
else:
    output.print_md(u"_Nenhuma familia com instancias encontrada._")
output.print_md(u"_Total: {} familias unicas com instancias_".format(len(fam_counter)))


# ============================================================
# 6. TOP TIPOS DE FAMILIA (por numero de instancias)
# ============================================================
output.print_md(u"---\n## 6. Top {} Tipos de Familia por Instancias".format(TOP_N))

# label = "FamiliaName : TipoName" para identificacao inequivoca.
# Usa el.Symbol (FamilyInstances) como fonte primaria — mais confiavel em
# IronPython do que doc.GetElement(el.GetTypeId()) para familias de sistema.
type_map = defaultdict(list)
for el in all_elements:
    try:
        fam_name  = u""
        type_name = u""
        # FamilyInstance: Symbol tem FamilyName e Name diretamente
        if hasattr(el, "Symbol") and el.Symbol is not None:
            fam_name  = u8(el.Symbol.FamilyName)
            type_name = u8(el.Symbol.Name)
        else:
            # Elementos de sistema (paredes, pisos, etc.): usa GetTypeId
            type_id = el.GetTypeId()
            if type_id is not None and eid_int(type_id) > 0:
                typ = doc.GetElement(type_id)
                if typ is not None:
                    type_name = u8(typ.Name)
                    fam_name  = u8(getattr(typ, "FamilyName", u""))
        # Fallback: usa nome da categoria como familia quando nao ha FamilyName
        if not fam_name and el.Category:
            fam_name = u8(el.Category.Name)
        if not type_name:
            type_name = u8(el.Name) if hasattr(el, "Name") else u""
        if not type_name:
            continue
        label = u"{} : {}".format(fam_name, type_name) if fam_name else type_name
        type_map[label].append(el.Id)
    except Exception:
        pass

type_counter  = Counter({k: len(v) for k, v in type_map.items()})
total_type_qty = sum(type_counter.values())
max_type_qty  = type_counter.most_common(1)[0][1] if type_counter else 1

types_table = []
for label, qty in type_counter.most_common(TOP_N):
    pct       = round(qty * 100.0 / total_type_qty, 1) if total_type_qty else 0
    sample_id = type_map[label][0] if type_map[label] else None
    link      = output.linkify(sample_id) if sample_id else u"—"
    types_table.append([u8(label), str(qty), u"{}%".format(pct), bar(qty, max_type_qty), link])

if types_table:
    output.print_table(types_table, columns=[u"Familia : Tipo", u"Instancias", u"%", u"Distribuicao", u"Exemplo"])
else:
    output.print_md(u"_Nenhum tipo de familia com instancias encontrado._")
output.print_md(u"_Total: {} tipos unicos com instancias_".format(len(type_counter)))


# ============================================================
# 7. TOP FOLHAS (Sheets por numero de vistas/viewports)
# ============================================================
output.print_md(u"---\n## 7. Top {} Folhas por Numero de Vistas".format(TOP_N))

sheets_data = []
for sh in sheets:
    try:
        number   = u8(sh.SheetNumber)
        name     = u8(sh.Name)
        vp_count = len(list(sh.GetAllViewports()))
        sheets_data.append((number, name, vp_count, sh.Id))
    except Exception:
        pass

# Ordena por quantidade de vistas (decrescente)
sheets_data.sort(key=lambda x: x[2], reverse=True)
max_vp_qty   = sheets_data[0][2] if sheets_data else 1
total_vp_qty = sum(vp for _, _, vp, _ in sheets_data)

sheets_table = []
for number, name, vp_count, sh_id in sheets_data[:TOP_N]:
    pct  = round(vp_count * 100.0 / total_vp_qty, 1) if total_vp_qty else 0
    link = output.linkify(sh_id)
    sheets_table.append([
        u8(number),
        u8(name),
        str(vp_count),
        u"{}%".format(pct),
        bar(vp_count, max_vp_qty),
        link,
    ])

if sheets_table:
    output.print_table(sheets_table, columns=[u"Numero", u"Nome", u"Vistas", u"%", u"Distribuicao", u"Abrir"])
else:
    output.print_md(u"_Nenhuma folha encontrada no modelo._")
output.print_md(u"_Total de folhas: {}_".format(len(sheets)))


# ============================================================
# 8. VISTAS POR TIPO
# ============================================================
output.print_md(u"---\n## 8. Vistas por Tipo")

view_type_counter = Counter()
for v in all_views:
    try:
        vt = u8(str(v.ViewType))
        view_type_counter[vt] += 1
    except Exception:
        pass

max_vt_qty   = view_type_counter.most_common(1)[0][1] if view_type_counter else 1
total_vt_qty = sum(view_type_counter.values())
vt_table = [
    [u8(vt), str(cnt), u"{}%".format(round(cnt * 100.0 / total_vt_qty, 1) if total_vt_qty else 0), bar(cnt, max_vt_qty)]
    for vt, cnt in view_type_counter.most_common()
]
if vt_table:
    output.print_table(vt_table, columns=[u"Tipo de Vista", u"Quantidade", u"%", u"Distribuicao"])
else:
    output.print_md(u"_Nenhuma vista encontrada._")
output.print_md(u"_Total de vistas: {}_".format(len(all_views)))


# ============================================================
# 9. NIVEIS / PAVIMENTOS
# ============================================================
output.print_md(u"---\n## 9. Niveis / Pavimentos")

# Ordena do pavimento mais baixo para o mais alto (por elevacao interna)
levels_sorted = sorted(levels, key=lambda lv: lv.Elevation)

# Conta instancias associadas a cada nivel via LevelId
level_elem_counter = Counter()
for el in all_elements:
    try:
        lv_id = el.LevelId
        if lv_id is not None:
            level_elem_counter[eid_int(lv_id)] += 1
    except Exception:
        pass

lvl_table = []
for lv in levels_sorted:
    elev_m   = to_meters(lv.Elevation)
    elem_cnt = level_elem_counter.get(eid_int(lv.Id), 0)
    link     = output.linkify(lv.Id)
    lvl_table.append([
        u8(lv.Name),
        u"{} m".format(elev_m),
        str(elem_cnt),
        link,
    ])

if lvl_table:
    output.print_table(lvl_table, columns=[u"Nome", u"Elevacao", u"Elementos", u"Selecionar"])
else:
    output.print_md(u"_Nenhum nivel encontrado._")
output.print_md(u"_Total: {} niveis_".format(len(levels)))


# ============================================================
# 10. WORKSETS (apenas se modelo for workshared)
# ============================================================
output.print_md(u"---\n## 10. Worksets (Conjuntos de Trabalho)")

try:
    if doc.IsWorkshared and worksets_list:
        # Conta elementos por workset usando o WorksetId de cada instancia
        ws_elem_counter = Counter()
        for el in all_elements:
            try:
                ws_elem_counter[eid_int(el.WorksetId)] += 1
            except Exception:
                pass

        ws_table = []
        for ws in sorted(
            worksets_list,
            key=lambda w: ws_elem_counter.get(eid_int(w.Id), 0),
            reverse=True,
        ):
            cnt   = ws_elem_counter.get(eid_int(ws.Id), 0)
            owner = u8(ws.Owner) if ws.Owner else u"(aberto)"
            state = u"Aberto" if ws.IsOpen else u"Fechado"
            ws_table.append([u8(ws.Name), owner, str(cnt), state])

        if ws_table:
            output.print_table(ws_table, columns=[u"Workset", u"Proprietario", u"Elementos", u"Estado"])
        output.print_md(u"_Total: {} worksets de usuario_".format(len(worksets_list)))
    else:
        output.print_md(u"_Modelo nao e workshared (sem worksets de usuario)._")
except Exception as ex_ws:
    output.print_md(u"_Erro ao ler worksets: {}_".format(u8(str(ex_ws))))


# ============================================================
# 11. TOP ALERTAS (Warnings)
# ============================================================
output.print_md(u"---\n## 11. Top Alertas do Modelo")

if warnings:
    warn_counter = Counter()
    for w in warnings:
        try:
            desc = u8(str(w.GetDescriptionText()))[:100]
        except Exception:
            desc = u"(erro ao ler descricao)"
        warn_counter[desc] += 1

    total_warn_qty = len(warnings)
    warn_table = [
        [u8(desc), str(cnt), u"{}%".format(round(cnt * 100.0 / total_warn_qty, 1) if total_warn_qty else 0)]
        for desc, cnt in warn_counter.most_common(TOP_N)
    ]
    if warn_table:
        output.print_table(warn_table, columns=[u"Descricao do Alerta", u"Ocorrencias", u"%"])
    output.print_md(u"_Total de alertas: {}_".format(len(warnings)))
else:
    warn_counter = Counter()
    output.print_md(u"_Nenhum alerta encontrado no modelo._")


# ============================================================
# 12. MODELOS LINKADOS (RVT Links)
# ============================================================
output.print_md(u"---\n## 12. Modelos Linkados (RVT Links)")

if links:
    lk_table = []
    for lk in links:
        try:
            lk_type = doc.GetElement(lk.GetTypeId())
            status  = u8(str(lk_type.GetLinkedFileStatus())) if lk_type else u"N/D"
        except Exception:
            status = u"Erro"
        lk_table.append([u8(lk.Name), status, output.linkify(lk.Id)])
    if lk_table:
        output.print_table(lk_table, columns=[u"Nome do Link", u"Status", u"Selecionar"])
    output.print_md(u"_Total: {} modelos linkados_".format(len(links)))
else:
    output.print_md(u"_Nenhum modelo linkado encontrado._")


# ============================================================
# 13. FASES DO PROJETO
# ============================================================
output.print_md(u"---\n## 13. Fases do Projeto")

phases_table = []
try:
    for i, ph in enumerate(doc.Phases):
        phases_table.append([str(i + 1), u8(ph.Name), u8(ph.Id.ToString())])
except Exception:
    phases_table = [[u"—", u"Erro ao ler fases", u"—"]]

if phases_table:
    output.print_table(phases_table, columns=[u"Ordem", u"Nome", u"ID"])
else:
    output.print_md(u"_Nenhuma fase encontrada._")


# ============================================================
# EXPORTACAO CSV — Top Categorias
# ============================================================
for path in [os.path.join(RUN_DIR, "top_categorias.csv"),
             os.path.join(LATEST_DIR, "top_categorias.csv")]:
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(csv_row("run_stamp", "model", "categoria", "quantidade", "percentual"))
    for cat, qty in cat_counter.most_common():
        pct = round(qty * 100.0 / total_elems, 2) if total_elems else 0
        w.writerow(csv_row(run_stamp, model, cat, qty, pct))
    f.close()


# EXPORTACAO CSV — Top Familias
for path in [os.path.join(RUN_DIR, "top_familias.csv"),
             os.path.join(LATEST_DIR, "top_familias.csv")]:
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(csv_row("run_stamp", "model", "familia", "instancias"))
    for fam, qty in fam_counter.most_common():
        w.writerow(csv_row(run_stamp, model, fam, qty))
    f.close()


# EXPORTACAO CSV — Top Folhas
for path in [os.path.join(RUN_DIR, "top_folhas.csv"),
             os.path.join(LATEST_DIR, "top_folhas.csv")]:
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(csv_row("run_stamp", "model", "numero", "nome", "vistas"))
    for number, name, vp_count, _ in sheets_data:
        w.writerow(csv_row(run_stamp, model, number, name, vp_count))
    f.close()


# ============================================================
# EXPORTACAO JSON — Snapshot completo do modelo
# ============================================================
resumo = {
    "run_stamp": u8(run_stamp),
    "model":     u8(model),
    "revit_version": {
        "number": u8(str(app.VersionNumber)),
        "name":   u8(str(app.VersionName)),
        "build":  u8(str(app.VersionBuild)),
        "lang":   u8(str(app.Language)),
    },
    "project_info": {
        "name":       u8(proj.Name),
        "number":     u8(proj.Number),
        "client":     u8(proj.ClientName),
        "address":    u8(proj.Address),
        "status":     u8(proj.Status),
        "author":     u8(proj.Author),
        "issue_date": u8(proj.IssueDate),
    },
    "stats": {
        "total_elements":    len(all_elements),
        "unique_categories": len(cat_counter),
        "unique_families":   len(fam_counter),
        "unique_types":      len(type_counter),
        "families_loaded":   len(families),
        "sheets":            len(sheets),
        "views":             len(all_views),
        "levels":            len(levels),
        "links":             len(links),
        "phases":            phases_count,
        "warnings":          len(warnings),
        "design_options":    len(d_opts),
        "is_workshared":     bool(doc.IsWorkshared),
        "worksets":          worksets_count,
    },
    "top_categories": [
        {"category": u8(k), "count": v}
        for k, v in cat_counter.most_common(TOP_N)
    ],
    "top_families": [
        {"family": u8(k), "count": v}
        for k, v in fam_counter.most_common(TOP_N)
    ],
    "top_types": [
        {"type": u8(k), "count": v}
        for k, v in type_counter.most_common(TOP_N)
    ],
    "top_sheets": [
        {"number": u8(n), "name": u8(nm), "viewports": vc}
        for n, nm, vc, _ in sheets_data[:TOP_N]
    ],
    "view_types": [
        {"type": u8(vt), "count": cnt}
        for vt, cnt in view_type_counter.most_common()
    ],
    "levels": [
        {"name": u8(lv.Name), "elevation_m": to_meters(lv.Elevation)}
        for lv in levels_sorted
    ],
    "top_warnings": [
        {"description": u8(d), "count": c}
        for d, c in warn_counter.most_common(TOP_N)
    ],
}

save_json(
    resumo,
    os.path.join(RUN_DIR,    "resumo_modelo.json"),
    os.path.join(LATEST_DIR, "resumo_modelo.json"),
)


# ============================================================
# RODAPE — caminhos de saida no painel HTML
# ============================================================
output.print_md(u"---\n## Arquivos Gerados")
output.print_table([
    [u"Historico (esta execucao)", u8(RUN_DIR)],
    [u"Ultima execucao (latest)",  u8(LATEST_DIR)],
], columns=[u"Destino", u"Caminho"])
output.print_md(
    u"> **Arquivos:** `top_categorias.csv` | `top_familias.csv` | "
    u"`top_folhas.csv` | `resumo_modelo.json`"
)


# ============================================================
# SAIDA NO CONSOLE (complementar ao painel HTML)
# ============================================================
print("=" * 65)
print("  RESUMO DO MODELO CONCLUIDO")
print("=" * 65)
print("  Modelo         : {}".format(model))
print("  Revit          : {} ({})".format(app.VersionNumber, app.VersionName))
print("  Data           : {}  Hora: {}".format(day, ts))
print("-" * 65)
print("  Elementos      : {}".format(len(all_elements)))
print("  Categorias     : {}".format(len(cat_counter)))
print("  Familias       : {}".format(len(fam_counter)))
print("  Tipos          : {}".format(len(type_counter)))
print("  Fam. carregadas: {}".format(len(families)))
print("  Folhas         : {}".format(len(sheets)))
print("  Vistas         : {}".format(len(all_views)))
print("  Niveis         : {}".format(len(levels)))
print("  Alertas        : {}".format(len(warnings)))
print("  Links RVT      : {}".format(len(links)))
print("  Workshared     : {}".format("Sim" if doc.IsWorkshared else "Nao"))
print("-" * 65)
print("  [Historico]    {}".format(RUN_DIR))
print("  [Latest]       {}".format(LATEST_DIR))
print("=" * 65)
print("PROCESSO REALIZADO COM SUCESSO")
