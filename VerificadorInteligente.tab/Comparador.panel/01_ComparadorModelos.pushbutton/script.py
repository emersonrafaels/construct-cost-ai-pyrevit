# -*- coding: utf-8 -*-

# ============================================================
# COMPARADOR DE MODELOS -- Diff direto entre dois arquivos .rvt
# ============================================================
# Seleciona dois arquivos .rvt via explorer, abre cada um
# desanexado do servidor central, extrai os elementos em memória
# e compara, produzindo:
#
#   comparacao.csv   -- uma linha por elemento/diferença com status:
#                       "adicionado" | "removido" | "alterado" | "inalterado"
#   comparacao.json  -- resumo + diff estruturado
#
# Campos comparados: category, family, type, level
#
# Saída:
#   Desktop/RevitScan/Comparadores/<carimbo>/
#   Desktop/RevitScan/Comparadores/latest/
# ============================================================

from pyrevit import forms, script, output
from Autodesk.Revit.DB import FilteredElementCollector, ElementId
import csv, json, os, datetime, sys

from path_utils import (
    BASE_DIR,
    ROOT_FOLDER_NAME,
    make_run_stamp,
    ensure_dirs,
    write_bom_csv,
    write_json_file,
)
from revit_open_utils import DialogSuppressor, open_rvt_detached, close_doc
from extracao_lib import (
    get_family_type,
    get_level_name,
    get_workset_name,
    get_phase_names,
    eid_int,
    safe_str,
)

# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Campos utilizados para detectar "alterado" vs "inalterado"
CAMPOS_COMPARAR = ["category", "family", "type", "level"]

# Incluir elementos inalterados no CSV de saída?
INCLUIR_INALTERADOS = True

# ============================================================
# EXTRAÇÃO EM MEMÓRIA
# ============================================================

def _extrair_elementos(doc):
    """
    Extrai elementos físicos do documento Revit aberto.

    Retorna dict { unique_id: { category, family, type, level,
                                workset, phase_created, phase_demolished } }
    """
    rows = {}
    elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())
    for el in elements:
        try:
            uid = el.UniqueId
            if not uid:
                continue
            cat = safe_str(el.Category.Name) if el.Category else u""
            fam, typ = get_family_type(el)
            lvl = get_level_name(el)
            wks = get_workset_name(el)
            phase_c, phase_d = get_phase_names(el)
            rows[uid] = {
                "category":        cat,
                "family":          fam,
                "type":            typ,
                "level":           lvl,
                "workset":         wks,
                "phase_created":   phase_c,
                "phase_demolished": phase_d,
            }
        except:
            continue
    return rows


# ============================================================
# COMPARAÇÃO
# ============================================================

def _sv(row, campo):
    """Valor seguro de campo em row."""
    return (row.get(campo) or u"").strip()


def _compare(rows_a, rows_b, incluir_inalterados=True):
    """
    Compara dois dicts {uid: row_dict}.
    Retorna lista de dicts com colunas:
        status, unique_id, category, family, type, level,
        workset, phase_created, phase_demolished,
        campo_alterado, valor_antes, valor_depois
    """
    uids_a = set(rows_a.keys())
    uids_b = set(rows_b.keys())
    diff   = []

    # --- Removidos (em A mas não em B) -------------------------
    for uid in sorted(uids_a - uids_b):
        row = rows_a[uid]
        diff.append({
            "status":           u"removido",
            "unique_id":        uid,
            "category":         _sv(row, "category"),
            "family":           _sv(row, "family"),
            "type":             _sv(row, "type"),
            "level":            _sv(row, "level"),
            "workset":          _sv(row, "workset"),
            "phase_created":    _sv(row, "phase_created"),
            "phase_demolished": _sv(row, "phase_demolished"),
            "campo_alterado":   u"",
            "valor_antes":      u"",
            "valor_depois":     u"",
        })

    # --- Adicionados (em B mas não em A) -----------------------
    for uid in sorted(uids_b - uids_a):
        row = rows_b[uid]
        diff.append({
            "status":           u"adicionado",
            "unique_id":        uid,
            "category":         _sv(row, "category"),
            "family":           _sv(row, "family"),
            "type":             _sv(row, "type"),
            "level":            _sv(row, "level"),
            "workset":          _sv(row, "workset"),
            "phase_created":    _sv(row, "phase_created"),
            "phase_demolished": _sv(row, "phase_demolished"),
            "campo_alterado":   u"",
            "valor_antes":      u"",
            "valor_depois":     u"",
        })

    # --- Presentes nos dois ------------------------------------
    for uid in sorted(uids_a & uids_b):
        row_a = rows_a[uid]
        row_b = rows_b[uid]

        campos_alterados = []
        for campo in CAMPOS_COMPARAR:
            va = _sv(row_a, campo)
            vb = _sv(row_b, campo)
            if va != vb:
                campos_alterados.append((campo, va, vb))

        if campos_alterados:
            # Uma linha por campo alterado
            for campo, va, vb in campos_alterados:
                diff.append({
                    "status":           u"alterado",
                    "unique_id":        uid,
                    "category":         _sv(row_b, "category"),
                    "family":           _sv(row_b, "family"),
                    "type":             _sv(row_b, "type"),
                    "level":            _sv(row_b, "level"),
                    "workset":          _sv(row_b, "workset"),
                    "phase_created":    _sv(row_b, "phase_created"),
                    "phase_demolished": _sv(row_b, "phase_demolished"),
                    "campo_alterado":   campo,
                    "valor_antes":      va,
                    "valor_depois":     vb,
                })
        elif incluir_inalterados:
            diff.append({
                "status":           u"inalterado",
                "unique_id":        uid,
                "category":         _sv(row_b, "category"),
                "family":           _sv(row_b, "family"),
                "type":             _sv(row_b, "type"),
                "level":            _sv(row_b, "level"),
                "workset":          _sv(row_b, "workset"),
                "phase_created":    _sv(row_b, "phase_created"),
                "phase_demolished": _sv(row_b, "phase_demolished"),
                "campo_alterado":   u"",
                "valor_antes":      u"",
                "valor_depois":     u"",
            })

    return diff


# ============================================================
# PERSISTÊNCIA
# ============================================================

def _u8(x):
    if sys.version_info[0] >= 3:
        return str(x) if x is not None else u""
    if x is None:
        return ""
    try:
        _unicode = unicode  # noqa: F821
        if isinstance(x, _unicode):
            return x.encode("utf-8")
        return str(x)
    except:
        return ""


COLS = [
    "status", "unique_id", "category", "family", "type", "level",
    "workset", "phase_created", "phase_demolished",
    "campo_alterado", "valor_antes", "valor_depois",
]


def _save_csv(path, diff):
    with write_bom_csv(path) as f:
        w = csv.writer(f)
        w.writerow([_u8(c) for c in COLS])
        for row in diff:
            w.writerow([_u8(row.get(c, u"")) for c in COLS])


def _save_json(path, rvt_a, rvt_b, diff, resumo):
    write_json_file(path, {
        "arquivo_a":  rvt_a,
        "arquivo_b":  rvt_b,
        "resumo":     resumo,
        "diferencas": [d for d in diff if d["status"] != "inalterado"],
    })


# ============================================================
# MAIN
# ============================================================

app = __revit__.Application
out = output.get_output()

out.print_md(u"## Comparador de Modelos Revit")
out.print_md(u"Selecione os dois arquivos .rvt para comparar.")

# --- Passo 1: seleciona os dois .rvt -------------------------
rvt_a = forms.pick_file(
    file_ext="rvt",
    title=u"Selecione o arquivo A (versão ANTERIOR)",
)
if not rvt_a:
    script.exit()

rvt_b = forms.pick_file(
    file_ext="rvt",
    title=u"Selecione o arquivo B (versão ATUAL)",
)
if not rvt_b:
    script.exit()

out.print_md(u"**A (anterior):** `{}`".format(rvt_a))
out.print_md(u"**B (atual):**    `{}`".format(rvt_b))
out.print_md(u"---")

# --- Passo 2: abre e extrai elementos de cada .rvt -----------
rows_a = {}
rows_b = {}
doc_a  = None
doc_b  = None

try:
    with DialogSuppressor(app):

        out.print_md(u"Abrindo arquivo A...")
        doc_a = open_rvt_detached(app, rvt_a)
        out.print_md(u"Extraindo elementos de A...")
        rows_a = _extrair_elementos(doc_a)
        out.print_md(u"- A: **{}** elemento(s)".format(len(rows_a)))
        close_doc(doc_a)
        doc_a = None

        out.print_md(u"Abrindo arquivo B...")
        doc_b = open_rvt_detached(app, rvt_b)
        out.print_md(u"Extraindo elementos de B...")
        rows_b = _extrair_elementos(doc_b)
        out.print_md(u"- B: **{}** elemento(s)".format(len(rows_b)))
        close_doc(doc_b)
        doc_b = None

except Exception as ex:
    close_doc(doc_a)
    close_doc(doc_b)
    forms.alert(
        u"Erro ao abrir/extrair modelo:\n\n{}".format(str(ex)),
        title=u"Comparador de Modelos",
        exitscript=True,
    )

# --- Passo 3: compara ----------------------------------------
out.print_md(u"Comparando modelos...")
diff = _compare(rows_a, rows_b, incluir_inalterados=INCLUIR_INALTERADOS)

n_add  = sum(1 for d in diff if d["status"] == u"adicionado")
n_rem  = sum(1 for d in diff if d["status"] == u"removido")
n_alt  = sum(1 for d in diff if d["status"] == u"alterado")
n_ign  = sum(1 for d in diff if d["status"] == u"inalterado")

# unique_ids com pelo menos um campo alterado
uids_alt = len(set(d["unique_id"] for d in diff if d["status"] == u"alterado"))

resumo = {
    "total_A":              len(rows_a),
    "total_B":              len(rows_b),
    "adicionados":          n_add,
    "removidos":            n_rem,
    "elementos_alterados":  uids_alt,
    "linhas_alteradas":     n_alt,
    "inalterados":          n_ign,
}

# --- Passo 4: salva arquivos ---------------------------------
now        = datetime.datetime.now()
run_stamp  = make_run_stamp(now)
plugin_dir = "Comparadores"

run_dir    = os.path.join(BASE_DIR, ROOT_FOLDER_NAME, plugin_dir, run_stamp)
latest_dir = os.path.join(BASE_DIR, ROOT_FOLDER_NAME, plugin_dir, "latest")
ensure_dirs(run_dir, latest_dir)

for dest in [run_dir, latest_dir]:
    _save_csv(os.path.join(dest, "comparacao.csv"), diff)
    _save_json(os.path.join(dest, "comparacao.json"), rvt_a, rvt_b, diff, resumo)

# --- Passo 5: exibe resultado --------------------------------
out.print_md(u"---")
out.print_md(u"## Resultado da Comparação")
out.print_md(
    u"| Adicionados | Removidos | Elem. Alterados | Inalterados |\n"
    u"|:-----------:|:---------:|:---------------:|:-----------:|\n"
    u"| **{}** | **{}** | **{}** | **{}** |".format(
        n_add, n_rem, uids_alt, n_ign)
)

if n_rem > 0:
    out.print_md(u"\n### Removidos ({})".format(n_rem))
    out.print_md(u"| unique_id | category | family | type | level |")
    out.print_md(u"|-----------|----------|--------|------|-------|")
    for d in [x for x in diff if x["status"] == u"removido"][:50]:
        out.print_md(u"| {} | {} | {} | {} | {} |".format(
            d["unique_id"], d["category"], d["family"], d["type"], d["level"]))

if n_add > 0:
    out.print_md(u"\n### Adicionados ({})".format(n_add))
    out.print_md(u"| unique_id | category | family | type | level |")
    out.print_md(u"|-----------|----------|--------|------|-------|")
    for d in [x for x in diff if x["status"] == u"adicionado"][:50]:
        out.print_md(u"| {} | {} | {} | {} | {} |".format(
            d["unique_id"], d["category"], d["family"], d["type"], d["level"]))

if n_alt > 0:
    out.print_md(u"\n### Alterados ({} elemento(s), {} linha(s))".format(uids_alt, n_alt))
    out.print_md(u"| unique_id | campo | antes | depois |")
    out.print_md(u"|-----------|-------|-------|--------|")
    for d in [x for x in diff if x["status"] == u"alterado"][:100]:
        out.print_md(u"| {} | {} | {} | {} |".format(
            d["unique_id"], d["campo_alterado"], d["valor_antes"], d["valor_depois"]))

out.print_md(u"---")
out.print_md(u"**Arquivos salvos em:** `{}`".format(run_dir))

if n_add == 0 and n_rem == 0 and uids_alt == 0:
    out.print_md(
        u"\n> **Os modelos são idênticos** nos campos comparados "
        u"(category, family, type, level)."
    )

forms.alert(
    u"Comparação concluída!\n\n"
    u"  Adicionados      : {}\n"
    u"  Removidos        : {}\n"
    u"  Elem. alterados  : {}\n"
    u"  Inalterados      : {}\n\n"
    u"Saída:\n  {}".format(n_add, n_rem, uids_alt, n_ign, run_dir),
    title=u"Comparador de Modelos",
    warn_icon=(n_add + n_rem + uids_alt > 0),
)
