# -*- coding: utf-8 -*-

from pyrevit import revit
from Autodesk.Revit.DB import *
from collections import Counter
import os, csv, datetime
import json

doc = revit.doc

# =========================
# PARAMS (AJUSTE AQUI)
# =========================
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop")  # default: Área de Trabalho
ROOT_FOLDER_NAME = "revit_dump"
TOP_N = 25
SHEET_KEYWORDS = [
    ("ELETRICA", ["EL", "ELE", "ELÉTR", "E-"]),
    ("INFRA",    ["INF", "TI", "DADOS", "REDE", "RACK", "CABEAMENTO"]),
    ("SEGURANCA",["SEG", "CFTV", "ALARME", "ACESSO", "CAMERA", "CÂMERA"]),
]
# =========================

ts = datetime.datetime.now().strftime("%H%M%S")
day = datetime.datetime.now().strftime("%Y-%m-%d")
model = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR = os.path.join(ROOT_DIR, model_safe, day)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)

out_topcats = os.path.join(RUN_DIR, "top_categorias_{}.csv".format(ts))
out_sheets  = os.path.join(RUN_DIR, "folhas_resumo_{}.csv".format(ts))
out_json    = os.path.join(RUN_DIR, "model_profile_{}.json".format(ts))


def u8(x):
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
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    return f


def norm_upper(s):
    try:
        if isinstance(s, unicode):
            return s.upper()
        return unicode(s, errors="ignore").upper()
    except:
        try:
            return str(s).upper()
        except:
            return ""


# -------------------------
# 1) Elementos e categorias
# -------------------------
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

cats = []
for el in elements:
    try:
        if el.Category:
            cats.append(el.Category.Name)
    except:
        pass

counter = Counter(cats)

# -------------------------
# 2) Sheets e heurística
# -------------------------
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

def classify_sheet(sh):
    name = norm_upper((sh.SheetNumber or "") + " " + (sh.Name or ""))
    hits = []
    for label, keys in SHEET_KEYWORDS:
        for k in keys:
            if k in name:
                hits.append(label)
                break
    if not hits:
        return "OUTROS"
    # se bater mais de um, concatena (ex: ELETRICA+INFRA)
    hits = sorted(list(set(hits)))
    return "+".join(hits)

sheet_classes = Counter()
for sh in sheets:
    try:
        sheet_classes[classify_sheet(sh)] += 1
    except:
        pass

# -------------------------
# 3) Salvar top categorias
# -------------------------
f1 = write_bom_csv(out_topcats)
w1 = csv.writer(f1)
w1.writerow(["model", "total_elementos", "categoria", "qtd"])
for cat, qtd in counter.most_common(TOP_N):
    w1.writerow([u8(model), u8(len(elements)), u8(cat), u8(qtd)])
f1.close()

# -------------------------
# 4) Salvar resumo de folhas
# -------------------------
f2 = write_bom_csv(out_sheets)
w2 = csv.writer(f2)
w2.writerow(["model", "sheet_number", "sheet_name", "classificacao"])
for sh in sheets:
    try:
        w2.writerow([u8(model), u8(sh.SheetNumber), u8(sh.Name), u8(classify_sheet(sh))])
    except:
        continue
f2.close()

# -------------------------
# 5) Salvar json de perfil
# -------------------------
profile = {
    "model": model,
    "run_dir": RUN_DIR,
    "day": day,
    "time": ts,
    "total_elements": len(elements),
    "unique_categories": len(counter),
    "top_categories": [{"category": c, "count": n} for c, n in counter.most_common(TOP_N)],
    "total_sheets": len(sheets),
    "sheet_classes": dict(sheet_classes),
}

# json no IronPython: garantir utf-8
f3 = open(out_json, "wb")
f3.write(u"\ufeff".encode("utf-8"))
f3.write(u8(json.dumps(profile, ensure_ascii=False, indent=2)))
f3.close()

# -------------------------
# Output
# -------------------------
print("✅ Resumo do Modelo gerado em:")
print(RUN_DIR)
print("•", out_topcats)
print("•", out_sheets)
print("•", out_json)
print("")
print("📌 Totais:")
print("Elementos:", len(elements))
print("Categorias unicas:", len(counter))
print("Folhas:", len(sheets))
print("")
print("📊 Top categorias:")
for cat, qtd in counter.most_common(10):
    print(" - {}: {}".format(cat, qtd))
print("")
print("📄 Folhas por classificacao:")
for k, v in sheet_classes.most_common():
    print(" - {}: {}".format(k, v))