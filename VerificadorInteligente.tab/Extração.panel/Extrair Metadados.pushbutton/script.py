# -*- coding: utf-8 -*-

from pyrevit import revit
from Autodesk.Revit.DB import *
import csv, os, datetime

doc = revit.doc

# =========================================
# PARAMS (AJUSTE AQUI)
# =========================================
# Opções:
# 1) Desktop do usuário (recomendado)
# 2) Pasta fixa (ex: r"D:\bim_exports")
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop")  # <- default: Área de Trabalho
ROOT_FOLDER_NAME = "revit_dump"  # <- nome da pasta raiz
# =========================================

ts = datetime.datetime.now().strftime("%H%M%S")
day = datetime.datetime.now().strftime("%Y-%m-%d")
model = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR = os.path.join(ROOT_DIR, model_safe, day)

if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)

out = os.path.join(RUN_DIR, "sheets_views_{}.csv".format(ts))


def u8(x):
    """Converte para bytes UTF-8 (Python 2 / IronPython)."""
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


def safe(x):
    return x if x else ""


sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

f = open(out, "wb")
f.write(u"\ufeff".encode("utf-8"))  # BOM p/ Excel
w = csv.writer(f)

w.writerow([
    "model", "sheet_number", "sheet_name", "sheet_id",
    "view_id", "view_name", "view_type", "view_discipline"
])

for sh in sheets:
    try:
        vports = sh.GetAllViewports()
        for vpid in vports:
            vp = doc.GetElement(vpid)
            v = doc.GetElement(vp.ViewId)

            discipline = ""
            try:
                discipline = str(v.Discipline)
            except:
                discipline = ""

            w.writerow([
                u8(model),
                u8(safe(sh.SheetNumber)),
                u8(safe(sh.Name)),
                u8(sh.Id.ToString()),
                u8(v.Id.ToString()),
                u8(safe(v.Name)),
                u8(str(v.ViewType)),
                u8(discipline)
            ])
    except:
        continue

f.close()

print("✅ Export Sheets/Views gerado em:")
print(out)
print("Sheets encontrados:", len(sheets))
print("Pasta do dia:", RUN_DIR)