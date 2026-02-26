# -*- coding: utf-8 -*-

from pyrevit import revit
from Autodesk.Revit.DB import *
import csv, os, datetime

doc = revit.doc

# =========================================
# PARAMS (AJUSTE AQUI)
# =========================================
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop")  # <- default: Área de Trabalho
ROOT_FOLDER_NAME = "revit_dump"  # <- nome da pasta raiz
MAX_PARAM_LEN = 500  # corta textos enormes (evita explosão)
# =========================================

ts = datetime.datetime.now().strftime("%H%M%S")
day = datetime.datetime.now().strftime("%Y-%m-%d")
model = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR = os.path.join(ROOT_DIR, model_safe, day)

if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)

elements_path = os.path.join(RUN_DIR, "elements_{}.csv".format(ts))
params_path   = os.path.join(RUN_DIR, "params_{}.csv".format(ts))


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


def eid_int(eid):
    """Compatível com diferentes versões do Revit/IronPython."""
    try:
        return eid.IntegerValue
    except:
        pass
    try:
        return eid.Value
    except:
        pass
    try:
        return int(eid.ToString())
    except:
        return None


def bbox_to_str(bb):
    if not bb:
        return ("", "")
    mn = bb.Min
    mx = bb.Max
    return ("{},{},{}".format(mn.X, mn.Y, mn.Z),
            "{},{},{}".format(mx.X, mx.Y, mx.Z))


def location_to_str(el):
    loc = el.Location
    if not loc:
        return ("", "")
    if isinstance(loc, LocationPoint):
        p = loc.Point
        return ("POINT", "{},{},{}".format(p.X, p.Y, p.Z))
    if isinstance(loc, LocationCurve):
        c = loc.Curve
        p0 = c.GetEndPoint(0)
        p1 = c.GetEndPoint(1)
        return ("CURVE", "{} -> {}".format(
            "{},{},{}".format(p0.X, p0.Y, p0.Z),
            "{},{},{}".format(p1.X, p1.Y, p1.Z)
        ))
    return ("", "")


def get_level_name(el):
    try:
        if el.LevelId and el.LevelId != ElementId.InvalidElementId:
            lv = doc.GetElement(el.LevelId)
            return lv.Name if lv else ""
    except:
        pass
    return ""


def get_family_type(el):
    try:
        t = doc.GetElement(el.GetTypeId())
        if not t:
            return ("", "")
        fam = ""
        try:
            fam = t.FamilyName
        except:
            fam = ""
        return (fam, t.Name)
    except:
        return ("", "")


def param_to_str(p):
    st = p.StorageType
    try:
        if st == StorageType.String:
            return p.AsString() or ""
        if st == StorageType.Integer:
            return str(p.AsInteger())
        if st == StorageType.Double:
            return str(p.AsDouble())  # raw
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


def write_params(owner_unique_id, owner_id, scope, el, writer, max_len):
    for p in el.Parameters:
        try:
            name = p.Definition.Name
            st = str(p.StorageType)
            val = param_to_str(p)

            if not val:
                continue

            if len(val) > max_len:
                val = val[:max_len]

            is_shared = "0"
            guid = ""
            try:
                if p.IsShared:
                    is_shared = "1"
                    guid = str(p.GUID)
            except:
                pass

            writer.writerow([
                u8(model),
                u8(owner_unique_id),
                u8(owner_id),
                u8(scope),
                u8(name),
                u8(st),
                u8(val),
                u8(is_shared),
                u8(guid)
            ])
        except:
            continue


elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

fe = open(elements_path, "wb")
fp = open(params_path, "wb")

# BOM p/ Excel reconhecer UTF-8
fe.write(u"\ufeff".encode("utf-8"))
fp.write(u"\ufeff".encode("utf-8"))

ew = csv.writer(fe)
pw = csv.writer(fp)

ew.writerow([
    "model","element_id","unique_id","category",
    "family","type","level","loc_kind","loc_value",
    "bbox_min","bbox_max"
])

pw.writerow([
    "model","unique_id","element_id","scope",
    "param_name","storage_type","value_str",
    "is_shared","guid"
])

ok = 0
skipped = 0

for el in elements:
    try:
        cat = safe(el.Category.Name) if el.Category else ""
        fam, typ = get_family_type(el)
        lvl = get_level_name(el)
        loc_kind, loc_val = location_to_str(el)
        bbmin, bbmax = bbox_to_str(el.get_BoundingBox(None))

        el_id = eid_int(el.Id)
        el_id_str = str(el_id) if el_id is not None else ""

        ew.writerow([
            u8(model),
            u8(el_id_str),
            u8(el.UniqueId),
            u8(cat),
            u8(fam),
            u8(typ),
            u8(lvl),
            u8(loc_kind),
            u8(loc_val),
            u8(bbmin),
            u8(bbmax)
        ])

        write_params(el.UniqueId, el_id_str, "instance", el, pw, MAX_PARAM_LEN)

        try:
            t = doc.GetElement(el.GetTypeId())
            if t:
                write_params(el.UniqueId, el_id_str, "type", t, pw, MAX_PARAM_LEN)
        except:
            pass

        ok += 1
    except:
        skipped += 1
        continue

fe.close()
fp.close()

print("✅ Dump concluído na pasta do dia:")
print(RUN_DIR)
print("elements:", elements_path)
print("params:", params_path)
print("Total coletado:", ok)
print("Total pulado:", skipped)
print("Total no collector:", len(elements))