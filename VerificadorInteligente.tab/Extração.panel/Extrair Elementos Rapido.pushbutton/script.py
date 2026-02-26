# -*- coding: utf-8 -*-

from pyrevit import revit
from Autodesk.Revit.DB import *
import csv, os, datetime

doc = revit.doc

# =========================
# PARAMS
# =========================
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"
MAX_PARAM_LEN = 300

# Categorias relevantes (ajustável conforme seu padrão de agência)
ALLOW_CATEGORIES = set([
    # Elétrica / MEP
    u"Conduites", u"Conexões do conduite", u"Fiação", u"Circuitos elétricos",
    u"Luminárias", u"Dispositivos elétricos", u"Equipamentos elétricos",
    u"Bandejas de cabos", u"Conexões da bandeja de cabos",
    u"Eletrocalhas", u"Conexões da eletrocalha",

    # Identificadores/Tags (aparecem muito no seu modelo)
    u"Identificadores de luminária", u"Identificadores de dispositivos elétricos",

    # Espaços úteis p/ features
    u"Ambientes", u"Salas", u"Níveis", u"Linhas de centro",
])

# Parâmetros-chave (Silver rápido)
TOP_PARAMS = set([
    u"Mark", u"Marca", u"Comentários", u"Comment",
    u"Nome do sistema", u"Classificação do sistema", u"System Classification",
    u"Número do circuito", u"Circuit Number",
    u"Painel", u"Panel",
    u"Tensão", u"Voltage",
    u"Carga", u"Load",
    u"Fabricante", u"Manufacturer",
    u"Modelo", u"Model",
])

# =========================

ts = datetime.datetime.now().strftime("%H%M%S")
day = datetime.datetime.now().strftime("%Y-%m-%d")
model = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR = os.path.join(ROOT_DIR, model_safe, day)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)

elements_path = os.path.join(RUN_DIR, "elements_rapido_{}.csv".format(ts))
params_path   = os.path.join(RUN_DIR, "params_top_{}.csv".format(ts))

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

def eid_int(eid):
    try: return eid.IntegerValue
    except: pass
    try: return eid.Value
    except: pass
    try: return int(eid.ToString())
    except: return None

def bbox_to_str(bb):
    if not bb: return ("","")
    mn, mx = bb.Min, bb.Max
    return ("{},{},{}".format(mn.X,mn.Y,mn.Z), "{},{},{}".format(mx.X,mx.Y,mx.Z))

def get_level_name(el):
    try:
        if el.LevelId and el.LevelId != ElementId.InvalidElementId:
            lv = doc.GetElement(el.LevelId)
            return lv.Name if lv else ""
    except:
        return ""
    return ""

def get_family_type(el):
    try:
        t = doc.GetElement(el.GetTypeId())
        if not t: return ("","")
        fam = ""
        try: fam = t.FamilyName
        except: fam = ""
        return (fam, t.Name)
    except:
        return ("","")

def param_to_str(p):
    st = p.StorageType
    try:
        if st == StorageType.String: return p.AsString() or ""
        if st == StorageType.Integer: return str(p.AsInteger())
        if st == StorageType.Double: return str(p.AsDouble())
        if st == StorageType.ElementId:
            eid = p.AsElementId()
            if eid == ElementId.InvalidElementId: return ""
            ref = doc.GetElement(eid)
            if ref:
                try: return "{}:{}".format(eid_int(eid), ref.Name)
                except: return str(eid_int(eid))
            return str(eid_int(eid))
    except:
        return ""
    return ""

def write_top_params(unique_id, element_id, el, writer):
    for p in el.Parameters:
        try:
            name = p.Definition.Name
            if name not in TOP_PARAMS:
                continue
            val = param_to_str(p)
            if not val:
                continue
            if len(val) > MAX_PARAM_LEN:
                val = val[:MAX_PARAM_LEN]
            writer.writerow([u8(model), u8(unique_id), u8(element_id), u8(name), u8(val)])
        except:
            continue

# CSVs com BOM
fe = open(elements_path, "wb"); fe.write(u"\ufeff".encode("utf-8"))
fp = open(params_path, "wb");   fp.write(u"\ufeff".encode("utf-8"))
ew = csv.writer(fe); pw = csv.writer(fp)

ew.writerow(["model","element_id","unique_id","category","family","type","level","bbox_min","bbox_max"])
pw.writerow(["model","unique_id","element_id","param_name","value_str"])

elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

kept = 0
for el in elements:
    try:
        cat = el.Category.Name if el.Category else ""
        # filtra por categoria
        if cat and (cat not in ALLOW_CATEGORIES):
            continue

        el_id = eid_int(el.Id)
        el_id_str = str(el_id) if el_id is not None else ""

        fam, typ = get_family_type(el)
        lvl = get_level_name(el)
        bbmin, bbmax = bbox_to_str(el.get_BoundingBox(None))

        ew.writerow([u8(model), u8(el_id_str), u8(el.UniqueId), u8(cat), u8(fam), u8(typ), u8(lvl), u8(bbmin), u8(bbmax)])
        write_top_params(el.UniqueId, el_id_str, el, pw)

        kept += 1
    except:
        continue

fe.close(); fp.close()

print("✅ Extrair Elementos (Rápido) concluído")
print("Pasta:", RUN_DIR)
print("elements:", elements_path)
print("params_top:", params_path)
print("Elementos mantidos:", kept)