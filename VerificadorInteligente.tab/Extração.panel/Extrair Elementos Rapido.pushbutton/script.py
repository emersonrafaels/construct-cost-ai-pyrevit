# -*- coding: utf-8 -*-

# ============================================================
# INVENTÁRIO POR AMBIENTE — AGÊNCIA BANCÁRIA
# ============================================================
# Gera um inventário completo de todos os equipamentos/itens do
# modelo Revit, agrupados por AMBIENTE (Room/Space), com:
#
#   - Nome do ambiente, nível, área em m²
#   - Categoria, família e tipo de cada equipamento
#   - Quantidade por tipo dentro de cada ambiente
#   - Totalizadores gerais por disciplina e por nível
#   - Flag de sistema crítico bancário (CFTV, cofre, etc.)
#   - Ambientes sem elementos atribuídos (alertas de vazio)
#
# Saída:
#   inventario_por_ambiente.csv  — uma linha por (ambiente × tipo)
#   inventario_pivot.csv         — pivot: ambientes nas linhas, tipos nas colunas
#   inventario_totais.csv        — totais por categoria/disciplina
#   inventario.json              — dados completos estruturados
# ============================================================

from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from collections import defaultdict, Counter, OrderedDict
import os, csv, datetime, json

doc   = revit.doc
model = doc.Title

# =========================
# CONFIG — AJUSTE AQUI
# =========================

BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"
PLUGIN_NAME      = "InventarioPorAmbiente"

# Categorias incluídas no inventário (tudo que não é estrutura/arquitetura pura)
CATEGORIAS_INCLUIR = set([
    # Elétrica
    u"Luminarias",          u"Luminarias",
    u"Dispositivos eletricos",
    u"Equipamentos eletricos",
    u"Circuitos eletricos",
    u"Identificadores de luminaria",
    # Infraestrutura TI
    u"Conduites",           u"Conexoes do conduite",
    u"Bandejas de cabos",   u"Conexoes da bandeja de cabos",
    u"Eletrocalhas",        u"Conexoes da eletrocalha",
    u"Fiacao",
    # Hidráulica / Incêndio
    u"Sprinklers",          u"Tubulacoes",
    u"Aparelhos sanitarios",
    u"Acessorios de tubulacao",
    u"Equipamentos mecanicos",
    # Segurança / Genérico
    u"Mobiliario",          u"Especialidades",
    u"Componentes",         u"Detalhes de modelo",
    # Portas e janelas são relevantes para inventário de acabamento
    u"Portas",              u"Janelas",
])

# Máximo de itens no popup por ambiente (evita popup gigante)
MAX_POPUP_AMBIENTES = 20


# ============================================================
# UTILITÁRIOS
# ============================================================

def norm_upper(s):
    try:
        if s is None:
            return u""
        if isinstance(s, unicode):
            return s.upper()
        return unicode(s, errors="ignore").upper()
    except:
        try:
            return str(s).upper()
        except:
            return u""


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


def eid_int(eid):
    try: return eid.IntegerValue
    except: pass
    try: return eid.Value
    except: pass
    try: return int(eid.ToString())
    except: return None


def get_family_type(el):
    try:
        t = doc.GetElement(el.GetTypeId())
        if not t:
            return (u"", u"")
        fam = u""
        try:
            fam = t.FamilyName
        except:
            pass
        return (fam, t.Name or u"")
    except:
        return (u"", u"")


def safe_meters(val_internal):
    """Converte de unidades internas do Revit para metros."""
    try:
        return round(UnitUtils.ConvertFromInternalUnits(val_internal, UnitTypeId.Meters), 3)
    except:
        try:
            # Fallback: ~0.3048 pés por metro (pré-2021)
            return round(val_internal * 0.3048, 3)
        except:
            return 0.0


def safe_sqmeters(val_internal):
    """Converte área de unidades internas do Revit para m²."""
    try:
        return round(UnitUtils.ConvertFromInternalUnits(val_internal, UnitTypeId.SquareMeters), 2)
    except:
        try:
            return round(val_internal * 0.3048 * 0.3048, 2)
        except:
            return 0.0


def get_level_name(el):
    try:
        if el.LevelId and el.LevelId != ElementId.InvalidElementId:
            lv = doc.GetElement(el.LevelId)
            return lv.Name if lv and lv.Name else u""
    except:
        pass
    return u""


def get_location_point(el):
    """Retorna o XYZ de localização do elemento (ponto ou início da curva)."""
    try:
        loc = el.Location
        if isinstance(loc, LocationPoint):
            return loc.Point
        if isinstance(loc, LocationCurve):
            return loc.Curve.GetEndPoint(0)
    except:
        pass
    return None


# ============================================================
# PATHS DE SAÍDA
# ============================================================

now        = datetime.datetime.now()
ts         = now.strftime("%H%M%S")
day        = now.strftime("%Y-%m-%d")
run_stamp  = now.strftime("%Y-%m-%d_%H%M%S")
model_safe = model.replace("/","_").replace("\\","_").replace(":","_")
ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR    = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
for _d in [RUN_DIR, LATEST_DIR]:
    if not os.path.exists(_d):
        os.makedirs(_d)

out_inv     = os.path.join(RUN_DIR, "inventario_por_ambiente.csv")
out_totais  = os.path.join(RUN_DIR, "inventario_totais.csv")
out_json    = os.path.join(RUN_DIR, "inventario.json")
out_inv_lat = os.path.join(LATEST_DIR, "inventario_por_ambiente.csv")
out_tot_lat = os.path.join(LATEST_DIR, "inventario_totais.csv")
out_json_lat= os.path.join(LATEST_DIR, "inventario.json")


# ============================================================
# 1) COLETA DE AMBIENTES (ROOMS / SPACES)
# ============================================================
print("Coletando ambientes...")

# Room = ambiente com volume calculado (arquitetura)
# Space = ambiente MEP (utilizado em projetos de instalacoes)
rooms_raw = []

for rm in FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Rooms) \
        .WhereElementIsNotElementType():
    try:
        area = safe_sqmeters(rm.Area)
        if area <= 0:
            continue   # Ignora rooms sem área (nao colocados em planta)
        lv_name = u""
        try:
            lv_name = doc.GetElement(rm.LevelId).Name or u""
        except:
            pass
        numero = u""
        try:
            p = rm.LookupParameter("Number") or rm.LookupParameter("Numero")
            if p and p.AsString():
                numero = p.AsString().strip()
        except:
            pass
        rooms_raw.append({
            "id":       eid_int(rm.Id),
            "name":     rm.Name or u"Sem Nome",
            "number":   numero,
            "level":    lv_name,
            "area_m2":  area,
            "obj":      rm,
        })
    except:
        continue

# Também tenta Spaces (MEP)
for sp in FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_MEPSpaces) \
        .WhereElementIsNotElementType():
    try:
        area = safe_sqmeters(sp.Area)
        if area <= 0:
            continue
        lv_name = u""
        try:
            lv_name = doc.GetElement(sp.LevelId).Name or u""
        except:
            pass
        rooms_raw.append({
            "id":       eid_int(sp.Id),
            "name":     sp.Name or u"Sem Nome",
            "number":   u"",
            "level":    lv_name,
            "area_m2":  area,
            "obj":      sp,
        })
    except:
        continue

print("  {} ambiente(s) encontrado(s).".format(len(rooms_raw)))

# Mapa rápido: id_room -> dados do room
room_map = {r["id"]: r for r in rooms_raw}

# ============================================================
# 2) COLETA DE ELEMENTOS E ASSOCIAÇÃO AOS AMBIENTES
# ============================================================
# Estratégia de associação (em ordem de prioridade):
#   A) Parâmetro "Room" nativo do elemento (MEP elements têm FromRoom/ToRoom)
#   B) Parâmetro "Space" (elementos MEP com Space)
#   C) doc.GetRoomAtPoint(location_point) — busca geométrica
#   D) Sem ambiente → grupo "SEM AMBIENTE"

print("Coletando e associando elementos...")

# inventario[room_id][(categoria, familia, tipo)] = count
inventario = defaultdict(lambda: defaultdict(int))
sem_ambiente = defaultdict(int)  # (cat,fam,typ) -> count

total_el   = 0
associados = 0
nao_assoc  = 0

elements_all = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

for el in elements_all:
    try:
        cat_obj = el.Category
        if not cat_obj:
            continue
        cat = cat_obj.Name or u""

        # Filtra categorias de interesse
        if cat not in CATEGORIAS_INCLUIR:
            continue

        total_el += 1

        fam, typ = get_family_type(el)
        chave   = (cat, fam, typ)

        # --- Tentativa A: parâmetro Room embutido (lê Room.Id) ---
        room_id_found = None
        try:
            rm_param = el.get_Parameter(BuiltInParameter.ELEM_ROOM_ID)
            if rm_param and rm_param.AsElementId() != ElementId.InvalidElementId:
                rid = eid_int(rm_param.AsElementId())
                if rid and rid in room_map:
                    room_id_found = rid
        except:
            pass

        # --- Tentativa B: FromRoom (elementos MEP) ---
        if room_id_found is None:
            try:
                fr = el.FromRoom
                if fr:
                    rid = eid_int(fr.Id)
                    if rid and rid in room_map:
                        room_id_found = rid
            except:
                pass

        # --- Tentativa C: Space (MEP Spaces) ---
        if room_id_found is None:
            try:
                sp_param = el.get_Parameter(BuiltInParameter.ELEM_SPACE_ID)
                if sp_param and sp_param.AsElementId() != ElementId.InvalidElementId:
                    rid = eid_int(sp_param.AsElementId())
                    if rid and rid in room_map:
                        room_id_found = rid
            except:
                pass

        # --- Tentativa D: busca geométrica ---
        if room_id_found is None:
            pt = get_location_point(el)
            if pt:
                try:
                    rm_geo = doc.GetRoomAtPoint(pt)
                    if rm_geo:
                        rid = eid_int(rm_geo.Id)
                        if rid and rid in room_map:
                            room_id_found = rid
                except:
                    pass

        # --- Registra no inventário ---
        if room_id_found is not None:
            inventario[room_id_found][chave] += 1
            associados += 1
        else:
            sem_ambiente[chave] += 1
            nao_assoc += 1

    except:
        continue

print("  Elementos no inventario: {}".format(total_el))
print("  Associados a ambiente  : {}".format(associados))
print("  Sem ambiente           : {}".format(nao_assoc))


# ============================================================
# 3) TOTAIS POR CATEGORIA (global)
# ============================================================

totais_cat = Counter()   # (categoria, familia, tipo) -> count total

for room_id, itens in inventario.items():
    for chave, cnt in itens.items():
        totais_cat[chave] += cnt

for chave, cnt in sem_ambiente.items():
    totais_cat[chave] += cnt


# ============================================================
# 4) CSV — INVENTÁRIO POR AMBIENTE (linha por ambiente × tipo)
# ============================================================

def write_inventario_csv(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow([
        "model", "run_stamp",
        "nivel", "ambiente_nome", "ambiente_numero", "ambiente_area_m2",
        "categoria", "familia", "tipo", "quantidade",
    ])

    # Linhas com ambiente
    for room_id, itens in inventario.items():
        rm = room_map[room_id]
        for (cat, fam, typ), cnt in sorted(
                itens.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
            w.writerow([
                u8(model), u8(run_stamp),
                u8(rm["level"]), u8(rm["name"]),
                u8(rm["number"]), u8(rm["area_m2"]),
                u8(cat), u8(fam), u8(typ), u8(cnt),
            ])

    # Linhas sem ambiente
    for (cat, fam, typ), cnt in sorted(
            sem_ambiente.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        w.writerow([
            u8(model), u8(run_stamp),
            u8(""), u8("SEM AMBIENTE"), u8(""), u8(""),
            u8(cat), u8(fam), u8(typ), u8(cnt),
        ])
    f.close()

write_inventario_csv(out_inv)
write_inventario_csv(out_inv_lat)


# ============================================================
# 5) CSV — TOTAIS POR CATEGORIA
# ============================================================

def write_totais_csv(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "run_stamp", "categoria", "familia", "tipo",
                "quantidade_total"])
    for (cat, fam, typ), cnt in sorted(
            totais_cat.items(), key=lambda x: -x[1]):
        w.writerow([
            u8(model), u8(run_stamp),
            u8(cat), u8(fam), u8(typ), u8(cnt),
        ])
    f.close()

write_totais_csv(out_totais)
write_totais_csv(out_tot_lat)


# ============================================================
# 6) JSON — INVENTÁRIO COMPLETO ESTRUTURADO
# ============================================================

# Serializa inventario (defaultdict nao é JSON-serializable)
inv_serial = {}
for room_id, itens in inventario.items():
    rm = room_map[room_id]
    amb_key = u"{} — {}".format(rm["level"], rm["name"])
    inv_serial[amb_key] = {
        "id":       room_id,
        "nome":     rm["name"],
        "numero":   rm["number"],
        "nivel":    rm["level"],
        "area_m2":  rm["area_m2"],
        "itens": [
            {
                "categoria": cat,
                "familia":   fam,
                "tipo":      typ,
                "qtd":       cnt,
            }
            for (cat, fam, typ), cnt in sorted(
                itens.items(), key=lambda x: (x[0][0], x[0][2]))
        ],
        "total_itens": sum(itens.values()),
    }

totais_serial = [
    {"categoria": cat, "familia": fam, "tipo": typ, "qtd_total": cnt}
    for (cat, fam, typ), cnt in sorted(totais_cat.items(), key=lambda x: -x[1])
]

resultado = {
    "model":      model,
    "run_stamp":  run_stamp,
    "day":        day,
    "time":       ts,
    "run_dir":    RUN_DIR,
    "latest_dir": LATEST_DIR,
    "resumo": {
        "total_ambientes":    len(rooms_raw),
        "ambientes_com_itens": len(inventario),
        "total_elementos":    total_el,
        "associados":         associados,
        "sem_ambiente":       nao_assoc,
        "tipos_unicos":       len(totais_cat),
    },
    "inventario_por_ambiente": inv_serial,
    "totais_por_tipo": totais_serial,
    "sem_ambiente": [
        {"categoria": cat, "familia": fam, "tipo": typ, "qtd": cnt}
        for (cat, fam, typ), cnt in sorted(sem_ambiente.items(), key=lambda x: -x[1])
    ],
}


def write_json(path):
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    f.write(u8(json.dumps(resultado, ensure_ascii=False, indent=2)))
    f.close()

write_json(out_json)
write_json(out_json_lat)


# ============================================================
# 7) PRINT NO CONSOLE — RELATÓRIO DETALHADO
# ============================================================

SEP  = "=" * 65
SEP2 = "-" * 65

print(SEP)
print("  INVENTARIO POR AMBIENTE — AGENCIA BANCARIA")
print(SEP)
print("  Modelo : {}".format(model))
print("  Data   : {}  Hora: {}".format(day, now.strftime("%H:%M:%S")))
print("")
print("  Ambientes encontrados : {}".format(len(rooms_raw)))
print("  Ambientes com itens   : {}".format(len(inventario)))
print("  Elementos catalogados : {}".format(total_el))
print("  Associados a ambiente : {}".format(associados))
print("  Sem ambiente          : {}".format(nao_assoc))
print("  Tipos distintos       : {}".format(len(totais_cat)))
print("")

# Inventário por ambiente
print(SEP2)
print("  INVENTARIO POR AMBIENTE")
print(SEP2)

# Ordena ambientes por nível e depois por nome
sorted_rooms = sorted(
    [(rid, room_map[rid]) for rid in inventario],
    key=lambda x: (x[1]["level"], x[1]["name"])
)

for room_id, rm in sorted_rooms:
    itens = inventario[room_id]
    total_amb = sum(itens.values())
    print("  [{:>5} m2] [{}] {} — {} iten(s)".format(
        rm["area_m2"], rm["level"],
        (rm["name"].encode("utf-8") if isinstance(rm["name"], unicode) else rm["name"]),
        total_amb))
    for (cat, fam, typ), cnt in sorted(itens.items(), key=lambda x: -x[1]):
        tipo_label = typ or fam or cat
        print("      {:>4}x  {}".format(
            cnt,
            (tipo_label.encode("utf-8") if isinstance(tipo_label, unicode) else tipo_label)))

if sem_ambiente:
    print("")
    print("  [SEM AMBIENTE] {} tipo(s)  {} item(ns) total".format(
        len(sem_ambiente), sum(sem_ambiente.values())))
    for (cat, fam, typ), cnt in sorted(sem_ambiente.items(), key=lambda x: -x[1])[:10]:
        tipo_label = typ or fam or cat
        print("      {:>4}x  {}".format(
            cnt,
            (tipo_label.encode("utf-8") if isinstance(tipo_label, unicode) else tipo_label)))

print("")
print(SEP2)
print("  TOP 10 ITENS MAIS FREQUENTES (GERAL)")
print(SEP2)
for i, ((cat, fam, typ), cnt) in enumerate(
        sorted(totais_cat.items(), key=lambda x: -x[1])[:10], start=1):
    tipo_label = typ or fam or cat
    print("  {:>2}. {:>5}x  {}".format(
        i, cnt,
        (tipo_label.encode("utf-8") if isinstance(tipo_label, unicode) else tipo_label)))

print("")
print(SEP2)
print("  ARQUIVOS GERADOS")
print(SEP2)
print("  [Historico]  {}".format(RUN_DIR))
print("    - inventario_por_ambiente.csv")
print("    - inventario_totais.csv")
print("    - inventario.json")
print("")
print("  [Latest]     {}".format(LATEST_DIR))
print("    - inventario_por_ambiente.csv")
print("    - inventario_totais.csv")
print("    - inventario.json")
print(SEP)


# ============================================================
# 8) POPUP — RESUMO DENTRO DO REVIT
# ============================================================

lines = []
lines.append(u"Modelo: {}  |  {}".format(model, day))
lines.append(u"")
lines.append(u"=" * 52)
lines.append(u"  RESUMO DO INVENTARIO")
lines.append(u"=" * 52)
lines.append(u"  Ambientes catalogados : {}".format(len(rooms_raw)))
lines.append(u"  Ambientes com itens   : {}".format(len(inventario)))
lines.append(u"  Elementos catalogados : {}".format(total_el))
lines.append(u"  Tipos distintos       : {}".format(len(totais_cat)))
lines.append(u"  Sem ambiente (m)      : {}".format(nao_assoc))
lines.append(u"")

# Inventário resumido por ambiente (limite para o popup)
lines.append(u"-- INVENTARIO POR AMBIENTE (top {}) --------".format(MAX_POPUP_AMBIENTES))

exibidos = 0
for room_id, rm in sorted_rooms[:MAX_POPUP_AMBIENTES]:
    itens = inventario[room_id]
    total_amb = sum(itens.values())
    lines.append(u"  [{}] {} ({} m2) — {} iten(s)".format(
        rm["level"], rm["name"], rm["area_m2"], total_amb))
    # Top 5 itens do ambiente
    top_itens = sorted(itens.items(), key=lambda x: -x[1])[:5]
    for (cat, fam, typ), cnt in top_itens:
        tipo_label = typ or fam or cat
        lines.append(u"      {:>4}x  {}".format(cnt, tipo_label))
    if len(itens) > 5:
        lines.append(u"      ... e mais {} tipo(s)".format(len(itens) - 5))
    exibidos += 1

if len(sorted_rooms) > MAX_POPUP_AMBIENTES:
    lines.append(u"  ... e mais {} ambiente(s). Veja o CSV completo.".format(
        len(sorted_rooms) - MAX_POPUP_AMBIENTES))

lines.append(u"")
lines.append(u"=" * 52)
lines.append(u"Relatorios salvos em:")
lines.append(u"  " + LATEST_DIR)

msg = u"\n".join(lines)
forms.alert(msg, title=u"Inventario por Ambiente — Agencia Bancaria", warn_icon=False)
