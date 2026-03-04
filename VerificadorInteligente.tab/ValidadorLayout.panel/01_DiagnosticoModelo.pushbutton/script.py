# -*- coding: utf-8 -*-
# ============================================================
# VALIDADOR DE LAYOUT -- AGENCIA BANCARIA
# ============================================================
# Executa verificacoes automatizadas de layout, estilo e
# conformidade para projetos de agencia bancaria no Revit.
#
# Validacoes realizadas:
#   1.  Informacoes do Projeto
#   2.  Inventario de Ambientes (Rooms / Spaces)
#   3.  Checklist de Ambientes Obrigatorios
#   4.  Acessibilidade -- NBR 9050
#         * Largura de portas (>= 0,80 m)
#         * Areas acessiveis identificadas no modelo
#   5.  Areas Minimas por Ambiente
#   6.  Pe-direito (altura livre) por Ambiente
#   7.  Seguranca
#         * Cabine/cilindro de seguranca
#         * Cofre / sala-forte
#         * Saidas de emergencia
#   8.  Nomenclatura Padronizada
#   9.  Alertas do Modelo (Warnings Revit)
#  10.  Pontuacao / Score Geral
#
# Arquivos gerados:
#   validacao_agencia.json  -- snapshot completo das validacoes
#   relatorio_validacao.csv -- resumo das pendencias
#
# Compativel com Revit 2022-2026  |  IronPython 2 / CPython 3
# ============================================================

from pyrevit import revit, script

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    Level,
    Wall,
    SpatialElement,
)
from Autodesk.Revit.DB.Architecture import Room

try:
    from Autodesk.Revit.DB import UnitUtils, UnitTypeId
    _HAS_UNIT = True
except ImportError:
    _HAS_UNIT = False

from collections import Counter, defaultdict
import os, csv, datetime, json, sys

# ============================================================
# CONFIGURACOES -- AJUSTE PARA O PROJETO
# ============================================================
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "RevitScan"
PLUGIN_NAME      = "ValidadorAgencia"

# ---------- LIMITES NBR 9050 / BOAS PRATICAS BANCARIAS ------
PORTA_ACESSIVEL_MIN_M  = 0.80
PORTA_PRINCIPAL_MIN_M  = 0.90
PE_DIREITO_MIN_M       = 2.50
PE_DIREITO_HALL_MIN_M  = 3.00
AREA_MIN_SALA_M2       = 9.0
AREA_MIN_ESPERA_M2     = 15.0
AREA_MIN_CAIXA_M2      = 12.0
AREA_MIN_COFRE_M2      = 6.0
AREA_MIN_GERENTE_M2    = 8.0
AREA_MIN_BANHEIRO_M2   = 3.50

# ---------- PALAVRAS-CHAVE PARA IDENTIFICACAO DE AMBIENTES --
KW_ESPERA    = [u"ESPERA", u"LOBBY", u"HALL", u"RECEPCAO", u"AGUARDO"]
KW_CAIXA     = [u"CAIXA", u"TELLER", u"ATENDIMENTO", u"OPERACAO"]
KW_GERENTE   = [u"GERENTE", u"GERENCIA", u"REUNIAO", u"DIRETORIA"]
KW_COFRE     = [u"COFRE", u"SALA FORTE", u"SALA-FORTE", u"TESOURARIA", u"VAULT"]
KW_SEGURANCA = [u"SEGURANCA", u"CILINDRO", u"CABINE", u"GUARITA", u"CONTROLE"]
KW_ATM       = [u"ATM", u"CAIXA ELETRONICO", u"AUTOATENDIMENTO", u"SELF SERVICE"]
KW_BANHEIRO  = [u"BANHEIRO", u"SANITARIO", u"WC", u"LAVABO", u"TOILET"]
KW_ACESSIVEL = [u"ACESSIVEL", u"ACESSIBILIDADE", u"PNE", u"DEFICIENTE"]
KW_SAIDA     = [u"SAIDA", u"EMERGENCIA", u"ROTA FUGA", u"EVACUACAO"]
KW_TI        = [u"TI", u"SERVIDOR", u"TELECOMUNICACAO", u"RACK", u"DATACENTER"]

AMBIENTES_OBRIGATORIOS = [
    (u"Hall / Espera",         KW_ESPERA),
    (u"Area de Caixas",        KW_CAIXA),
    (u"Sala de Gerente",       KW_GERENTE),
    (u"Cofre / Sala-Forte",    KW_COFRE),
    (u"Seguranca / Cabine",    KW_SEGURANCA),
    (u"ATM / Autoatendimento", KW_ATM),
    (u"Banheiro",              KW_BANHEIRO),
    (u"TI / Servidores",       KW_TI),
]
# ============================================================

doc = revit.doc
app = __revit__.Application

now       = datetime.datetime.now()
run_stamp = now.strftime("%Y-%m-%d_%H%M%S")
model     = doc.Title
model_safe= model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
RUN_DIR    = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
for _d in [RUN_DIR, LATEST_DIR]:
    if not os.path.exists(_d):
        os.makedirs(_d)


# ============================================================
# UTILITARIOS
# ============================================================
try:
    _unicode = unicode
except NameError:
    _unicode = str


def u8(x):
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
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def sqft_to_sqm(sqft):
    try:
        if _HAS_UNIT:
            return round(UnitUtils.ConvertFromInternalUnits(sqft, UnitTypeId.SquareMeters), 2)
        return round(sqft * 0.092903, 2)
    except Exception:
        return 0.0


def ft_to_m(ft):
    try:
        if _HAS_UNIT:
            return round(UnitUtils.ConvertFromInternalUnits(ft, UnitTypeId.Meters), 3)
        return round(ft * 0.3048, 3)
    except Exception:
        return 0.0


def match_kw(name, keywords):
    name_up = u8(name).upper()
    return any(kw in name_up for kw in keywords)


def status_icon(ok):
    return u"OK" if ok else u"ATENCAO"


def normalize_json(obj):
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
    if sys.version_info[0] >= 3:
        return open(path, "w", encoding="utf-8-sig", newline="")
    f = open(path, "wb")
    f.write(b"\xef\xbb\xbf")
    return f


def csv_row(*cols):
    if sys.version_info[0] >= 3:
        return [u8(c) for c in cols]
    return [u8(c).encode("utf-8") for c in cols]


def save_json(obj, *paths):
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


def rname(r):
    """Retorna o nome do Room de forma segura via parametro BuiltIn."""
    try:
        p = r.get_Parameter(BuiltInParameter.ROOM_NAME)
        if p:
            return u8(p.AsString())
    except Exception:
        pass
    try:
        return u8(r.Name)
    except Exception:
        return u""


def rnumber(r):
    """Retorna o numero do Room de forma segura via parametro BuiltIn."""
    try:
        p = r.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if p:
            return u8(p.AsString())
    except Exception:
        pass
    try:
        return u8(r.Number)
    except Exception:
        return u""


def rlevel(r):
    """Retorna o nome do Level do Room de forma segura."""
    try:
        lv = r.Level
        if lv:
            return u8(lv.Name)
    except Exception:
        pass
    return u"--"


# ============================================================
# INICIALIZA PAINEL HTML
# ============================================================
output = script.get_output()
output.set_height(960)
output.set_title(u"Validador Agencia -- " + u8(model))

output.print_md(u"# Validador de Layout -- Agencia Bancaria")
output.print_md(
    u"**Modelo:** `{}`  \n**Execucao:** `{}`  \n**Revit:** {} ({})".format(
        u8(model), run_stamp,
        u8(str(app.VersionNumber)), u8(str(app.VersionName))
    )
)

# Contadores de score
total_checks  = 0
passed_checks = 0
pendencias    = []


def register(ok, categoria, descricao, detalhe=u""):
    global total_checks, passed_checks
    total_checks  += 1
    if ok:
        passed_checks += 1
    else:
        pendencias.append({
            "categoria": u8(categoria),
            "descricao": u8(descricao),
            "detalhe":   u8(detalhe),
            "status":    u"ATENCAO",
        })
    return ok


# ============================================================
# COLETA PRINCIPAL
# ============================================================
all_rooms = [
    r for r in FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements()
    if isinstance(r, Room) and r.Area > 0
]

all_doors = list(
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Doors)
    .WhereElementIsNotElementType()
    .ToElements()
)

all_walls = list(FilteredElementCollector(doc).OfClass(Wall).ToElements())

levels = sorted(
    FilteredElementCollector(doc).OfClass(Level).ToElements(),
    key=lambda lv: lv.Elevation
)

model_warnings = []
try:
    model_warnings = list(doc.GetWarnings())
except Exception:
    pass


# ============================================================
# 1. INFORMACOES DO PROJETO
# ============================================================
output.print_md(u"---\n## 1. Informacoes do Projeto")

proj = doc.ProjectInformation
output.print_table([
    [u"Nome do Projeto", u8(proj.Name)],
    [u"Numero",          u8(proj.Number)],
    [u"Cliente",         u8(proj.ClientName)],
    [u"Endereco",        u8(proj.Address)],
    [u"Status",          u8(proj.Status)],
    [u"Autor",           u8(proj.Author)],
    [u"Data de Emissao", u8(proj.IssueDate)],
], columns=[u"Campo", u"Valor"])


# ============================================================
# 2. INVENTARIO DE AMBIENTES
# ============================================================
output.print_md(u"---\n## 2. Inventario de Ambientes (Rooms)")

if not all_rooms:
    output.print_md(u"> **ATENCAO:** Nenhum Room com area encontrado. "
                    u"Verifique se os ambientes estao fechados e com area calculada.")
else:
    inv_table = []
    for r in sorted(all_rooms, key=lambda x: -x.Area):
        area_m2 = sqft_to_sqm(r.Area)
        link    = output.linkify(r.Id)
        inv_table.append([rnumber(r), rname(r), u"{} m2".format(area_m2), rlevel(r), link])
    output.print_table(inv_table, columns=[u"Num.", u"Nome", u"Area", u"Nivel", u"Selecionar"])
    output.print_md(u"_Total: {} ambientes com area calculada_".format(len(all_rooms)))


# ============================================================
# 3. CHECKLIST DE AMBIENTES OBRIGATORIOS
# ============================================================
output.print_md(u"---\n## 3. Checklist de Ambientes Obrigatorios")

room_names  = [rname(r) for r in all_rooms]
check_table = []
for amb_label, kw_list in AMBIENTES_OBRIGATORIOS:
    encontrados = [n for n in room_names if match_kw(n, kw_list)]
    ok  = len(encontrados) > 0
    qtd = str(len(encontrados))
    st  = status_icon(ok)
    register(ok, u"Ambientes Obrigatorios", amb_label,
             u", ".join(encontrados[:3]) if encontrados else u"Nao encontrado")
    check_table.append([
        amb_label,
        qtd,
        u", ".join(encontrados[:3]) if encontrados else u"--",
        st,
    ])
output.print_table(check_table,
    columns=[u"Ambiente Esperado", u"Qtd", u"Encontrado(s)", u"Status"])


# ============================================================
# 4. ACESSIBILIDADE -- NBR 9050
# ============================================================
output.print_md(u"---\n## 4. Acessibilidade -- NBR 9050")

# 4a. Largura de portas
output.print_md(u"### 4a. Largura de Portas")

porta_ok_count  = 0
porta_nok_list  = []
porta_total     = 0

for door in all_doors:
    try:
        w_param = door.get_Parameter(BuiltInParameter.DOOR_WIDTH)
        if w_param is None:
            typ = doc.GetElement(door.GetTypeId())
            w_param = typ.get_Parameter(BuiltInParameter.DOOR_WIDTH) if typ else None
        if w_param is None:
            continue
        largura_m = ft_to_m(w_param.AsDouble())
        room_name = u""
        try:
            fr = door.FromRoom
            if fr:
                room_name = u8(fr.Name)
        except Exception:
            pass
        ok = largura_m >= PORTA_ACESSIVEL_MIN_M
        porta_total += 1
        if ok:
            porta_ok_count += 1
        else:
            porta_nok_list.append([
                u8(door.Name), room_name,
                u"{:.3f} m".format(largura_m),
                u">= {:.2f} m".format(PORTA_ACESSIVEL_MIN_M),
                status_icon(ok),
                output.linkify(door.Id),
            ])
            register(False, u"Acessibilidade",
                     u"Porta abaixo da largura minima",
                     u"{} ({:.3f} m)".format(u8(door.Name), largura_m))
    except Exception:
        pass

if porta_nok_list:
    output.print_table(porta_nok_list,
        columns=[u"Tipo de Porta", u"Ambiente", u"Largura", u"Minimo", u"Status", u"Selecionar"])
    output.print_md(u"> **{}** porta(s) abaixo de {:.2f} m  |  **{}** porta(s) em conformidade".format(
        len(porta_nok_list), PORTA_ACESSIVEL_MIN_M, porta_ok_count))
else:
    if porta_total:
        output.print_md(u"> Todas as **{}** portas verificadas atendem a largura minima de {:.2f} m.".format(
            porta_total, PORTA_ACESSIVEL_MIN_M))
    else:
        output.print_md(u"> _Nenhuma porta com parametro de largura encontrada._")

register(len(porta_nok_list) == 0, u"Acessibilidade",
         u"Todas as portas >= {:.2f} m".format(PORTA_ACESSIVEL_MIN_M),
         u"{}/{} em conformidade".format(porta_ok_count, porta_total))

# 4b. Banheiros acessiveis
output.print_md(u"### 4b. Banheiros Acessiveis")

banheiros  = [r for r in all_rooms if match_kw(rname(r), KW_BANHEIRO)]
ban_acess  = [r for r in banheiros if match_kw(rname(r), KW_ACESSIVEL)]
ban_table  = []
for r in banheiros:
    area_m2  = sqft_to_sqm(r.Area)
    eh_acess = match_kw(rname(r), KW_ACESSIVEL)
    ok_area  = area_m2 >= AREA_MIN_BANHEIRO_M2
    ban_table.append([
        rname(r), u"{} m2".format(area_m2),
        u"Sim" if eh_acess else u"Nao",
        status_icon(ok_area), output.linkify(r.Id),
    ])
    if not ok_area:
        register(False, u"Acessibilidade",
                 u"Banheiro abaixo de {:.2f} m2".format(AREA_MIN_BANHEIRO_M2),
                 u"{} = {:.2f} m2".format(rname(r), area_m2))

if ban_table:
    output.print_table(ban_table,
        columns=[u"Nome", u"Area", u"Acessivel", u"Area OK", u"Selecionar"])
    tem_acessivel = len(ban_acess) > 0
    register(tem_acessivel, u"Acessibilidade",
             u"Banheiro acessivel (NBR 9050) presente",
             u"{} banheiro(s) com 'acessivel' no nome".format(len(ban_acess)))
    output.print_md(u"> {} banheiro(s) total | {} acessivel(is) identificado(s)".format(
        len(banheiros), len(ban_acess)))
else:
    register(False, u"Acessibilidade", u"Nenhum banheiro encontrado no modelo", u"")
    output.print_md(u"> _Nenhum banheiro encontrado._")


# ============================================================
# 5. AREAS MINIMAS POR TIPO DE AMBIENTE
# ============================================================
output.print_md(u"---\n## 5. Areas Minimas por Ambiente")

AREA_REGRAS = [
    (u"Hall / Espera",      KW_ESPERA,  AREA_MIN_ESPERA_M2),
    (u"Area de Caixas",     KW_CAIXA,   AREA_MIN_CAIXA_M2),
    (u"Sala de Gerente",    KW_GERENTE, AREA_MIN_GERENTE_M2),
    (u"Cofre / Sala-Forte", KW_COFRE,   AREA_MIN_COFRE_M2),
]

area_table = []
for label, kws, min_m2 in AREA_REGRAS:
    ambientes_match = [r for r in all_rooms if match_kw(rname(r), kws)]
    if not ambientes_match:
        area_table.append([label, u"--", u"Nao encontrado", u">= {} m2".format(min_m2), u"PENDENTE"])
        register(False, u"Areas Minimas", label + u" nao encontrado no modelo", u"")
        continue
    for r in ambientes_match:
        area_m2 = sqft_to_sqm(r.Area)
        ok      = area_m2 >= min_m2
        area_table.append([label, rname(r), u"{} m2".format(area_m2), u">= {} m2".format(min_m2), status_icon(ok)])
        if not ok:
            register(False, u"Areas Minimas", u"{} abaixo do minimo".format(label),
                     u"{} = {:.2f} m2 (min: {:.2f} m2)".format(rname(r), area_m2, min_m2))
        else:
            register(True, u"Areas Minimas", u"{} em conformidade".format(label))

output.print_table(area_table,
    columns=[u"Tipo Esperado", u"Room", u"Area Real", u"Minimo", u"Status"])


# ============================================================
# 6. PE-DIREITO POR AMBIENTE
# ============================================================
output.print_md(u"---\n## 6. Pe-Direito por Ambiente")

pd_table = []
for r in all_rooms:
    try:
        h_param = r.get_Parameter(BuiltInParameter.ROOM_HEIGHT)
        if h_param is None:
            continue
        altura_m = ft_to_m(h_param.AsDouble())
        if altura_m <= 0:
            continue
        if match_kw(rname(r), KW_ESPERA + KW_CAIXA):
            minimo = PE_DIREITO_HALL_MIN_M
        else:
            minimo = PE_DIREITO_MIN_M
        ok = altura_m >= minimo
        pd_table.append([
            rname(r), u"{:.3f} m".format(altura_m),
            u">= {:.2f} m".format(minimo), status_icon(ok), output.linkify(r.Id),
        ])
        if not ok:
            register(False, u"Pe-Direito", u"Pe-direito abaixo do minimo",
                     u"{}: {:.3f} m (min {:.2f} m)".format(rname(r), altura_m, minimo))
    except Exception:
        pass

if pd_table:
    nok_pd = [row for row in pd_table if row[3] == u"ATENCAO"]
    ok_pd  = [row for row in pd_table if row[3] == u"OK"]
    if nok_pd:
        output.print_table(nok_pd,
            columns=[u"Ambiente", u"Altura", u"Minimo", u"Status", u"Selecionar"])
    output.print_md(u"> {}/{} ambientes com pe-direito em conformidade".format(
        len(ok_pd), len(pd_table)))
    if not nok_pd:
        output.print_md(u"> Todos os ambientes atendem o pe-direito minimo.")
else:
    output.print_md(u"> _Nenhum ambiente com parametro de altura encontrado._")


# ============================================================
# 7. SEGURANCA -- ELEMENTOS CRITICOS
# ============================================================
output.print_md(u"---\n## 7. Seguranca -- Elementos Criticos")

seg_table = []

def check_security_zone(label, keywords, required=True):
    matches = [r for r in all_rooms if match_kw(rname(r), keywords)]
    ok    = len(matches) > 0
    qtd   = str(len(matches))
    nomes = u", ".join(rname(r) for r in matches[:3])
    st    = status_icon(ok) if required else (u"OK" if ok else u"INFO")
    if required:
        register(ok, u"Seguranca", label, nomes if ok else u"Nao encontrado no modelo")
    seg_table.append([label, qtd, nomes if ok else u"--", st])

check_security_zone(u"Cabine / Cilindro de Seguranca", KW_SEGURANCA, required=True)
check_security_zone(u"Cofre / Sala-Forte",             KW_COFRE,     required=True)
check_security_zone(u"Saida de Emergencia / Rota Fuga",KW_SAIDA,     required=True)
check_security_zone(u"Area de Autoatendimento (ATM)",   KW_ATM,       required=True)
check_security_zone(u"Sala de TI / Servidores",         KW_TI,        required=False)

output.print_table(seg_table,
    columns=[u"Elemento de Seguranca", u"Qtd", u"Ambiente(s)", u"Status"])

# Paredes especiais (blindagem / corta-fogo)
output.print_md(u"### 7a. Paredes Especiais (Blindagem / Seguranca)")

paredes_seg = []
for w in all_walls:
    try:
        nome_tipo = u""
        typ = doc.GetElement(w.GetTypeId())
        if typ:
            nome_tipo = u8(typ.Name)
        if any(kw in nome_tipo.upper() for kw in
               [u"BLIND", u"SEGUR", u"CORTA-FOGO", u"CORTA FOGO", u"FIRE", u"REFOR"]):
            paredes_seg.append([nome_tipo, output.linkify(w.Id)])
    except Exception:
        pass

if paredes_seg:
    output.print_table(paredes_seg, columns=[u"Tipo de Parede", u"Selecionar"])
    output.print_md(u"_Total: {} parede(s) especial(is) identificada(s)_".format(len(paredes_seg)))
else:
    output.print_md(
        u"> _Nenhuma parede com nomenclatura de blindagem/corta-fogo encontrada. "
        u"Verifique os tipos de parede do projeto._"
    )


# ============================================================
# 8. NOMENCLATURA PADRONIZADA
# ============================================================
output.print_md(u"---\n## 8. Nomenclatura dos Ambientes")

sem_nome   = [r for r in all_rooms if not rname(r).strip()]
sem_numero = [r for r in all_rooms if not rnumber(r).strip()]
duplicados = [name for name, cnt in Counter(rname(r) for r in all_rooms).items() if cnt > 1]

output.print_table([
    [u"Rooms sem nome",    str(len(sem_nome)),   status_icon(len(sem_nome) == 0)],
    [u"Rooms sem numero",  str(len(sem_numero)), status_icon(len(sem_numero) == 0)],
    [u"Nomes duplicados",  str(len(duplicados)), status_icon(len(duplicados) == 0)],
], columns=[u"Verificacao", u"Qtd", u"Status"])

register(len(sem_nome) == 0,   u"Nomenclatura", u"Rooms sem nome",   str(len(sem_nome)))
register(len(sem_numero) == 0, u"Nomenclatura", u"Rooms sem numero", str(len(sem_numero)))
register(len(duplicados) == 0, u"Nomenclatura", u"Nomes duplicados", u", ".join(duplicados[:5]))

if sem_nome:
    output.print_table([[output.linkify(r.Id), u8(r.Number)] for r in sem_nome[:10]],
        columns=[u"Selecionar", u"Numero Room"])
if duplicados:
    output.print_md(u"**Nomes duplicados:** " + u", ".join(u"`{}`".format(d) for d in duplicados))



# ============================================================
# 9. ALERTAS DO MODELO (Warnings Revit)
# ============================================================
output.print_md(u"---\n## 9. Alertas do Modelo (Warnings)")

if model_warnings:
    warn_counter = Counter()
    for w in model_warnings:
        try:
            desc = u8(str(w.GetDescriptionText()))[:100]
        except Exception:
            desc = u"(erro ao ler descricao)"
        warn_counter[desc] += 1

    total_warn = len(model_warnings)
    output.print_table([
        [u8(desc), str(cnt), u"{:.1f}%".format(cnt * 100.0 / total_warn)]
        for desc, cnt in warn_counter.most_common(15)
    ], columns=[u"Descricao", u"Ocorrencias", u"%"])
    output.print_md(u"_Total: {} alertas no modelo_".format(total_warn))
    register(total_warn == 0, u"Alertas", u"Modelo sem warnings", str(total_warn))
else:
    output.print_md(u"> Nenhum alerta encontrado no modelo.")
    register(True, u"Alertas", u"Modelo sem warnings", u"0")


# ============================================================
# 10. PONTUACAO / SCORE GERAL
# ============================================================
output.print_md(u"---\n## 10. Score de Conformidade")

score_pct = round(passed_checks * 100.0 / total_checks, 1) if total_checks else 0

if score_pct >= 90:
    nivel = u"EXCELENTE"
elif score_pct >= 75:
    nivel = u"BOM"
elif score_pct >= 50:
    nivel = u"REGULAR"
else:
    nivel = u"CRITICO"

output.print_table([
    [u"Verificacoes realizadas", str(total_checks)],
    [u"Aprovadas",               "{} ({:.1f}%)".format(passed_checks, score_pct)],
    [u"Pendencias",              str(total_checks - passed_checks)],
    [u"Nivel de Conformidade",   u"{} -- {}%".format(nivel, score_pct)],
], columns=[u"Indicador", u"Resultado"])

if pendencias:
    output.print_md(u"### Pendencias Detectadas")
    output.print_table([
        [p["categoria"], p["descricao"], p["detalhe"]]
        for p in pendencias
    ], columns=[u"Categoria", u"Pendencia", u"Detalhe"])


# ============================================================
# EXPORTACAO CSV
# ============================================================
for path in [
    os.path.join(RUN_DIR,    "relatorio_validacao.csv"),
    os.path.join(LATEST_DIR, "relatorio_validacao.csv"),
]:
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(csv_row("run_stamp", "model", "categoria", "descricao", "detalhe", "status"))
    for p in pendencias:
        w.writerow(csv_row(run_stamp, model,
                           p["categoria"], p["descricao"], p["detalhe"], p["status"]))
    f.close()


# ============================================================
# EXPORTACAO JSON
# ============================================================
resultado = {
    "run_stamp": u8(run_stamp),
    "model":     u8(model),
    "revit_version": {
        "number": u8(str(app.VersionNumber)),
        "name":   u8(str(app.VersionName)),
    },
    "project_info": {
        "name":    u8(proj.Name),
        "number":  u8(proj.Number),
        "client":  u8(proj.ClientName),
        "address": u8(proj.Address),
    },
    "score": {
        "total_checks": total_checks,
        "passed":       passed_checks,
        "failed":       total_checks - passed_checks,
        "score_pct":    score_pct,
        "nivel":        nivel,
    },
    "pendencias": normalize_json(pendencias),
    "inventario": [
        {
            "number":  rnumber(r),
            "name":    rname(r),
            "area_m2": sqft_to_sqm(r.Area),
            "level":   rlevel(r),
        }
        for r in all_rooms
    ],
}
save_json(
    resultado,
    os.path.join(RUN_DIR,    "validacao_agencia.json"),
    os.path.join(LATEST_DIR, "validacao_agencia.json"),
)


# ============================================================
# RODAPE
# ============================================================
output.print_md(u"---\n## Arquivos Gerados")
output.print_table([
    [u"Historico (esta execucao)", u8(RUN_DIR)],
    [u"Ultima execucao (latest)",  u8(LATEST_DIR)],
], columns=[u"Destino", u"Caminho"])
output.print_md(u"> **Arquivos:** `relatorio_validacao.csv` | `validacao_agencia.json`")


# ============================================================
# CONSOLE
# ============================================================
print("=" * 65)
print("  VALIDADOR DE LAYOUT -- AGENCIA BANCARIA")
print("=" * 65)
print("  Modelo     : {}".format(model))
print("  Revit      : {} ({})".format(app.VersionNumber, app.VersionName))
print("  Score      : {}% -- {}".format(score_pct, nivel))
print("-" * 65)
print("  Checks     : {}".format(total_checks))
print("  Aprovados  : {}".format(passed_checks))
print("  Pendencias : {}".format(total_checks - passed_checks))
print("  Rooms      : {}".format(len(all_rooms)))
print("  Portas     : {}".format(len(all_doors)))
print("  Alertas    : {}".format(len(model_warnings)))
print("-" * 65)
print("  [Historico] {}".format(RUN_DIR))
print("  [Latest]    {}".format(LATEST_DIR))
print("=" * 65)
print("PROCESSO REALIZADO COM SUCESSO")