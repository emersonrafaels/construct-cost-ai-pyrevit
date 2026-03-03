# -*- coding: utf-8 -*-

# ============================================================
# RESUMO DO MODELO — AGÊNCIA BANCÁRIA
# ============================================================
# Gera um perfil completo do modelo Revit aberto, com foco em
# indicadores relevantes para projetos de agências bancárias:
#
#   - Metadados do projeto (autor, organização, versão Revit)
#   - Estatísticas gerais (elementos, categorias, folhas, níveis)
#   - Detecção de sistemas críticos bancários (CFTV, SPDA, UPS, etc.)
#   - Análise de ambientes (cofre, atendimento, servidores, etc.)
#   - Índice de Maturidade BIM (preenchimento de parâmetros-chave)
#   - Índice de Cobertura de Documentação (folhas por disciplina)
#   - Alertas contextuais para agência bancária
#
# Saída: popup dentro do Revit + CSV + JSON completo
# ============================================================

from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from collections import Counter, defaultdict
import os, csv, datetime, json

doc   = revit.doc
app   = doc.Application

# =========================
# CONFIG — AJUSTE AQUI
# =========================

BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"
TOP_N            = 20    # Ranking de categorias exibidas

# Disciplinas detectadas nas folhas (pranchas)
SHEET_KEYWORDS = [
    ("ELETRICA",    ["EL", "ELE", "ELETR", "E-"]),
    ("INFRA",       ["INF", "TI", "DADOS", "REDE", "RACK", "CABEAMENTO", "TELECOM"]),
    ("SEGURANCA",   ["SEG", "CFTV", "ALARME", "ACESSO", "CAMERA", "CAMARA", "SPDA"]),
    ("HIDRAULICA",  ["HID", "AGUA", "SPKL", "SPRINK", "INCEND", "H-"]),
    ("ARQUITETURA", ["ARQ", "ARQ-", "A-", "PLANT", "LAYOUT", "ARQT"]),
    ("ESTRUTURA",   ["EST", "STRUCT", "FUND", "PILAR", "VIGA"]),
    ("CLIMATIZACAO",["AR COND", "HVAC", "CLIMATIZ", "AC-", "VRF"]),
]

# Sistemas críticos bancários — detectados por palavras-chave no nome de família/tipo
KW_BANCO = {
    "CFTV / Cameras":         ["CFTV","CAMERA","CAMARA","DOME","BULLET","PTZ"],
    "Controle de Acesso":     ["CATRACA","CANCELA","BIOMETRI","LEITOR","TORNIQUETE","ACCESS CONTROL"],
    "UPS / No-break":         ["UPS","NO-BREAK","NOBREAK","NO BREAK","GERADOR","GE_"],
    "SPDA / Aterramento":     ["SPDA","PARA-RAIO","PARA RAIO","ATERRAMENTO","MALHA TERRA","DPS"],
    "Sprinkler / Incendio":   ["SPRINK","SPKL","BOCAL","CHUVEIRO AUTOM","DETECTOR","FUMACA","SMOKE","ACIONADOR"],
    "Rack TI / Servidores":   ["RACK","SERVIDOR","SERVER","DATA CENTER","SWITCH","PATCH PANEL"],
    "Quadros Eletricos":      ["QDL","QGBT","QUADRO DISTR","QUADRO GERAL","PAINEL ELET","QD_"],
    "Alarme Anti-Intrusao":   ["ALARME","ALARM","ANTI-INTRUS","SENSOR MOVIM","PIR","INFRAVERM"],
    "Acessibilidade PCD":     ["RAMPA PCD","ACESSIB","PCD","DEFICIENTE","PISO TATIL","VAGA PCD"],
    "Cofre / Tesouraria":     ["COFRE","VAULT","TESOURARIA","SALA COFRE","STRONG ROOM"],
}

# Ambientes críticos bancários — detectados em Room/Space/categoria
KW_AMBIENTES = {
    "Cofre / Tesouraria":     ["COFRE","VAULT","TESOURARIA"],
    "Sala de Segurança":      ["VIGILANCIA","SEGURANCA","MONITORAMENTO","CENTRAL SEG"],
    "Sala de Servidores":     ["SERVIDOR","DATACENTER","DATA CENTER","RACK"],
    "Atendimento / Caixas":   ["ATENDIMENTO","CAIXA","AGENCIA","BANCARIO"],
    "Gerência / Diretoria":   ["GERENCIA","DIRETORIA","SALA REUNIAO"],
    "Sanitários PCD":         ["WC PCD","BANHEIRO PCD","SANITARIO PCD"],
}

# Parâmetros BIM cujo preenchimento é avaliado no índice de maturidade
BIM_PARAMS = [
    "Mark", "Comments", "Phase Created", "Workset",
    "OmniClass Number", "Assembly Code", "URL",
]

# =========================
# UTILITÁRIOS
# =========================

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
            return ""


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


def safe_param_str(el, param_name):
    """Lê o valor de um parâmetro do Revit como string, retorna '' se ausente."""
    try:
        p = el.LookupParameter(param_name)
        if p is None:
            return u""
        if p.StorageType == StorageType.String:
            v = p.AsString()
            return v.strip() if v else u""
        if p.StorageType == StorageType.Integer:
            return unicode(p.AsInteger())
        if p.StorageType == StorageType.Double:
            return unicode(round(p.AsDouble(), 4))
        if p.StorageType == StorageType.ElementId:
            eid = p.AsElementId()
            if eid and eid.IntegerValue != -1:
                return unicode(eid.IntegerValue)
        return u""
    except:
        return u""


def get_family_type_text(el):
    """Retorna 'FamilyName / TypeName' do elemento."""
    try:
        t = doc.GetElement(el.GetTypeId())
        if not t:
            return u""
        fam = u""
        try:
            fam = t.FamilyName
        except:
            pass
        return norm_upper(fam + u" " + (t.Name or u""))
    except:
        return u""


def classify_sheet(sh):
    text = norm_upper((sh.SheetNumber or "") + " " + (sh.Name or ""))
    hits = []
    for label, keys in SHEET_KEYWORDS:
        for k in keys:
            if k in text:
                hits.append(label)
                break
    if not hits:
        return "OUTROS"
    hits = sorted(list(set(hits)))
    return "+".join(hits)


def pct(part, total):
    if total == 0:
        return 0.0
    return round(part * 100.0 / total, 1)


def bar(value, max_val=100, width=20):
    """Barra de progresso ASCII."""
    filled = int(round(value * width / float(max_val))) if max_val else 0
    filled = max(0, min(width, filled))
    return u"[" + u"#" * filled + u"." * (width - filled) + u"]"


# =========================
# PATHS DE SAÍDA
# =========================

now        = datetime.datetime.now()
ts         = now.strftime("%H%M%S")
day        = now.strftime("%Y-%m-%d")
run_stamp  = now.strftime("%Y-%m-%d_%H%M%S")
model      = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")
ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
PLUGIN_NAME = "ResumoModelo"
RUN_DIR    = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
for _d in [RUN_DIR, LATEST_DIR]:
    if not os.path.exists(_d):
        os.makedirs(_d)

out_topcats        = os.path.join(RUN_DIR,    "top_categorias.csv")
out_sheets_csv     = os.path.join(RUN_DIR,    "folhas_resumo.csv")
out_banco_csv      = os.path.join(RUN_DIR,    "sistemas_bancarios.csv")
out_json           = os.path.join(RUN_DIR,    "model_profile.json")
out_topcats_latest = os.path.join(LATEST_DIR, "top_categorias.csv")
out_sheets_latest  = os.path.join(LATEST_DIR, "folhas_resumo.csv")
out_banco_latest   = os.path.join(LATEST_DIR, "sistemas_bancarios.csv")
out_json_latest    = os.path.join(LATEST_DIR, "model_profile.json")


# ============================================================
# 1) METADADOS DO PROJETO (ProjectInfo)
# ============================================================
# ProjectInfo é um singleton no Revit que armazena dados gerais do projeto.

pi = doc.ProjectInformation

def pi_str(param_name):
    try:
        p = pi.LookupParameter(param_name)
        if p and p.AsString():
            return p.AsString().strip()
        return u""
    except:
        return u""

meta = {
    "nome_projeto":    pi_str("Project Name") or pi_str("Nome do Projeto") or model,
    "numero_projeto":  pi_str("Project Number") or pi_str("Numero do Projeto") or u"",
    "cliente":         pi_str("Client Name") or pi_str("Cliente") or u"",
    "autor":           pi_str("Author") or pi_str("Autor") or u"",
    "organizacao":     pi_str("Organization Name") or pi_str("Organizacao") or u"",
    "endereco":        pi_str("Project Address") or pi_str("Endereco") or u"",
    "status":          pi_str("Project Status") or pi_str("Status") or u"",
    "emissao":         pi_str("Project Issue Date") or pi_str("Data de Emissao") or u"",
    "versao_revit":    u"{}".format(app.VersionNumber) if app else u"",
    "nome_build":      u"{}".format(app.VersionName) if app else u"",
}


# ============================================================
# 2) ELEMENTOS, CATEGORIAS E COLETA GERAL
# ============================================================

elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())
total_el = len(elements)

cats = []
for el in elements:
    try:
        if el.Category:
            cats.append(el.Category.Name)
    except:
        pass

counter       = Counter(cats)
cat_unique    = len(counter)

# Tipos (FamilySymbol) — variedade de tipos diferentes no projeto
types_all = list(FilteredElementCollector(doc).WhereElementIsElementType())
total_tipos = len(types_all)

# Famílias carregadas (FamilySymbol com FamilyName disponível)
family_names = set()
for t in types_all:
    try:
        fn = t.FamilyName
        if fn:
            family_names.add(fn)
    except:
        pass
total_familias = len(family_names)


# ============================================================
# 3) NÍVEIS (ANDARES)
# ============================================================
levels_col = FilteredElementCollector(doc).OfClass(Level).ToElements()
levels     = sorted(list(levels_col), key=lambda l: l.Elevation)
total_niveis = len(levels)

level_names = []
for lv in levels:
    try:
        elev_m = round(UnitUtils.ConvertFromInternalUnits(lv.Elevation, UnitTypeId.Meters), 2)
        level_names.append(u"{} ({} m)".format(lv.Name, elev_m))
    except:
        try:
            level_names.append(lv.Name or u"")
        except:
            pass


# ============================================================
# 4) ROOMS / SPACES — AMBIENTES MODELADOS
# ============================================================
rooms_col = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rooms) \
    .WhereElementIsNotElementType().ToElements()
total_rooms = len(list(rooms_col))

# Ambientes críticos bancários
ambientes_count = {k: 0 for k in KW_AMBIENTES}
area_total_m2   = 0.0

for rm in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType():
    try:
        rname = norm_upper(rm.Name or "")
        # Área do ambiente (em m²)
        try:
            area_m2 = UnitUtils.ConvertFromInternalUnits(rm.Area, UnitTypeId.SquareMeters)
            area_total_m2 += area_m2
        except:
            pass
        for amb_label, kws in KW_AMBIENTES.items():
            for kw in kws:
                if norm_upper(kw) in rname:
                    ambientes_count[amb_label] += 1
                    break
    except:
        continue


# ============================================================
# 5) FOLHAS (PRANCHAS) — CLASSIFICAÇÃO E REVISÃO
# ============================================================
sheets_col    = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
sheet_classes = Counter()
sheets_data   = []

for sh in sheets_col:
    try:
        disc = classify_sheet(sh)
        sheet_classes[disc] += 1

        # Revisão atual da folha
        rev_num  = u""
        rev_date = u""
        rev_desc = u""
        try:
            rev_num  = safe_param_str(sh, "Current Revision") or \
                       safe_param_str(sh, "Revisao Atual") or u""
            rev_date = safe_param_str(sh, "Current Revision Date") or \
                       safe_param_str(sh, "Data da Revisao") or u""
            rev_desc = safe_param_str(sh, "Current Revision Description") or u""
        except:
            pass

        # Aprovado para impressão
        approved = u""
        try:
            p = sh.LookupParameter("Approved By") or sh.LookupParameter("Aprovado Por")
            if p and p.AsString():
                approved = p.AsString().strip()
        except:
            pass

        sheets_data.append({
            "number":   sh.SheetNumber or u"",
            "name":     sh.Name or u"",
            "disc":     disc,
            "rev":      rev_num,
            "rev_date": rev_date,
            "rev_desc": rev_desc,
            "approved": approved,
        })
    except:
        continue

total_sheets = len(sheets_data)

# Contagem de revisões emitidas (folhas com revisão preenchida)
sheets_com_rev = sum(1 for s in sheets_data if s["rev"])


# ============================================================
# 6) SISTEMAS CRÍTICOS BANCÁRIOS
# ============================================================
banco_count    = {k: 0 for k in KW_BANCO}
banco_exemplos = {k: [] for k in KW_BANCO}

for el in elements:
    try:
        full = get_family_type_text(el)
        cat_name = el.Category.Name if el.Category else u""
        full_ext = full + u" " + norm_upper(cat_name)
        for sistema, kws in KW_BANCO.items():
            for kw in kws:
                if norm_upper(kw) in full_ext:
                    banco_count[sistema] += 1
                    ex = full.strip()
                    if ex and len(banco_exemplos[sistema]) < 3:
                        if ex not in banco_exemplos[sistema]:
                            banco_exemplos[sistema].append(ex)
                    break
    except:
        continue


# ============================================================
# 7) ÍNDICE DE MATURIDADE BIM
# ============================================================
# Avalia quão bem os parâmetros-chave estão preenchidos no modelo.
# Amostra até 500 elementos (para performance) e verifica cada parâmetro.

sample_size = min(500, total_el)
import random
try:
    sample = random.sample(elements, sample_size)
except:
    sample = elements[:sample_size]

param_fill = {p: 0 for p in BIM_PARAMS}   # Quantos elementos têm o param preenchido

for el in sample:
    for param_name in BIM_PARAMS:
        val = safe_param_str(el, param_name)
        if val:
            param_fill[param_name] += 1

# % de preenchimento por parâmetro (sobre o sample)
param_pct = {}
for param_name, count_filled in param_fill.items():
    param_pct[param_name] = pct(count_filled, sample_size)

# Índice geral de maturidade = média dos % de preenchimento
bim_index = round(sum(param_pct.values()) / float(len(param_pct)), 1) if param_pct else 0.0

def bim_nivel(score):
    if score >= 70:
        return u"ALTO"
    elif score >= 40:
        return u"MEDIO"
    else:
        return u"BAIXO"

bim_nivel_str = bim_nivel(bim_index)


# ============================================================
# 8) ÍNDICE DE COBERTURA DE DOCUMENTAÇÃO
# ============================================================
# Verifica quais disciplinas críticas de um banco têm folhas documentadas.

DISC_OBRIGATORIAS = ["ELETRICA", "INFRA", "SEGURANCA", "ARQUITETURA"]
DISC_RECOMENDADAS = ["HIDRAULICA", "ESTRUTURA", "CLIMATIZACAO"]

disc_cobertas_obrig = 0
disc_cobertas_recom = 0

def has_disc(disc):
    for k in sheet_classes:
        if disc in k:
            return True
    return False

for d in DISC_OBRIGATORIAS:
    if has_disc(d):
        disc_cobertas_obrig += 1
for d in DISC_RECOMENDADAS:
    if has_disc(d):
        disc_cobertas_recom += 1

cob_obrig_pct = pct(disc_cobertas_obrig, len(DISC_OBRIGATORIAS))
cob_recom_pct = pct(disc_cobertas_recom, len(DISC_RECOMENDADAS))


# ============================================================
# 9) ALERTAS CONTEXTUAIS BANCÁRIOS
# ============================================================

alertas   = []   # Criticos — indicam risco real
sugestoes = []   # Informativos — melhorias recomendadas

# Ambientes
if ambientes_count.get("Cofre / Tesouraria", 0) == 0:
    alertas.append(u"COFRE/Tesouraria nao detectado como Room/Space no modelo.")
if ambientes_count.get("Sala de Segurança", 0) == 0:
    alertas.append(u"Sala de Seguranca/Monitoramento nao modelada como ambiente.")
if ambientes_count.get("Sala de Servidores", 0) == 0:
    sugestoes.append(u"Sala de Servidores/Datacenter nao identificada como ambiente.")

# Sistemas criticos
if banco_count.get("CFTV / Cameras", 0) == 0:
    alertas.append(u"Nenhuma camera CFTV detectada — sistema de videovigilancia ausente.")
elif banco_count.get("CFTV / Cameras", 0) < 4:
    sugestoes.append(u"Apenas {} camera(s) CFTV — recomendado minimo de 4.".format(
        banco_count["CFTV / Cameras"]))

if banco_count.get("Controle de Acesso", 0) == 0:
    alertas.append(u"Controle de Acesso biometrico/catraca ausente — obrigatorio em banco.")

if banco_count.get("UPS / No-break", 0) == 0:
    alertas.append(u"UPS/No-break ausente — sistemas criticos sem redundancia de energia.")

if banco_count.get("SPDA / Aterramento", 0) == 0:
    alertas.append(u"SPDA/Aterramento ausente — exigido pela NBR 5419.")

if banco_count.get("Alarme Anti-Intrusao", 0) == 0:
    alertas.append(u"Sistema de alarme anti-intrusao ausente — obrigatorio para agencias.")

if banco_count.get("Rack TI / Servidores", 0) == 0:
    sugestoes.append(u"Rack de TI/Servidores nao detectado — verificar infraestrutura de TI.")

if banco_count.get("Sprinkler / Incendio", 0) == 0:
    alertas.append(u"Sistema de sprinkler/deteccao de incendio ausente.")

if banco_count.get("Acessibilidade PCD", 0) == 0:
    sugestoes.append(u"Elementos de acessibilidade PCD ausentes — NBR 9050.")

# Documentacao
if not has_disc("SEGURANCA"):
    alertas.append(u"Nenhuma folha de SEGURANCA detectada no projeto.")
if not has_disc("ELETRICA"):
    alertas.append(u"Nenhuma folha de ELETRICA detectada no projeto.")
if sheets_com_rev == 0 and total_sheets > 0:
    sugestoes.append(u"Nenhuma folha possui revisao registrada — verificar controle de emissao.")

# Maturidade BIM
if bim_index < 20:
    sugestoes.append(u"Maturidade BIM muito baixa ({} %) — parametros essenciais vazios.".format(bim_index))

# Níveis
if total_niveis == 0:
    alertas.append(u"Nenhum nivel (andar) definido no modelo.")
elif total_niveis == 1:
    sugestoes.append(u"Apenas 1 nivel definido — verificar se modelo contempla todos os pavimentos.")

# Rooms
if total_rooms == 0:
    sugestoes.append(u"Nenhum Room/Space modelado — analise de ambientes nao disponivel.")


# ============================================================
# 10) PRINT NO CONSOLE — RELATÓRIO DETALHADO
# ============================================================

SEP  = "=" * 65
SEP2 = "-" * 65

print(SEP)
print("  RESUMO DO MODELO — AGENCIA BANCARIA")
print(SEP)
print("  Arquivo  : {}".format(model))
print("  Projeto  : {}".format(meta["nome_projeto"] or "-"))
print("  Numero   : {}".format(meta["numero_projeto"] or "-"))
print("  Cliente  : {}".format(meta["cliente"] or "-"))
print("  Autor    : {}".format(meta["autor"] or "-"))
print("  Org      : {}".format(meta["organizacao"] or "-"))
print("  Endereco : {}".format(meta["endereco"] or "-"))
print("  Status   : {}".format(meta["status"] or "-"))
print("  Emissao  : {}".format(meta["emissao"] or "-"))
print("  Revit    : {} ({})".format(meta["nome_build"] or "-", meta["versao_revit"] or "-"))
print("  Executado: {}  {}".format(day, now.strftime("%H:%M:%S")))
print("")

print(SEP2)
print("  ESTATISTICAS GERAIS")
print(SEP2)
print("  Elementos       : {:>7}".format(total_el))
print("  Categorias      : {:>7}".format(cat_unique))
print("  Familias        : {:>7}".format(total_familias))
print("  Tipos           : {:>7}".format(total_tipos))
print("  Folhas          : {:>7}".format(total_sheets))
print("  Niveis (andares): {:>7}".format(total_niveis))
print("  Ambientes (Room): {:>7}".format(total_rooms))
if area_total_m2 > 0:
    print("  Area modelada   : {:>7} m2".format(round(area_total_m2, 1)))
print("")

print(SEP2)
print("  NIVEIS DO PROJETO")
print(SEP2)
for lname in level_names:
    print("  " + (lname.encode("utf-8") if isinstance(lname, unicode) else lname))
print("")

print(SEP2)
print("  TOP {} CATEGORIAS MAIS FREQUENTES".format(TOP_N))
print(SEP2)
top10 = counter.most_common(10)
max_cnt = top10[0][1] if top10 else 1
for cat, qtd in top10:
    b = bar(qtd, max_cnt, 15)
    print("  {} {:>6}  {}".format(b, qtd,
        cat.encode("utf-8") if isinstance(cat, unicode) else cat))
print("")

print(SEP2)
print("  FOLHAS POR DISCIPLINA")
print(SEP2)
for k, v in sheet_classes.most_common():
    print("  {:>6}  {}".format(v, k))
print("  --- {} folha(s) com revisao registrada".format(sheets_com_rev))
print("")
cob_str = "  Cobertura obrig. ({}/{}) : {} %  {}".format(
    disc_cobertas_obrig, len(DISC_OBRIGATORIAS),
    cob_obrig_pct, bar(cob_obrig_pct, 100, 15))
print(cob_str)
print("")

print(SEP2)
print("  SISTEMAS CRITICOS BANCARIOS")
print(SEP2)
for sistema, cnt in banco_count.items():
    exs = ", ".join(banco_exemplos[sistema]) if banco_exemplos[sistema] else "-"
    st  = "OK  " if cnt > 0 else "AUSENTE"
    print("  [{}] {:28s}: {:>4}  ex: {}".format(
        st, sistema, cnt,
        exs.encode("utf-8") if isinstance(exs, unicode) else exs))
print("")

print(SEP2)
print("  AMBIENTES CRITICOS BANCARIOS")
print(SEP2)
for amb, cnt in ambientes_count.items():
    tag = "OK" if cnt > 0 else "--"
    print("  [{}] {:30s}: {}".format(
        tag, amb.encode("utf-8") if isinstance(amb, unicode) else amb, cnt))
print("")

print(SEP2)
print("  MATURIDADE BIM  ({} %)  {}  [{}]".format(
    bim_index, bim_nivel_str, bar(bim_index, 100, 20)))
print(SEP2)
for param_name, ppct in sorted(param_pct.items(), key=lambda x: -x[1]):
    print("  {:25s}: {:>5} %  {}".format(param_name, ppct, bar(ppct, 100, 15)))
print("  (Amostra: {} elementos)".format(sample_size))
print("")

print(SEP2)
print("  ALERTAS E SUGESTOES")
print(SEP2)
if alertas:
    for al in alertas:
        print("  [ALERTA] " + (al.encode("utf-8") if isinstance(al, unicode) else al))
if sugestoes:
    for sg in sugestoes:
        print("  [INFO]   " + (sg.encode("utf-8") if isinstance(sg, unicode) else sg))
if not alertas and not sugestoes:
    print("  Nenhum alerta. Modelo dentro dos padroes esperados!")
print(SEP)


# ============================================================
# 11) CSV — TOP CATEGORIAS
# ============================================================

def write_topcats(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "run_stamp", "rank", "categoria", "qtd", "pct_total"])
    for i, (cat, qtd) in enumerate(counter.most_common(TOP_N), start=1):
        w.writerow([u8(model), u8(run_stamp), u8(i),
                    u8(cat), u8(qtd), u8(pct(qtd, total_el))])
    f.close()

write_topcats(out_topcats)
write_topcats(out_topcats_latest)


# ============================================================
# 12) CSV — FOLHAS DETALHADAS
# ============================================================

def write_sheets_csv(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "run_stamp", "numero", "nome", "disciplina",
                "revisao", "data_revisao", "descricao_revisao", "aprovado_por"])
    for s in sheets_data:
        w.writerow([u8(model), u8(run_stamp),
                    u8(s["number"]), u8(s["name"]), u8(s["disc"]),
                    u8(s["rev"]), u8(s["rev_date"]),
                    u8(s["rev_desc"]), u8(s["approved"])])
    f.close()

write_sheets_csv(out_sheets_csv)
write_sheets_csv(out_sheets_latest)


# ============================================================
# 13) CSV — SISTEMAS BANCÁRIOS CRÍTICOS
# ============================================================

def write_banco_csv(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "run_stamp", "sistema", "quantidade",
                "presente", "exemplos"])
    for sistema, cnt in banco_count.items():
        exs = " | ".join(banco_exemplos[sistema])
        w.writerow([u8(model), u8(run_stamp), u8(sistema), u8(cnt),
                    "SIM" if cnt > 0 else "NAO", u8(exs)])
    f.close()

write_banco_csv(out_banco_csv)
write_banco_csv(out_banco_latest)


# ============================================================
# 14) JSON — PERFIL COMPLETO DO MODELO
# ============================================================

profile = {
    "model":        model,
    "run_stamp":    run_stamp,
    "day":          day,
    "time":         ts,
    "run_dir":      RUN_DIR,
    "latest_dir":   LATEST_DIR,
    "metadata":     meta,
    "stats": {
        "total_elements":    total_el,
        "unique_categories": cat_unique,
        "total_families":    total_familias,
        "total_types":       total_tipos,
        "total_sheets":      total_sheets,
        "sheets_com_revisao": sheets_com_rev,
        "total_niveis":      total_niveis,
        "total_rooms":       total_rooms,
        "area_total_m2":     round(area_total_m2, 2),
    },
    "levels":            level_names,
    "top_categories": [
        {"rank": i+1, "category": c, "count": n, "pct": pct(n, total_el)}
        for i, (c, n) in enumerate(counter.most_common(TOP_N))
    ],
    "sheets": {
        "by_discipline": dict(sheet_classes),
        "detail":        sheets_data,
        "cobertura_obrigatorias_pct": cob_obrig_pct,
        "cobertura_recomendadas_pct": cob_recom_pct,
    },
    "sistemas_bancarios": {
        "contagens":  banco_count,
        "exemplos":   banco_exemplos,
    },
    "ambientes_criticos": ambientes_count,
    "bim_maturidade": {
        "index_pct":   bim_index,
        "nivel":       bim_nivel_str,
        "por_parametro": param_pct,
        "amostra":     sample_size,
    },
    "alertas":   alertas,
    "sugestoes": sugestoes,
}


def write_json(path):
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    f.write(u8(json.dumps(profile, ensure_ascii=False, indent=2)))
    f.close()

write_json(out_json)
write_json(out_json_latest)


# ============================================================
# 15) POPUP — RESUMO DENTRO DO REVIT
# ============================================================

def fmt_bar(val, max_v=100, w=16):
    filled = int(round(val * w / float(max_v))) if max_v else 0
    filled = max(0, min(w, filled))
    return u"[" + u"#" * filled + u"." * (w - filled) + u"]"


lines = []
lines.append(u"Modelo: {}  |  {}".format(model, day))
if meta["nome_projeto"]:
    lines.append(u"Projeto: {}".format(meta["nome_projeto"]))
if meta["cliente"]:
    lines.append(u"Cliente: {}".format(meta["cliente"]))
if meta["autor"]:
    lines.append(u"Autor: {}".format(meta["autor"]))
lines.append(u"")

# Estatísticas gerais
lines.append(u"=" * 52)
lines.append(u"  ESTATISTICAS GERAIS")
lines.append(u"=" * 52)
lines.append(u"  Elementos   : {:>6}    Familias  : {:>4}".format(total_el, total_familias))
lines.append(u"  Categorias  : {:>6}    Tipos     : {:>4}".format(cat_unique, total_tipos))
lines.append(u"  Folhas      : {:>6}    Niveis    : {:>4}".format(total_sheets, total_niveis))
lines.append(u"  Ambientes   : {:>6}".format(total_rooms))
if area_total_m2 > 0:
    lines.append(u"  Area total  : {:>6} m2".format(round(area_total_m2, 1)))
lines.append(u"")

# Níveis
if level_names:
    lines.append(u"-- NIVEIS DO PROJETO ---------------------------")
    for ln in level_names[:8]:
        lines.append(u"  " + ln)
    if len(level_names) > 8:
        lines.append(u"  ... e mais {} nivel(is)".format(len(level_names) - 8))
    lines.append(u"")

# Folhas por disciplina
lines.append(u"-- FOLHAS POR DISCIPLINA -----------------------")
for k, v in sheet_classes.most_common():
    lines.append(u"  {:>4}  {}".format(v, k))
lines.append(u"  Cobertura obrig.: {} %  {}".format(
    cob_obrig_pct, fmt_bar(cob_obrig_pct)))
lines.append(u"")

# Sistemas críticos bancários
lines.append(u"-- SISTEMAS CRITICOS BANCARIOS -----------------")
for sistema, cnt in banco_count.items():
    icon = u"[OK]" if cnt > 0 else u"[--]"
    lines.append(u"  {} {:28s}: {}".format(icon, sistema, cnt))
lines.append(u"")

# Ambientes críticos
lines.append(u"-- AMBIENTES CRITICOS --------------------------")
for amb, cnt in ambientes_count.items():
    icon = u"[OK]" if cnt > 0 else u"[--]"
    lines.append(u"  {} {}".format(icon, amb))
lines.append(u"")

# Maturidade BIM
lines.append(u"-- MATURIDADE BIM ------------------------------")
lines.append(u"  Indice: {} %  [{}]  {}".format(
    bim_index, bim_nivel_str, fmt_bar(bim_index)))
lines.append(u"")

# Alertas
if alertas or sugestoes:
    lines.append(u"-- ALERTAS E SUGESTOES -------------------------")
    for al in alertas:
        lines.append(u"  [!] " + al)
    for sg in sugestoes:
        lines.append(u"  [>] " + sg)
    lines.append(u"")

lines.append(u"=" * 52)
lines.append(u"Relatorios salvos em:")
lines.append(u"  " + LATEST_DIR)

msg = u"\n".join(lines)
has_alert = len(alertas) > 0
forms.alert(msg, title=u"Resumo do Modelo — Agencia Bancaria", warn_icon=has_alert)
