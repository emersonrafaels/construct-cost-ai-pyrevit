# -*- coding: utf-8 -*-

# ============================================================
# VALIDADOR DE LAYOUT — AGÊNCIA BANCÁRIA
# ============================================================
# Valida requisitos de layout para projetos de agências bancárias:
#   - Disciplinas: Elétrica, Infraestrutura TI, Segurança, Hidráulica
#   - Ambientes críticos: cofre, sala de segurança, sala de servidores
#   - Sistemas de proteção: SPDA, sprinkler, detecção de fumaça
#   - Acessibilidade PCD (NBR 9050)
#   - Controle de acesso biométrico, CFTV, UPS/No-break, Aterramento
#   - Quadros elétricos (QDL/QGBT), Rack de TI, Alarme anti-intrusão
#   - Score de qualidade de layout (0–100) com classificação em níveis
#
# Resultado: popup APROVADO/NAO APROVADO + CSV + JSON estruturado
# ============================================================

from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from collections import Counter, defaultdict
import os, csv, datetime, json

doc   = revit.doc
model = doc.Title

# =========================
# CONFIG — AJUSTE AQUI
# =========================

BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "RevitScan"   # Pasta raiz para salvar resultados (na área de trabalho)

# ---------------------------------------------------------------
# LIMIARES DE APROVAÇÃO — valores mínimos por categoria
# ---------------------------------------------------------------

# Folhas mínimas por disciplina (0 = não valida)
MIN_SHEETS = {
    "ELETRICA":    1,
    "INFRA":       1,
    "SEGURANCA":   1,   # Obrigatório em banco
    "HIDRAULICA":  0,   # Opcional no MVP
    "ARQUITETURA": 1,
}

# Elementos físicos mínimos modelados por disciplina
MIN_ELEMS = {
    "ELETRICA":  30,
    "INFRA":     10,
    "SEGURANCA":  4,
}

# Sistemas críticos obrigatórios para agência bancária
MIN_SISTEMAS = {
    "CFTV":            4,   # 4 câmeras (entrada, caixa, cofre, pátio)
    "CONTROLE_ACESSO": 2,   # Leitores biométricos / cancelas
    "UPS_NOBREAK":     1,   # No-break para sistemas críticos
    "SPDA":            1,   # Para-raios (NBR 5419)
    "SPRINKLER":       5,   # Sprinklers na área coberta
    "DETECTOR_FUMACA": 4,   # Detectores de fumaça
    "RACK_TI":         1,   # Rack de TI / sala de servidores
    "QUADRO_ELETRICO": 2,   # QDL / QGBT
    "ALARME":          1,   # Sistema de alarme anti-intrusão
    "ACESSIBILIDADE":  1,   # Elemento PCD (NBR 9050)
}

# Peso de cada check no score de qualidade (soma = 100 pontos)
WEIGHTS = {
    # Documentação (folhas) ............ 20 pts
    "sheets_eletrica":    4,
    "sheets_infra":       3,
    "sheets_seguranca":   4,
    "sheets_arquitetura": 4,
    "sheets_hidraulica":  1,
    # Modelagem (elementos) ............ 15 pts
    "elems_eletrica":     6,
    "elems_infra":        4,
    "elems_seguranca":    5,
    # Sistemas críticos bancários ...... 65 pts
    "cftv":              12,
    "controle_acesso":    8,
    "ups_nobreak":        8,
    "spda":               6,
    "sprinkler":          7,
    "detector_fumaca":    6,
    "rack_ti":            5,
    "quadro_eletrico":    5,
    "alarme":             5,
    "acessibilidade":     3,
}

# ---------------------------------------------------------------
# PALAVRAS-CHAVE para classificar folhas por disciplina
# ---------------------------------------------------------------
SHEET_RULES = [
    ("ELETRICA",    ["EL", "ELE", "ELETR", "E-", "ELETRI"]),
    ("INFRA",       ["INF", "TI", "DADOS", "REDE", "RACK", "CABEAMENTO", "TELECOM"]),
    ("SEGURANCA",   ["SEG", "CFTV", "ALARME", "ACESSO", "CAMERA", "CAMARA", "SPDA"]),
    ("HIDRAULICA",  ["HID", "AGUA", "SPKL", "SPRINK", "INCEND", "H-"]),
    ("ARQUITETURA", ["ARQ", "ARQ-", "A-", "PLANT", "LAYOUT", "ARQT"]),
]

# ---------------------------------------------------------------
# CATEGORIAS do Revit que identificam cada disciplina
# ---------------------------------------------------------------
CAT_ELETRICA = set([
    u"Luminarias",
    u"Dispositivos eletricos",
    u"Equipamentos eletricos",
    u"Circuitos eletricos",
    u"Identificadores de luminaria",
    u"Identificadores de dispositivos eletricos",
])

CAT_INFRA = set([
    u"Conduites",
    u"Conexoes do conduite",
    u"Fiacao",
    u"Bandejas de cabos",
    u"Conexoes da bandeja de cabos",
    u"Eletrocalhas",
    u"Conexoes da eletrocalha",
])

CAT_HIDRAULICA = set([
    u"Tubulacoes",
    u"Conexoes de tubulacao",
    u"Sprinklers",
    u"Equipamentos mecanicos",
    u"Aparelhos sanitarios",
    u"Acessorios de tubulacao",
])

# ---------------------------------------------------------------
# PALAVRAS-CHAVE para detectar sistemas críticos bancários
# (buscadas no FamilyName + TypeName + CategoryName de cada elemento)
# ---------------------------------------------------------------
KW_SISTEMAS = {
    "CFTV":            [u"CFTV", u"CAMERA", u"CAMARA", u"DOME", u"BULLET",
                        u"PTZ", u"VIDEOVIGILANCIA", u"CAM_"],
    "CONTROLE_ACESSO": [u"CATRACA", u"CANCELA", u"BIOMETRI", u"LEITOR",
                        u"CONTROLE ACESSO", u"ACCESS CONTROL", u"TORNIQUETE",
                        u"FECHADURA ELET", u"ELETROMAGNETICA"],
    "UPS_NOBREAK":     [u"UPS", u"NO-BREAK", u"NOBREAK", u"NO BREAK",
                        u"GRUPO GERADORA", u"GERADOR", u"GE_", u"BATERIAS"],
    "SPDA":            [u"SPDA", u"PARA-RAIO", u"PARA RAIO", u"ATERRAMENTO",
                        u"MALHA TERRA", u"HASTE", u"DPS", u"DESCARGADOR"],
    "SPRINKLER":       [u"SPRINK", u"CHUVEIRO AUTOM", u"SPKL", u"SPRINKLER",
                        u"BOCAL", u"CABECA SPRINK"],
    "DETECTOR_FUMACA": [u"DETECT", u"FUMACA", u"SMOKE", u"INCEND",
                        u"ACIONADOR", u"CENTRAL INCEND"],
    "RACK_TI":         [u"RACK", u"GABINETE TI", u"SERVIDOR", u"SERVER",
                        u"DATA CENTER", u"SWITCH", u"PATCH PANEL"],
    "QUADRO_ELETRICO": [u"QDL", u"QGBT", u"QUADRO DISTR", u"QUADRO GERAL",
                        u"PAINEL ELET", u"QD_", u"CDL", u"DISTRIBUICAO"],
    "ALARME":          [u"ALARME", u"ALARM", u"ANTI-INTRUS", u"ANTI INTRUS",
                        u"SENSOR MOVIM", u"PIR", u"INFRAVERM"],
    "ACESSIBILIDADE":  [u"RAMPA PCD", u"ACESSIB", u"PCD", u"DEFICIENTE",
                        u"SANITARIO PCD", u"VAGA PCD", u"ROTA ACESS",
                        u"PISO TATIL", u"CALCADA TATIL"],
}

# ---------------------------------------------------------------
# AMBIENTES CRÍTICOS bancários (detectados via Name de Room/Space)
# ---------------------------------------------------------------
KW_AMBIENTES = {
    "COFRE_TESOURARIA": [u"COFRE", u"VAULT", u"SALA_COFRE", u"TESOURARIA"],
    "SALA_SEGURANCA":   [u"VIGILANCIA", u"SEGURANCA", u"MONITORAMENTO", u"CENTRAL SEG"],
    "SALA_SERVIDORES":  [u"SERVIDOR", u"DATACENTER", u"DATA CENTER", u"RACK"],
    "AREA_ATENDIMENTO": [u"ATENDIMENTO", u"CAIXA", u"AGENCIA", u"BANCARIO"],
    "GERENCIA":         [u"GERENCIA", u"DIRETORIA", u"SALA REUNIAO"],
}


# ============================================================
# UTILITARIOS
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
            return ""


def get_family_type(el):
    try:
        t = doc.GetElement(el.GetTypeId())
        if not t:
            return (u"", u"")
        fam = u""
        try:
            fam = t.FamilyName
        except:
            fam = u""
        return (fam, t.Name)
    except:
        return (u"", u"")


def classify_sheet(sh):
    text = norm_upper((sh.SheetNumber or "") + " " + (sh.Name or ""))
    hits = []
    for label, keys in SHEET_RULES:
        for k in keys:
            if k in text:
                hits.append(label)
                break
    if not hits:
        return "OUTROS"
    hits = sorted(list(set(hits)))
    return "+".join(hits)


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


def stars(score):
    filled = int(round(score / 20.0))
    filled = max(0, min(5, filled))
    return u"[" + u"*" * filled + u"." * (5 - filled) + u"]"


def classify_score(score):
    if score >= 90:
        return u"EXCELENTE"
    elif score >= 75:
        return u"BOM"
    elif score >= 55:
        return u"REGULAR"
    elif score >= 35:
        return u"CRITICO"
    else:
        return u"INSUFICIENTE"


# ============================================================
# PATHS DE SAIDA
# ============================================================

now        = datetime.datetime.now()
ts         = now.strftime("%H%M%S")
day        = now.strftime("%Y-%m-%d")
run_stamp  = now.strftime("%Y-%m-%d_%H%M%S")
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")
ROOT_DIR   = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)
PLUGIN_NAME = "ValidadorLayout"
RUN_DIR    = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
for _d in [RUN_DIR, LATEST_DIR]:
    if not os.path.exists(_d):
        os.makedirs(_d)

out_csv        = os.path.join(RUN_DIR,    "validacao_layout.csv")
out_json       = os.path.join(RUN_DIR,    "validacao_layout.json")
out_csv_latest = os.path.join(LATEST_DIR, "validacao_layout.csv")
out_json_latest= os.path.join(LATEST_DIR, "validacao_layout.json")


# ============================================================
# 1) FOLHAS — CONTAGEM POR DISCIPLINA
# ============================================================

sheets        = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
sheet_classes = Counter()
sheet_detail  = defaultdict(list)

for sh in sheets:
    try:
        label = classify_sheet(sh)
        sheet_classes[label] += 1
        tag = u"{} {}".format(sh.SheetNumber or "", sh.Name or "").strip()
        for disc in label.split("+"):
            sheet_detail[disc].append(tag)
    except:
        pass


def sum_sheets(disc):
    total = 0
    for k, v in sheet_classes.items():
        if disc in k:
            total += v
    return total


s_ele = sum_sheets("ELETRICA")
s_inf = sum_sheets("INFRA")
s_seg = sum_sheets("SEGURANCA")
s_hid = sum_sheets("HIDRAULICA")
s_arq = sum_sheets("ARQUITETURA")


# ============================================================
# 2) ELEMENTOS FISICOS — CONTAGEM POR DISCIPLINA E SISTEMAS
# ============================================================

elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

count_eletrica   = 0
count_infra      = 0
count_seguranca  = 0
count_hidraulica = 0

count_sistemas   = {k: 0 for k in KW_SISTEMAS}
exemplos_sistemas= {k: [] for k in KW_SISTEMAS}   # max 3 exemplos por sistema

ambientes_encontrados = {k: 0 for k in KW_AMBIENTES}

# Palavras-chave de segurança (union dos sistemas de seg)
KW_SEG_ALL = (KW_SISTEMAS["CFTV"] + KW_SISTEMAS["CONTROLE_ACESSO"] +
              KW_SISTEMAS["ALARME"])

for el in elements:
    try:
        cat     = el.Category.Name if el.Category else u""
        fam, typ = get_family_type(el)
        full    = norm_upper(fam + u" " + typ + u" " + cat)

        # Disciplinas por categoria
        if cat in CAT_ELETRICA:
            count_eletrica += 1
        if cat in CAT_INFRA:
            count_infra += 1
        if cat in CAT_HIDRAULICA:
            count_hidraulica += 1

        # Segurança por palavra-chave
        for kw in KW_SEG_ALL:
            if norm_upper(kw) in full:
                count_seguranca += 1
                break

        # Sistemas críticos bancários
        for sistema, keywords in KW_SISTEMAS.items():
            matched = False
            for kw in keywords:
                if norm_upper(kw) in full:
                    matched = True
                    break
            if matched:
                count_sistemas[sistema] += 1
                label_ex = (fam.strip() + u" / " + typ.strip()).strip(u" /")
                if label_ex and len(exemplos_sistemas[sistema]) < 3:
                    if label_ex not in exemplos_sistemas[sistema]:
                        exemplos_sistemas[sistema].append(label_ex)

        # Ambientes criticos (Name do elemento Room/Space)
        room_name = u""
        try:
            p = el.LookupParameter("Name")
            if p:
                room_name = norm_upper(p.AsString() or "")
        except:
            pass
        for amb, kws in KW_AMBIENTES.items():
            for kw in kws:
                if norm_upper(kw) in (full + room_name):
                    ambientes_encontrados[amb] += 1
                    break
    except:
        continue


# ============================================================
# 3) VALIDACAO DAS REGRAS E CALCULO DO SCORE
# ============================================================

issues = []   # Falhas criticas (reprovam)
warns  = []   # Avisos (nao reprovam)
checks = {}   # {check_id: {pass, valor, minimo, pontos, max}}


def run_check(check_id, valor, minimo, fail_msg, warn_msg=None, is_warn=False):
    passed = (minimo == 0) or (valor >= minimo)
    pontos = WEIGHTS.get(check_id, 0) if passed else 0
    checks[check_id] = {
        "pass":   passed,
        "valor":  valor,
        "minimo": minimo,
        "pontos": pontos,
        "max":    WEIGHTS.get(check_id, 0),
    }
    if not passed:
        if is_warn:
            if warn_msg:
                warns.append(warn_msg)
        else:
            issues.append(fail_msg)
    return passed


# -- Folhas --
run_check("sheets_eletrica",    s_ele, MIN_SHEETS["ELETRICA"],
          u"Folhas Eletrica ausentes ({}/{}).".format(s_ele, MIN_SHEETS["ELETRICA"]))
run_check("sheets_infra",       s_inf, MIN_SHEETS["INFRA"],
          u"Folhas Infraestrutura TI ausentes ({}/{}).".format(s_inf, MIN_SHEETS["INFRA"]))
run_check("sheets_seguranca",   s_seg, MIN_SHEETS["SEGURANCA"],
          u"Folhas Seguranca ausentes ({}/{}).".format(s_seg, MIN_SHEETS["SEGURANCA"]))
run_check("sheets_arquitetura", s_arq, MIN_SHEETS["ARQUITETURA"],
          u"Folhas Arquitetura/Layout ausentes ({}/{}).".format(s_arq, MIN_SHEETS["ARQUITETURA"]))
run_check("sheets_hidraulica",  s_hid, MIN_SHEETS["HIDRAULICA"],
          u"Folhas Hidraulica/Incendio ausentes.",
          warn_msg=u"Sem folhas de Hidraulica/Incendio detectadas.", is_warn=True)

# -- Elementos --
run_check("elems_eletrica",  count_eletrica,  MIN_ELEMS["ELETRICA"],
          u"Elementos eletricos insuficientes ({}/min {}).".format(
              count_eletrica, MIN_ELEMS["ELETRICA"]))
run_check("elems_infra",     count_infra,     MIN_ELEMS["INFRA"],
          u"Elementos de infraestrutura TI insuficientes ({}/min {}).".format(
              count_infra, MIN_ELEMS["INFRA"]))
run_check("elems_seguranca", count_seguranca, MIN_ELEMS["SEGURANCA"],
          u"Elementos de seguranca insuficientes ({}/min {}).".format(
              count_seguranca, MIN_ELEMS["SEGURANCA"]))

# -- Sistemas criticos bancarios --
run_check("cftv",            count_sistemas["CFTV"],            MIN_SISTEMAS["CFTV"],
          u"CFTV insuficiente: {} camera(s), minimo {}.".format(
              count_sistemas["CFTV"], MIN_SISTEMAS["CFTV"]))
run_check("controle_acesso", count_sistemas["CONTROLE_ACESSO"], MIN_SISTEMAS["CONTROLE_ACESSO"],
          u"Controle de Acesso insuficiente: {} dispositivo(s), minimo {}.".format(
              count_sistemas["CONTROLE_ACESSO"], MIN_SISTEMAS["CONTROLE_ACESSO"]))
run_check("ups_nobreak",     count_sistemas["UPS_NOBREAK"],     MIN_SISTEMAS["UPS_NOBREAK"],
          u"UPS/No-break ausente — sistemas criticos sem redundancia de energia.")
run_check("spda",            count_sistemas["SPDA"],            MIN_SISTEMAS["SPDA"],
          u"SPDA/Aterramento ausente — obrigatorio por NBR 5419.")
run_check("sprinkler",       count_sistemas["SPRINKLER"],       MIN_SISTEMAS["SPRINKLER"],
          u"Sprinklers insuficientes: {} detectado(s), minimo {}.".format(
              count_sistemas["SPRINKLER"], MIN_SISTEMAS["SPRINKLER"]))
run_check("detector_fumaca", count_sistemas["DETECTOR_FUMACA"], MIN_SISTEMAS["DETECTOR_FUMACA"],
          u"Detectores de fumaca insuficientes: {} detectado(s), minimo {}.".format(
              count_sistemas["DETECTOR_FUMACA"], MIN_SISTEMAS["DETECTOR_FUMACA"]))
run_check("rack_ti",         count_sistemas["RACK_TI"],         MIN_SISTEMAS["RACK_TI"],
          u"Rack de TI / Sala de servidores ausente.")
run_check("quadro_eletrico", count_sistemas["QUADRO_ELETRICO"], MIN_SISTEMAS["QUADRO_ELETRICO"],
          u"Quadros eletricos (QDL/QGBT) insuficientes: {} detectado(s), minimo {}.".format(
              count_sistemas["QUADRO_ELETRICO"], MIN_SISTEMAS["QUADRO_ELETRICO"]))
run_check("alarme",          count_sistemas["ALARME"],          MIN_SISTEMAS["ALARME"],
          u"Sistema de alarme anti-intrusao ausente — obrigatorio para agencias bancarias.")
run_check("acessibilidade",  count_sistemas["ACESSIBILIDADE"],  MIN_SISTEMAS["ACESSIBILIDADE"],
          u"Elementos de acessibilidade PCD ausentes (NBR 9050).",
          warn_msg=u"Nenhum elemento de acessibilidade PCD detectado.", is_warn=True)

# -- Consistencia extra (avisos) --
if s_ele >= 1 and count_eletrica < 10:
    warns.append(u"Folhas de eletrica existem mas poucos elementos foram modelados.")
if s_inf >= 1 and count_infra < 5:
    warns.append(u"Folhas de infraestrutura TI existem mas poucos elementos foram modelados.")
if count_sistemas["CFTV"] >= 1 and count_sistemas["CONTROLE_ACESSO"] == 0:
    warns.append(u"CFTV modelado mas sem Controle de Acesso — sistema de seguranca incompleto.")
if count_sistemas["UPS_NOBREAK"] >= 1 and count_sistemas["QUADRO_ELETRICO"] == 0:
    warns.append(u"No-break presente mas nenhum quadro eletrico detectado.")
if ambientes_encontrados.get("COFRE_TESOURARIA", 0) == 0:
    warns.append(u"Ambiente de COFRE/Tesouraria nao detectado como Room/Space.")
if ambientes_encontrados.get("AREA_ATENDIMENTO", 0) == 0:
    warns.append(u"Area de atendimento/caixa nao detectada como ambiente.")

# -- Score --
score_obtido = sum(c["pontos"] for c in checks.values())
score_total  = sum(WEIGHTS.values())
score_pct    = int(round(score_obtido * 100.0 / score_total)) if score_total else 0
nivel        = classify_score(score_pct)
estrelas     = stars(score_pct)

# Aprovado = sem falhas criticas + score >= 55
approved = (len(issues) == 0) and (score_pct >= 55)


# ============================================================
# 4) PRINT NO CONSOLE — RELATORIO DETALHADO
# ============================================================

SEP  = "=" * 65
SEP2 = "-" * 65

print(SEP)
print("  VALIDADOR DE LAYOUT — AGENCIA BANCARIA")
print(SEP)
print("  Modelo : {}".format(model))
print("  Data   : {}  Hora: {}".format(day, now.strftime("%H:%M:%S")))
print("")
print(SEP2)
print("  SCORE DE QUALIDADE: {} % — {}  {}".format(score_pct, nivel, estrelas))
print("  Pontos: {}/{}".format(score_obtido, score_total))
print(SEP2)
print("")
print("  [FOLHAS]")
print("  Arquitetura:{:3d}  Eletrica:{:3d}  Infra TI:{:3d}  Seg:{:3d}  Hid:{:3d}".format(
    s_arq, s_ele, s_inf, s_seg, s_hid))
print("")
print("  [ELEMENTOS FISICOS]")
print("  Eletrica:{:4d}  Infra TI:{:4d}  Seguranca:{:4d}  Hidraulica:{:4d}".format(
    count_eletrica, count_infra, count_seguranca, count_hidraulica))
print("")
print("  [SISTEMAS CRITICOS BANCARIOS]")
for sistema, cnt in count_sistemas.items():
    minimo = MIN_SISTEMAS.get(sistema, 0)
    status = "OK  " if cnt >= minimo else "FAIL"
    exs = ", ".join(exemplos_sistemas[sistema]) if exemplos_sistemas[sistema] else "-"
    print("  [{}] {:20s}: {:3d} / min {:2d}  ex: {}".format(status, sistema, cnt, minimo, exs))
print("")
print("  [AMBIENTES CRITICOS]")
for amb, cnt in ambientes_encontrados.items():
    print("  {:22s}: {}".format(amb, cnt if cnt else "NAO DETECTADO"))
print("")
print(SEP2)
status_str = "APROVADO" if approved else "NAO APROVADO"
print("  RESULTADO: {}  |  {}  |  Score: {} %".format(status_str, nivel, score_pct))
print(SEP2)
if issues:
    print("  FALHAS:")
    for it in issues:
        print("    [FAIL] " + (it.encode("utf-8") if isinstance(it, unicode) else it))
if warns:
    print("  AVISOS:")
    for wm in warns:
        print("    [WARN] " + (wm.encode("utf-8") if isinstance(wm, unicode) else wm))
if not issues and not warns:
    print("  Nenhuma falha ou aviso!")
print(SEP)


# ============================================================
# 5) CSV — RESULTADO DETALHADO
# ============================================================

def write_validation_csv(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    w.writerow(["model", "run_stamp", "categoria", "check_id", "valor", "minimo",
                "pontos_obtidos", "pontos_max", "status", "mensagem"])
    for chk_id, c in checks.items():
        if chk_id.startswith("sheets_"):
            cat = "FOLHAS"
        elif chk_id.startswith("elems_"):
            cat = "ELEMENTOS"
        else:
            cat = "SISTEMA_CRITICO"
        w.writerow([u8(model), u8(run_stamp), cat, u8(chk_id),
                    u8(c["valor"]), u8(c["minimo"]),
                    u8(c["pontos"]), u8(c["max"]),
                    "PASS" if c["pass"] else "FAIL", ""])
    for issue in issues:
        w.writerow([u8(model), u8(run_stamp), "FALHA", "issue",
                    "", "", "0", "0", "FAIL", u8(issue)])
    for wi in warns:
        w.writerow([u8(model), u8(run_stamp), "AVISO", "warn",
                    "", "", "0", "0", "WARN", u8(wi)])
    w.writerow([u8(model), u8(run_stamp), "SCORE", "qualidade",
                u8(score_pct), "55", u8(score_obtido), u8(score_total),
                "PASS" if score_pct >= 55 else "FAIL", u8(nivel)])
    f.close()


write_validation_csv(out_csv)
write_validation_csv(out_csv_latest)


# ============================================================
# 6) JSON — RESULTADO ESTRUTURADO
# ============================================================

result = {
    "model":      model,
    "run_stamp":  run_stamp,
    "day":        day,
    "time":       ts,
    "run_dir":    RUN_DIR,
    "latest_dir": LATEST_DIR,
    "score": {
        "obtido":  score_obtido,
        "total":   score_total,
        "percent": score_pct,
        "nivel":   nivel,
    },
    "approved":  approved,
    "sheets": {
        "eletrica":    s_ele,
        "infra":       s_inf,
        "seguranca":   s_seg,
        "hidraulica":  s_hid,
        "arquitetura": s_arq,
        "by_class":    dict(sheet_classes),
    },
    "elements": {
        "eletrica":   count_eletrica,
        "infra":      count_infra,
        "seguranca":  count_seguranca,
        "hidraulica": count_hidraulica,
    },
    "sistemas_criticos": {
        "contagens": count_sistemas,
        "minimos":   MIN_SISTEMAS,
        "exemplos":  exemplos_sistemas,
    },
    "ambientes_criticos": ambientes_encontrados,
    "checks":  checks,
    "issues":  issues,
    "warns":   warns,
}


def write_json(path):
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    f.write(u8(json.dumps(result, ensure_ascii=False, indent=2)))
    f.close()


write_json(out_json)
write_json(out_json_latest)


# ============================================================
# 7) POPUP — RESULTADO DENTRO DO REVIT
# ============================================================

def chk_icon(passed):
    return u"[OK]" if passed else u"[!!]"


def fmt_line(label, valor, minimo, passed):
    return u"{} {:22s}: {} / {}".format(chk_icon(passed), label, valor, minimo)


lines = []
lines.append(u"Modelo: {}".format(model))
lines.append(u"")
lines.append(u"=" * 50)
lines.append(u"  SCORE DE QUALIDADE: {} %  {}".format(score_pct, estrelas))
lines.append(u"  NIVEL: {}".format(nivel))
lines.append(u"=" * 50)
lines.append(u"")

lines.append(u"-- FOLHAS (Pranchas de Projeto) ---------------")
lines.append(fmt_line(u"Arquitetura",       s_arq, MIN_SHEETS["ARQUITETURA"], checks["sheets_arquitetura"]["pass"]))
lines.append(fmt_line(u"Eletrica",          s_ele,  MIN_SHEETS["ELETRICA"],    checks["sheets_eletrica"]["pass"]))
lines.append(fmt_line(u"Infraestrutura TI", s_inf,  MIN_SHEETS["INFRA"],       checks["sheets_infra"]["pass"]))
lines.append(fmt_line(u"Seguranca",         s_seg,  MIN_SHEETS["SEGURANCA"],   checks["sheets_seguranca"]["pass"]))
lines.append(fmt_line(u"Hidraulica/Incend", s_hid,  MIN_SHEETS["HIDRAULICA"],  checks["sheets_hidraulica"]["pass"]))
lines.append(u"")

lines.append(u"-- ELEMENTOS MODELADOS -------------------------")
lines.append(fmt_line(u"Eletrica",     count_eletrica,  MIN_ELEMS["ELETRICA"],  checks["elems_eletrica"]["pass"]))
lines.append(fmt_line(u"Infra TI",     count_infra,     MIN_ELEMS["INFRA"],     checks["elems_infra"]["pass"]))
lines.append(fmt_line(u"Seguranca",    count_seguranca, MIN_ELEMS["SEGURANCA"], checks["elems_seguranca"]["pass"]))
lines.append(u"")

lines.append(u"-- SISTEMAS CRITICOS BANCARIOS -----------------")
sistemas_labels = [
    ("cftv",            u"CFTV (Cameras)"),
    ("controle_acesso", u"Controle de Acesso"),
    ("ups_nobreak",     u"UPS / No-break"),
    ("spda",            u"SPDA / Aterramento"),
    ("sprinkler",       u"Sprinklers"),
    ("detector_fumaca", u"Detec. Fumaca"),
    ("rack_ti",         u"Rack TI / Servidores"),
    ("quadro_eletrico", u"Quadros Eletricos"),
    ("alarme",          u"Alarme Anti-Intrusao"),
    ("acessibilidade",  u"Acessibilidade PCD"),
]
for chk_id, label in sistemas_labels:
    c = checks[chk_id]
    lines.append(fmt_line(label, c["valor"], c["minimo"], c["pass"]))
lines.append(u"")

lines.append(u"=" * 50)
if approved:
    title = u"APROVADO"
    lines.append(u"  RESULTADO: APROVADO")
else:
    title = u"NAO APROVADO"
    lines.append(u"  RESULTADO: NAO APROVADO  ({} falha(s))".format(len(issues)))
lines.append(u"  Score: {} %  |  {}  |  Pontos: {}/{}".format(
    score_pct, nivel, score_obtido, score_total))
lines.append(u"=" * 50)

if issues:
    lines.append(u"")
    lines.append(u"FALHAS CRITICAS:")
    for it in issues:
        lines.append(u"  [!!] " + it)

if warns:
    lines.append(u"")
    lines.append(u"AVISOS:")
    for wm in warns:
        lines.append(u"  [>] " + wm)

lines.append(u"")
lines.append(u"Relatorio salvo em:")
lines.append(u"  " + LATEST_DIR)

msg = u"\n".join(lines)
forms.alert(msg, title=title, warn_icon=(not approved))
