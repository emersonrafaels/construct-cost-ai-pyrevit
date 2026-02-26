# -*- coding: utf-8 -*-

from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from collections import Counter

doc = revit.doc
model = doc.Title

# =========================
# CONFIG (AJUSTE AQUI)
# =========================

# Limiares mínimos (MVP)
MIN_SHEETS_ELETRICA = 1
MIN_SHEETS_INFRA    = 1
MIN_SHEETS_SEG      = 0   # opcional no MVP (muitos modelos não têm)

MIN_ELETRICA_ELEMS = 50   # ex: luminárias + dispositivos
MIN_INFRA_ELEMS    = 50   # ex: conduítes + fiação + bandeja
MIN_SEG_ELEMS      = 1    # opcional (depende do modelo)

# Palavras-chave para classificar sheets (mesmo padrão do Resumo)
SHEET_RULES = [
    ("ELETRICA", ["EL", "ELE", "ELÉTR", "E-"]),
    ("INFRA",    ["INF", "TI", "DADOS", "REDE", "RACK", "CABEAMENTO"]),
    ("SEGURANCA",["SEG", "CFTV", "ALARME", "ACESSO", "CAMERA", "CÂMERA"]),
]

# Categorias relevantes (ajuste conforme seu Revit PT-BR)
CAT_ELETRICA = set([u"Luminárias", u"Dispositivos elétricos", u"Equipamentos elétricos",
                    u"Circuitos elétricos", u"Identificadores de luminária",
                    u"Identificadores de dispositivos elétricos"])

CAT_INFRA = set([u"Conduites", u"Conexões do conduite", u"Fiação",
                 u"Bandejas de cabos", u"Conexões da bandeja de cabos",
                 u"Eletrocalhas", u"Conexões da eletrocalha"])

# Segurança geralmente aparece misturado em "Dispositivos elétricos" etc.
# então usamos keywords em family/type:
KW_SEG = [u"CFTV", u"CAMERA", u"CÂMERA", u"ALARME", u"ALARM", u"ACESSO", u"ACCESS", u"SENSOR", u"DETECT"]

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


# -------------------------
# 1) Sheets check
# -------------------------
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
sheet_classes = Counter()

for sh in sheets:
    try:
        sheet_classes[classify_sheet(sh)] += 1
    except:
        pass

# soma por disciplina (quando vem "ELETRICA+INFRA" etc.)
s_ele = 0
s_inf = 0
s_seg = 0
for k, v in sheet_classes.items():
    if "ELETRICA" in k:
        s_ele += v
    if "INFRA" in k:
        s_inf += v
    if "SEGURANCA" in k:
        s_seg += v

# -------------------------
# 2) Elements check
# -------------------------
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

count_eletrica = 0
count_infra = 0
count_seg = 0

for el in elements:
    try:
        cat = el.Category.Name if el.Category else u""
        if cat in CAT_ELETRICA:
            count_eletrica += 1
        if cat in CAT_INFRA:
            count_infra += 1

        # segurança por keyword no family/type
        fam, typ = get_family_type(el)
        text = norm_upper(fam + u" " + typ)
        for kw in KW_SEG:
            if kw in text:
                count_seg += 1
                break
    except:
        continue

# -------------------------
# 3) Regras e decisão
# -------------------------
issues = []
warns = []

# Sheets mínimos
if s_ele < MIN_SHEETS_ELETRICA:
    issues.append("Folhas Elétrica abaixo do mínimo ({} < {}).".format(s_ele, MIN_SHEETS_ELETRICA))
if s_inf < MIN_SHEETS_INFRA:
    issues.append("Folhas Infra abaixo do mínimo ({} < {}).".format(s_inf, MIN_SHEETS_INFRA))
if s_seg < MIN_SHEETS_SEG:
    issues.append("Folhas Segurança abaixo do mínimo ({} < {}).".format(s_seg, MIN_SHEETS_SEG))

# Elementos mínimos
if count_eletrica < MIN_ELETRICA_ELEMS:
    issues.append("Elementos Elétrica abaixo do mínimo ({} < {}).".format(count_eletrica, MIN_ELETRICA_ELEMS))
if count_infra < MIN_INFRA_ELEMS:
    issues.append("Elementos Infra abaixo do mínimo ({} < {}).".format(count_infra, MIN_INFRA_ELEMS))
if MIN_SEG_ELEMS > 0 and count_seg < MIN_SEG_ELEMS:
    issues.append("Indícios de Segurança abaixo do mínimo ({} < {}).".format(count_seg, MIN_SEG_ELEMS))

# Consistência: tem folhas mas não tem elementos
if s_ele >= 1 and count_eletrica < 10:
    warns.append("Há folhas Elétrica, mas quase nenhum elemento elétrico (possível modelagem incompleta).")
if s_inf >= 1 and count_infra < 10:
    warns.append("Há folhas Infra, mas quase nenhum elemento de infra (possível modelagem incompleta).")

approved = (len(issues) == 0)

# -------------------------
# 4) Mensagem final (popup)
# -------------------------
lines = []
lines.append("Modelo: {}".format(model))
lines.append("")
lines.append("📄 Folhas (inferidas): Elétrica={} | Infra={} | Segurança={}".format(s_ele, s_inf, s_seg))
lines.append("🏗️ Elementos: Elétrica={} | Infra={} | Segurança(kw)={}".format(count_eletrica, count_infra, count_seg))
lines.append("")

if approved:
    title = "✅ APROVADO"
    lines.append("Status: APROVADO")
else:
    title = "❌ NÃO APROVADO"
    lines.append("Status: NÃO APROVADO")
    lines.append("")
    lines.append("Falhas:")
    for it in issues:
        lines.append("- " + it)

if warns:
    lines.append("")
    lines.append("Avisos:")
    for w in warns:
        lines.append("- " + w)

msg = "\n".join(lines)

# show popup
forms.alert(msg, title=title, warn_icon=(not approved))