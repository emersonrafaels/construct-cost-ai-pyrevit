# -*- coding: utf-8 -*-

# ============================================================
# CONTEXTO RÁPIDO PARA QUEM NÃO CONHECE REVIT
# ============================================================
# Este script valida se um modelo Revit atende aos requisitos mínimos
# de um projeto de instalações (elétrica, infraestrutura de TI e segurança).
#
# A validação é feita em duas frentes:
#   1. Folhas (pranchas): verifica se existem folhas de cada disciplina.
#   2. Elementos: verifica se existem elementos físicos suficientes
#      para cada disciplina (luminárias, conduítes, câmeras, etc.).
#
# Ao final, exibe um popup APROVADO/NÃO APROVADO dentro do Revit
# e salva os resultados detalhados em CSV e JSON para rastreabilidade.
# ============================================================

# pyrevit: ponte entre Python e o Revit.
# forms: módulo do pyRevit para exibir caixas de diálogo nativas do Revit.
from pyrevit import revit, forms

# API oficial do Revit — acesso a elementos, folhas, famílias, etc.
from Autodesk.Revit.DB import *

from collections import Counter  # Contagem de ocorrências por categoria/disciplina
import os, csv, datetime         # Utilitários padrão do Python
import json                      # Para salvar o resultado em formato JSON estruturado

# "doc" = documento (modelo) atualmente aberto no Revit.
doc   = revit.doc
model = doc.Title  # Nome do arquivo .rvt sem extensão

# =========================
# CONFIG (AJUSTE AQUI)
# =========================

# Pasta raiz onde os arquivos de resultado serão salvos.
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"

# ---------------------------------------------------------------
# Limiares mínimos de aprovação (MVP).
# Se o modelo tiver menos do que esses valores, será reprovado.
# ---------------------------------------------------------------
MIN_SHEETS_ELETRICA = 1   # Pelo menos 1 folha de elétrica
MIN_SHEETS_INFRA    = 1   # Pelo menos 1 folha de infraestrutura
MIN_SHEETS_SEG      = 0   # Segurança é opcional no MVP (0 = não valida)

MIN_ELETRICA_ELEMS = 50   # Mín. de elementos elétricos (luminárias, dispositivos, etc.)
MIN_INFRA_ELEMS    = 50   # Mín. de elementos de infra (conduítes, bandejas, etc.)
MIN_SEG_ELEMS      = 1    # Mín. de elementos de segurança detectados por palavra-chave

# ---------------------------------------------------------------
# Palavras-chave para identificar a disciplina de cada folha.
# O script concatena SheetNumber + Name e busca essas strings.
# Exemplo: folha "EL-01 Quadro de Distribuição" → ELETRICA.
# ---------------------------------------------------------------
SHEET_RULES = [
    ("ELETRICA", ["EL", "ELE", "ELÉTR", "E-"]),
    ("INFRA",    ["INF", "TI", "DADOS", "REDE", "RACK", "CABEAMENTO"]),
    ("SEGURANCA",["SEG", "CFTV", "ALARME", "ACESSO", "CAMERA", "CÂMERA"]),
]

# ---------------------------------------------------------------
# Conjuntos de nomes de categoria do Revit (em PT-BR) usados para
# identificar elementos elétricos e de infraestrutura.
# As categorias são os grupos nativos do Revit — equivalente às
# "camadas" de um projeto (cada tipo de objeto tem sua categoria).
# ---------------------------------------------------------------
CAT_ELETRICA = set([
    u"Luminárias",
    u"Dispositivos elétricos",
    u"Equipamentos elétricos",
    u"Circuitos elétricos",
    u"Identificadores de luminária",
    u"Identificadores de dispositivos elétricos",
])

CAT_INFRA = set([
    u"Conduites",
    u"Conexões do conduite",
    u"Fiação",
    u"Bandejas de cabos",
    u"Conexões da bandeja de cabos",
    u"Eletrocalhas",
    u"Conexões da eletrocalha",
])

# Elementos de segurança (câmeras, sensores, etc.) frequentemente são
# modelados como "Dispositivos elétricos" genéricos no Revit PT-BR,
# e não têm categoria própria. Por isso usamos palavras-chave no
# nome da família ou do tipo (FamilyName / Type.Name) para detectá-los.
KW_SEG = [u"CFTV", u"CAMERA", u"CÂMERA", u"ALARME", u"ALARM",
          u"ACESSO", u"ACCESS", u"SENSOR", u"DETECT"]

# =========================

# Captura o momento exato da execução para organizar as pastas de saída.
now       = datetime.datetime.now()
ts        = now.strftime("%H%M%S")           # Hora no formato HHMMSS
day       = now.strftime("%Y-%m-%d")         # Data no formato YYYY-MM-DD
run_stamp = now.strftime("%Y-%m-%d_%H%M%S") # Carimbo completo usado no nome da pasta

# Remove caracteres inválidos em nomes de pasta no Windows.
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)

# Nome do plugin — usado como subpasta dentro de cada execução e do "latest",
# permitindo que vários scripts salvem arquivos no mesmo modelo sem conflito.
PLUGIN_NAME = "ValidadorLayout"

# ---------------------------------------------------------------
# Pasta com carimbo de data/hora — preserva o histórico de cada
# execução. Estrutura gerada:
#   revit_dump/NomeModelo/2026-02-26_143021/ValidadorLayout/
# ---------------------------------------------------------------
RUN_DIR = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)  # Cria a pasta e todos os diretórios intermediários

# ---------------------------------------------------------------
# Pasta "latest" — sempre sobrescrita pela última execução.
# Útil para automatizações que precisam ler sempre o resultado
# mais recente sem precisar descobrir qual pasta é a mais nova.
# Estrutura gerada: revit_dump/NomeModelo/latest/ValidadorLayout/
# ---------------------------------------------------------------
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
if not os.path.exists(LATEST_DIR):
    os.makedirs(LATEST_DIR)

# Caminhos dos arquivos na pasta com carimbo (histórico permanente).
out_csv  = os.path.join(RUN_DIR, "validacao_layout.csv")
out_json = os.path.join(RUN_DIR, "validacao_layout.json")

# Caminhos dos mesmos arquivos na pasta "latest" (sempre sobrescritos).
out_csv_latest  = os.path.join(LATEST_DIR, "validacao_layout.csv")
out_json_latest = os.path.join(LATEST_DIR, "validacao_layout.json")


# ---------------------------------------------------------------
# Normaliza uma string para maiúsculas de forma segura.
# Necessário pelo IronPython 2 ter dois tipos de string (str/unicode)
# e nomes de folhas no Revit poderem conter caracteres especiais.
# ---------------------------------------------------------------
def norm_upper(s):
    try:
        if s is None:
            return u""
        if isinstance(s, unicode):
            return s.upper()
        return unicode(s, errors="ignore").upper()  # "ignore" descarta bytes inválidos
    except:
        try:
            return str(s).upper()
        except:
            return ""


# ---------------------------------------------------------------
# Classifica uma folha por disciplina usando as SHEET_RULES.
# Concatena SheetNumber + Name e verifica cada palavra-chave.
# Se nenhuma bater, retorna "OUTROS".
# ---------------------------------------------------------------
def classify_sheet(sh):
    # Junta número e nome da folha em uma string única normalizada.
    text = norm_upper((sh.SheetNumber or "") + " " + (sh.Name or ""))
    hits = []  # Disciplinas encontradas nesta folha
    for label, keys in SHEET_RULES:
        for k in keys:
            if k in text:
                hits.append(label)
                break  # Já detectou essa disciplina, vai para a próxima
    if not hits:
        return "OUTROS"
    # Múltiplas disciplinas na mesma folha são concatenadas com "+".
    hits = sorted(list(set(hits)))
    return "+".join(hits)


# ---------------------------------------------------------------
# Retorna (FamilyName, TypeName) de um elemento do Revit.
#
# No Revit, cada elemento é uma instância de um "tipo" (Type), que
# pertence a uma "família" (Family). Exemplo: uma câmera pode ser
# do tipo "HD 2MP Dome" da família "Camera_Segurança".
# Usamos FamilyName + TypeName para detectar elementos de segurança
# quando a categoria deles não é específica o suficiente.
# ---------------------------------------------------------------
def get_family_type(el):
    try:
        t = doc.GetElement(el.GetTypeId())  # Obtém o tipo a partir do ID do elemento
        if not t:
            return (u"", u"")
        fam = u""
        try:
            fam = t.FamilyName  # Nome da família (ex: "Camera_Segurança")
        except:
            fam = u""
        return (fam, t.Name)    # t.Name = nome do tipo (ex: "HD 2MP Dome")
    except:
        return (u"", u"")


# ---------------------------------------------------------------
# Converte qualquer valor para bytes UTF-8 de forma segura.
# Necessário para gravar CSVs no IronPython 2 (modo binário "wb").
# ---------------------------------------------------------------
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


# ---------------------------------------------------------------
# Abre um arquivo CSV para escrita com BOM UTF-8 no início.
# O BOM faz o Excel reconhecer automaticamente o encoding UTF-8,
# exibindo acentos e caracteres especiais corretamente.
# ---------------------------------------------------------------
def write_bom_csv(path):
    f = open(path, "wb")               # "wb" = modo de escrita binária
    f.write(u"\ufeff".encode("utf-8")) # \ufeff = caractere BOM
    return f


# ============================================================
# 1) COLETA E CONTAGEM DE FOLHAS POR DISCIPLINA
# ============================================================
# ViewSheet = classe do Revit que representa pranchas de impressão.
# Cada folha tem SheetNumber (ex: "EL-01") e Name (ex: "Planta Baixa").
sheets       = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
sheet_classes = Counter()  # Acumula {disciplina: quantidade_de_folhas}

for sh in sheets:
    try:
        sheet_classes[classify_sheet(sh)] += 1
    except:
        pass

# Uma folha pode ter classificação composta como "ELETRICA+INFRA".
# Por isso, somamos por disciplina verificando se o label CONTÉM a string.
s_ele = 0  # Total de folhas com disciplina elétrica
s_inf = 0  # Total de folhas com disciplina de infraestrutura
s_seg = 0  # Total de folhas com disciplina de segurança
for k, v in sheet_classes.items():
    if "ELETRICA"  in k: s_ele += v
    if "INFRA"     in k: s_inf += v
    if "SEGURANCA" in k: s_seg += v


# ============================================================
# 2) COLETA E CONTAGEM DE ELEMENTOS POR DISCIPLINA
# ============================================================
# Varredura de todos os elementos físicos do modelo.
# .WhereElementIsNotElementType() exclui templates/tipos, traz só instâncias.
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

count_eletrica = 0  # Elementos elétricos (por categoria)
count_infra    = 0  # Elementos de infra (por categoria)
count_seg      = 0  # Elementos de segurança (por palavra-chave no nome da família/tipo)

for el in elements:
    try:
        # Obtém o nome da categoria do elemento (ou string vazia se não tiver).
        cat = el.Category.Name if el.Category else u""

        if cat in CAT_ELETRICA:
            count_eletrica += 1  # Categoria confirmada como elétrica

        if cat in CAT_INFRA:
            count_infra += 1     # Categoria confirmada como infraestrutura

        # Para segurança, busca por palavra-chave no nome da família/tipo.
        fam, typ = get_family_type(el)
        text = norm_upper(fam + u" " + typ)
        for kw in KW_SEG:
            if kw in text:
                count_seg += 1
                break  # Já contou este elemento como segurança, passa para o próximo
    except:
        continue  # Ignora elementos que causarem qualquer exceção


# ============================================================
# 3) APLICAÇÃO DAS REGRAS E DECISÃO DE APROVAÇÃO
# ============================================================
issues = []  # Falhas críticas → reprovam o modelo
warns  = []  # Avisos → não reprovam, mas alertam inconsistências

# --- Validação de folhas mínimas por disciplina ---
if s_ele < MIN_SHEETS_ELETRICA:
    issues.append("Folhas Elétrica abaixo do mínimo ({} < {}).".format(s_ele, MIN_SHEETS_ELETRICA))
if s_inf < MIN_SHEETS_INFRA:
    issues.append("Folhas Infra abaixo do mínimo ({} < {}).".format(s_inf, MIN_SHEETS_INFRA))
if s_seg < MIN_SHEETS_SEG:
    issues.append("Folhas Segurança abaixo do mínimo ({} < {}).".format(s_seg, MIN_SHEETS_SEG))

# --- Validação de elementos mínimos por disciplina ---
if count_eletrica < MIN_ELETRICA_ELEMS:
    issues.append("Elementos Elétrica abaixo do mínimo ({} < {}).".format(count_eletrica, MIN_ELETRICA_ELEMS))
if count_infra < MIN_INFRA_ELEMS:
    issues.append("Elementos Infra abaixo do mínimo ({} < {}).".format(count_infra, MIN_INFRA_ELEMS))
if MIN_SEG_ELEMS > 0 and count_seg < MIN_SEG_ELEMS:
    issues.append("Elementos de Segurança abaixo do mínimo ({} < {}).".format(count_seg, MIN_SEG_ELEMS))

# --- Validação de consistência: tem folhas mas não tem elementos ---
# Isso indica que o projeto foi documentado mas não modelado em 3D.
if s_ele >= 1 and count_eletrica < 10:
    warns.append("Há folhas Elétrica, mas quase nenhum elemento elétrico (possível modelagem incompleta).")
if s_inf >= 1 and count_infra < 10:
    warns.append("Há folhas Infra, mas quase nenhum elemento de infra (possível modelagem incompleta).")

# Aprovado somente se não houver nenhuma falha crítica.
approved = (len(issues) == 0)


# ============================================================
# 4) SALVAR CSV — RESULTADO DA VALIDAÇÃO
# ============================================================
# Grava uma linha por regra verificada, com status PASS/FAIL/WARN.
def write_validation_csv(path):
    f = write_bom_csv(path)
    w = csv.writer(f)
    # Cabeçalho do CSV
    w.writerow(["model", "run_stamp", "tipo", "disciplina", "valor", "minimo", "status", "mensagem"])

    # --- Regras de folhas ---
    w.writerow([u8(model), u8(run_stamp), "sheets", "ELETRICA",
                u8(s_ele), u8(MIN_SHEETS_ELETRICA),
                "PASS" if s_ele >= MIN_SHEETS_ELETRICA else "FAIL",
                "" if s_ele >= MIN_SHEETS_ELETRICA else "Folhas Elétrica abaixo do mínimo."])

    w.writerow([u8(model), u8(run_stamp), "sheets", "INFRA",
                u8(s_inf), u8(MIN_SHEETS_INFRA),
                "PASS" if s_inf >= MIN_SHEETS_INFRA else "FAIL",
                "" if s_inf >= MIN_SHEETS_INFRA else "Folhas Infra abaixo do mínimo."])

    w.writerow([u8(model), u8(run_stamp), "sheets", "SEGURANCA",
                u8(s_seg), u8(MIN_SHEETS_SEG),
                "PASS" if s_seg >= MIN_SHEETS_SEG else "FAIL",
                "" if s_seg >= MIN_SHEETS_SEG else "Folhas Segurança abaixo do mínimo."])

    # --- Regras de elementos ---
    w.writerow([u8(model), u8(run_stamp), "elements", "ELETRICA",
                u8(count_eletrica), u8(MIN_ELETRICA_ELEMS),
                "PASS" if count_eletrica >= MIN_ELETRICA_ELEMS else "FAIL",
                "" if count_eletrica >= MIN_ELETRICA_ELEMS else "Elementos Elétrica abaixo do mínimo."])

    w.writerow([u8(model), u8(run_stamp), "elements", "INFRA",
                u8(count_infra), u8(MIN_INFRA_ELEMS),
                "PASS" if count_infra >= MIN_INFRA_ELEMS else "FAIL",
                "" if count_infra >= MIN_INFRA_ELEMS else "Elementos Infra abaixo do mínimo."])

    w.writerow([u8(model), u8(run_stamp), "elements", "SEGURANCA",
                u8(count_seg), u8(MIN_SEG_ELEMS),
                "PASS" if (MIN_SEG_ELEMS == 0 or count_seg >= MIN_SEG_ELEMS) else "FAIL",
                "" if (MIN_SEG_ELEMS == 0 or count_seg >= MIN_SEG_ELEMS) else "Segurança abaixo do mínimo."])

    # --- Avisos de consistência (não reprovam, mas são registrados) ---
    for w_msg in warns:
        w.writerow([u8(model), u8(run_stamp), "warn", "CONSISTENCIA",
                    "", "", "WARN", u8(w_msg)])

    f.close()

write_validation_csv(out_csv)         # Salva no histórico (pasta com carimbo data/hora)
write_validation_csv(out_csv_latest)  # Salva no "latest" (sobrescreve a execução anterior)


# ============================================================
# 5) SALVAR JSON — RESULTADO COMPLETO DA VALIDAÇÃO
# ============================================================
# JSON estruturado com todos os dados da validação.
# Útil para integrar com dashboards, APIs ou pipelines de BI.
result = {
    "model":      model,       # Nome do arquivo Revit
    "run_stamp":  run_stamp,   # Carimbo de data/hora desta execução
    "day":        day,         # Data (YYYY-MM-DD)
    "time":       ts,          # Hora (HHMMSS)
    "run_dir":    RUN_DIR,     # Pasta do histórico desta execução
    "latest_dir": LATEST_DIR,  # Pasta "latest" (sempre a mais recente)
    "approved":   approved,    # True = APROVADO, False = REPROVADO
    "sheets": {                # Contagem de folhas por disciplina
        "eletrica":  s_ele,
        "infra":     s_inf,
        "seguranca": s_seg,
        "by_class":  dict(sheet_classes),  # Dicionário completo {classificação: qtd}
    },
    "elements": {              # Contagem de elementos físicos por disciplina
        "eletrica":  count_eletrica,
        "infra":     count_infra,
        "seguranca": count_seg,
    },
    "issues": issues,          # Lista de falhas críticas (vazia se aprovado)
    "warns":  warns,           # Lista de avisos de consistência
}


def write_json(path):
    # IronPython 2 não suporta "encoding" em json.dump,
    # por isso escrevemos em binário ("wb") com BOM UTF-8 manualmente.
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    f.write(u8(json.dumps(result, ensure_ascii=False, indent=2)))
    f.close()

write_json(out_json)         # Salva no histórico
write_json(out_json_latest)  # Salva no "latest"

# Console — confirmação de onde os arquivos foram salvos e resumo da validação.
print("=" * 55)
print("  VALIDACAO DE LAYOUT CONCLUIDA")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(RUN_DIR))
print("    - validacao_layout.csv")
print("    - validacao_layout.json")
print("")
print("  [Ultima execucao]")
print("    {}".format(LATEST_DIR))
print("    - validacao_layout.csv")
print("    - validacao_layout.json")
print("")
print("-" * 55)
print("  CONTAGENS")
print("-" * 55)
print("  Folhas   Eletrica  : {}".format(s_ele))
print("  Folhas   Infra     : {}".format(s_inf))
print("  Folhas   Seguranca : {}".format(s_seg))
print("  Elem.    Eletrica  : {}".format(count_eletrica))
print("  Elem.    Infra     : {}".format(count_infra))
print("  Elem.    Seguranca : {}".format(count_seg))
print("")
print("-" * 55)
status_str = "APROVADO" if approved else "NAO APROVADO"
print("  RESULTADO: {}".format(status_str))
print("-" * 55)
if issues:
    print("  Falhas:")
    for it in issues:
        print("    - {}".format(it))
if warns:
    print("  Avisos:")
    for w in warns:
        print("    - {}".format(w))
if approved and not warns:
    print("  Nenhuma falha ou aviso encontrado.")
print("=" * 55)


# ============================================================
# 6) POPUP — RESULTADO FINAL DENTRO DO REVIT
# ============================================================
# forms.alert() exibe uma caixa de diálogo nativa do Revit.
# warn_icon=True → ícone de aviso (triângulo amarelo) quando reprovado.
lines = []
lines.append("Modelo: {}".format(model))
lines.append("")
lines.append("Folhas (inferidas): Eletrica={} | Infra={} | Seguranca={}".format(s_ele, s_inf, s_seg))
lines.append("Elementos: Eletrica={} | Infra={} | Seguranca(kw)={}".format(count_eletrica, count_infra, count_seg))
lines.append("")

if approved:
    title = "APROVADO"
    lines.append("Status: APROVADO")
else:
    title = "NAO APROVADO"
    lines.append("Status: NAO APROVADO")
    lines.append("")
    lines.append("Falhas:")
    for it in issues:
        lines.append("  - " + it)

if warns:
    lines.append("")
    lines.append("Avisos:")
    for w in warns:
        lines.append("  - " + w)

lines.append("")
lines.append("Salvo em: " + LATEST_DIR)

msg = "\n".join(lines)

# Exibe o popup — bloqueia a execução até o usuário fechar.
forms.alert(msg, title=title, warn_icon=(not approved))