# -*- coding: utf-8 -*-

# ============================================================
# REALIZAR INVENTARIO POR FOLHA -- Layout Itau
# ============================================================
# Gera inventario de equipamentos agrupado por Folha (Sheet).
# Cada elemento e contado apenas uma vez por folha, mesmo que
# apareca em multiplas vistas da mesma folha.
#
# Arquivos gerados:
#   inventario_por_folha.csv -- (folha x elemento x quantidade)
#   inventario_totais.csv    -- totais globais por tipo
#   inventario.json          -- dados estruturados completos
#
# A logica de extracao vive em:
#   VerificadorInteligente.extension/lib/extracao_lib.py
# ============================================================

from pyrevit import revit, forms
import datetime

from extracao_lib import extrair_inventario_por_folha

# ============================================================
# CONFIGURACOES
# ============================================================
MAX_POPUP_FOLHAS = 20  # Maximo de folhas exibidas no popup de resumo

# ============================================================
# CONTEXTO DO MODELO ATUAL
# ============================================================
doc   = revit.doc
model = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

now       = datetime.datetime.now()
run_stamp = now.strftime("%Y-%m-%d_%H%M%S")
day       = now.strftime("%Y-%m-%d")

# ============================================================
# EXECUCAO
# ============================================================
print("Coletando folhas e associando elementos...")
res = extrair_inventario_por_folha(doc, model, model_safe, run_stamp, day, now)

sheets_data      = res["sheets_data"]
inventario       = res["inventario"]
totais           = res["totais"]
folhas_com_itens = res["folhas_com_itens"]
folhas_sem_itens = res["folhas_sem_itens"]
total_seen       = res["total_seen"]
RUN_DIR          = res["run_dir"]
LATEST_DIR       = res["latest_dir"]

# ============================================================
# RELATORIO NO CONSOLE DO PYREVIT
# ============================================================
SEP  = "=" * 65
SEP2 = "-" * 65

print(SEP)
print("  INVENTARIO POR FOLHA -- LAYOUT ITAU")
print(SEP)
print("  Modelo : {}".format(model))
print("  Data   : {}  Hora: {}".format(day, now.strftime("%H:%M:%S")))
print("")
print("  Total de folhas       : {}".format(len(sheets_data)))
print("  Folhas com itens      : {}".format(folhas_com_itens))
print("  Folhas sem itens      : {}".format(folhas_sem_itens))
print("  Elementos catalogados : {}".format(total_seen))
print("  Tipos distintos       : {}".format(len(totais)))
print("")

sorted_sheets = sorted(
    [(sid, sheets_data[sid]) for sid in inventario if sid in sheets_data],
    key=lambda x: x[1]["numero"]
)

print(SEP2)
print("  INVENTARIO POR FOLHA")
print(SEP2)

for sid, sh in sorted_sheets:
    itens    = inventario[sid]
    total_sh = sum(itens.values())
    print("  [{}] {} -- {} item(ns)".format(sh["numero"], sh["nome"], total_sh))
    for (fam, typ), cnt in sorted(itens.items(), key=lambda x: -x[1]):
        label = typ or fam or u"(sem nome)"
        print("      {:>4}x  {}".format(cnt, label))

print("")
print(SEP2)
print("  TOP 15 MAIS FREQUENTES (GLOBAL)")
print(SEP2)
for i, ((fam, typ), cnt) in enumerate(
        sorted(totais.items(), key=lambda x: -x[1])[:15], start=1):
    label = typ or fam or u"(sem nome)"
    print("  {:>2}. {:>5}x  {}".format(i, cnt, label))

print("")
print(SEP2)
print("  ARQUIVOS GERADOS")
print(SEP2)
print("  [Historico]  {}".format(RUN_DIR))
print("    inventario_por_folha.csv / inventario_totais.csv / inventario.json")
print("  [Latest]     {}".format(LATEST_DIR))
print(SEP)


# ============================================================
# POPUP DE RESUMO
# ============================================================
lines = [
    u"Modelo: {}  |  {}".format(model, day),
    u"",
    u"=" * 52,
    u"  RESUMO DO INVENTARIO POR FOLHA",
    u"=" * 52,
    u"  Total de folhas       : {}".format(len(sheets_data)),
    u"  Folhas com itens      : {}".format(folhas_com_itens),
    u"  Elementos catalogados : {}".format(total_seen),
    u"  Tipos distintos       : {}".format(len(totais)),
    u"",
    u"-- TOP {} FOLHAS --------------------".format(MAX_POPUP_FOLHAS),
]

for sid, sh in sorted_sheets[:MAX_POPUP_FOLHAS]:
    itens    = inventario[sid]
    total_sh = sum(itens.values())
    lines.append(u"  [{}] {} -- {} item(ns)".format(
        sh["numero"], sh["nome"], total_sh))
    for (fam, typ), cnt in sorted(itens.items(), key=lambda x: -x[1])[:5]:
        label = typ or fam or u"(sem nome)"
        lines.append(u"      {:>4}x  {}".format(cnt, label))
    if len(itens) > 5:
        lines.append(u"      ... e mais {} tipo(s)".format(len(itens) - 5))

if len(sorted_sheets) > MAX_POPUP_FOLHAS:
    lines.append(u"  ... e mais {} folha(s). Veja o CSV!".format(
        len(sorted_sheets) - MAX_POPUP_FOLHAS))

lines += [
    u"",
    u"-- TOP 10 ELEMENTOS -----------------",
]
for i, ((fam, typ), cnt) in enumerate(
        sorted(totais.items(), key=lambda x: -x[1])[:10], start=1):
    label = typ or fam or u"(sem nome)"
    lines.append(u"  {:>2}. {:>4}x  {}".format(i, cnt, label))

lines += [
    u"",
    u"=" * 52,
    u"Relatorios salvos em:",
    u"  " + LATEST_DIR,
]

forms.alert(
    u"\n".join(lines),
    title=u"Inventario por Folha -- Layout Itau",
    warn_icon=False,
)
