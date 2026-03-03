# -*- coding: utf-8 -*-

# ============================================================
# EXTRAIR METADADOS -- Folhas e Vistas
# ============================================================
# Extrai metadados de folhas (ViewSheet) e vistas do modelo.
# Cada folha pode conter multiplas vistas via Viewports.
#
# Arquivo gerado:
#   sheets_views.csv -- um par (folha, vista) por linha
#
# Util para auditar a documentacao do projeto e verificar
# a organizacao das pranchas.
#
# A logica de extracao vive em:
#   VerificadorInteligente.extension/lib/extracao_lib.py
# ============================================================

from pyrevit import revit
import datetime

from extracao_lib import extrair_metadados

# ============================================================
# CONTEXTO DO MODELO ATUAL
# ============================================================
doc   = revit.doc
model = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

now       = datetime.datetime.now()
run_stamp = now.strftime("%Y-%m-%d_%H%M%S")
day       = now.strftime("%Y-%m-%d")
ts        = now.strftime("%H%M%S")

# ============================================================
# EXECUCAO
# ============================================================
print("Iniciando extracao de metadados (folhas e vistas)...")
res = extrair_metadados(doc, model, model_safe, run_stamp, day, ts)

# ============================================================
# RELATORIO NO CONSOLE DO PYREVIT
# ============================================================
print("=" * 55)
print("  EXTRACAO DE METADADOS CONCLUIDA")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(res["run_dir"]))
print("    - sheets_views.csv")
print("")
print("  [Ultima execucao]")
print("    {}".format(res["latest_dir"]))
print("    - sheets_views.csv")
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Folhas encontradas  : {}".format(res["n_sheets"]))
print("  Pares folha-vista   : {}".format(res["total_pairs"]))
print("  Registros ignorados : {}".format(res["skipped"]))
print("=" * 55)
print("PROCESSO REALIZADO COM SUCESSO")
