# -*- coding: utf-8 -*-

# ============================================================
# EXTRAIR ELEMENTOS COMPLETO
# ============================================================
# Extrai todos os elementos fisicos do modelo Revit gerando:
#   elements.csv         -- um elemento por linha com metadados
#   params.csv           -- todos os parametros (instancia + tipo)
#   model_hierarchy.json -- hierarquia Categoria > Familia > Tipo
#
# A logica de extracao vive em:
#   VerificadorInteligente.extension/lib/extracao_lib.py
# ============================================================

from pyrevit import revit
import datetime

from extracao_lib import extrair_completo

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
print("Iniciando extracao de elementos completa...")
res = extrair_completo(doc, model, model_safe, run_stamp, day, ts)

# ============================================================
# RELATORIO NO CONSOLE DO PYREVIT
# ============================================================
hierarchy = res["hierarchy"]

print("=" * 55)
print("  EXTRACAO DE ELEMENTOS COMPLETA")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(res["run_dir"]))
print("    - elements.csv")
print("    - params.csv")
print("    - model_hierarchy.json")
print("")
print("  [Ultima execucao]")
print("    {}".format(res["latest_dir"]))
print("    - elements.csv")
print("    - params.csv")
print("    - model_hierarchy.json")
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Elementos extraidos : {}".format(res["ok"]))
print("  Elementos pulados   : {}".format(res["skipped"]))
print("  Total no collector  : {}".format(res["n_elements"]))
print("  Categorias unicas   : {}".format(len(hierarchy)))
print("")
print("-" * 55)
print("  HIERARQUIA (Top 15 categorias)")
print("-" * 55)
for cat_name, cat_data in sorted(hierarchy.items(),
                                  key=lambda x: -x[1]["total"])[:15]:
    n_fams  = len(cat_data["families"])
    n_types = sum(len(f["types"]) for f in cat_data["families"].values())
    print("  {:>5}  {}  ({} familias, {} tipos)".format(
        cat_data["total"], cat_name, n_fams, n_types))
print("=" * 55)
print("PROCESSO REALIZADO COM SUCESSO")
