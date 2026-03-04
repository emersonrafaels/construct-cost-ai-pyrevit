# -*- coding: utf-8 -*-

# ============================================================
# REALIZAR INVENTARIO POR PAVIMENTO / AMBIENTE
# ============================================================
# Gera inventario hierarquico:
#   Pavimento -> Ambiente -> Categoria -> Familia -> Tipo
#
# Para cada Tipo coleta:
#   Codigo SAP  (parametro da instancia ou do tipo)
#   Nome da familia
#   Quantidade de instancias aprovadas
#
# Configuracao centralizada em:
#   VerificadorInteligente.extension/lib/config_inventario.py
#
# Arquivos gerados:
#   inventario_por_pavimento.csv  -- linha por Tipo com totais
#   inventario_por_pavimento.json -- hierarquia completa + flat_items
#   conformidade.json             -- relatorio de qualidade
# ============================================================

from pyrevit import revit, forms
import datetime

# ----------------------------------------------------------
# Forca reimport a cada execucao (cache IronPython)
# ----------------------------------------------------------
import config_inventario
try:
    reload(config_inventario)
except NameError:
    import importlib
    importlib.reload(config_inventario)

import extracao_lib
try:
    reload(extracao_lib)
except NameError:
    import importlib
    importlib.reload(extracao_lib)

from extracao_lib import extrair_inventario_por_pavimento

# ============================================================
# CONFIGURACOES LOCAIS (sobrepõem config_inventario.py)
# ============================================================
MAX_POPUP_PAV = 10   # Pavimentos exibidos no popup de resumo
MAX_POPUP_TOP = 15   # Tipos mais frequentes no popup

# Descomente e ajuste para sobrescrever o config_inventario.py:
# FILTRO_FASE_CRIACAO    = u"Construção nova"
# FILTRO_FASE_DEMOLICAO  = True

# Quando None/False, os valores vêm de config_inventario.py:
FILTRO_FASE_CRIACAO   = None
FILTRO_FASE_DEMOLICAO = False

# ============================================================
# CONTEXTO DO MODELO ATUAL
# ============================================================
doc        = revit.doc
model      = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

now       = datetime.datetime.now()
run_stamp = now.strftime("%Y-%m-%d_%H%M%S")
day       = now.strftime("%Y-%m-%d")

# ============================================================
# EXECUCAO
# ============================================================
print("Coletando elementos e montando inventario por pavimento...")
_t0 = datetime.datetime.now()
try:
    res = extrair_inventario_por_pavimento(
        doc, model, model_safe, run_stamp, day, now,
        filtro_fase_criacao=FILTRO_FASE_CRIACAO,
        filtro_fase_demolicao=FILTRO_FASE_DEMOLICAO,
        cfg=config_inventario,
    )
except Exception as _exc:
    forms.alert(
        u"Erro ao gerar inventario:\n\n{}".format(_exc),
        title=u"Inventario por Pavimento -- Erro",
        warn_icon=True,
    )
    raise
_elapsed = (datetime.datetime.now() - _t0).seconds

inventario     = res["inventario"]      # {pav: {amb: {cat: {fam: {typ: info}}}}}
totais         = res["totais_globais"]  # {(cat, fam, typ): count}
total_incluido = res["total_incluido"]
total_ignorado = res["total_ignorado"]
conformidade   = res["conformidade"]
RUN_DIR        = res["run_dir"]
LATEST_DIR     = res["latest_dir"]

# ============================================================
# RELATORIO NO CONSOLE DO PYREVIT
# ============================================================
SEP  = "=" * 68
SEP2 = "-" * 68

_fc   = getattr(config_inventario, "DEFAULT_FASE_CRIACAO", None)
_fd   = getattr(config_inventario, "DEFAULT_EXCLUIR_DEMOLIDOS", False)
_fase_c = FILTRO_FASE_CRIACAO  or _fc  or u"(desativado)"
_fase_d = FILTRO_FASE_DEMOLICAO or _fd

print(SEP)
print("  INVENTARIO POR PAVIMENTO / AMBIENTE")
print(SEP)
print("  Modelo  : {}".format(model))
print("  Data    : {}  Hora: {}".format(day, now.strftime("%H:%M:%S")))
print(u"  Config  : lib/config_inventario.py")
print(u"  Filtros : fase_criacao={}  |  excluir_demolidos={}".format(
    _fase_c, _fase_d))
_ign_cats = getattr(config_inventario, "IGNORED_CATEGORIES", [])
if _ign_cats:
    print(u"  Categorias ignoradas ({}) : {}".format(
        len(_ign_cats), u", ".join(_ign_cats[:6]) +
        (u"..." if len(_ign_cats) > 6 else u"")))
print("")
print("  Instancias incluidas : {}".format(total_incluido))
print("  Instancias ignoradas : {}".format(total_ignorado))
print("  Pavimentos           : {}".format(len(inventario)))
print("  Tempo de execucao    : {}s".format(_elapsed))
print("")

for pav in sorted(inventario.keys()):
    ambientes = inventario[pav]
    total_pav = sum(
        info["count"]
        for amb in ambientes.values()
        for cat in amb.values()
        for fam in cat.values()
        for info in fam.values()
    )
    print(SEP2)
    print("  PAVIMENTO: {}  ({} instancia(s))".format(pav, total_pav))
    print(SEP2)
    for amb in sorted(ambientes.keys()):
        categorias = ambientes[amb]
        total_amb  = sum(
            info["count"]
            for cat in categorias.values()
            for fam in cat.values()
            for info in fam.values()
        )
        print("    Ambiente: {}  ({} instancia(s))".format(amb, total_amb))
        for cat in sorted(categorias.keys()):
            familias = categorias[cat]
            print("      [{}]".format(cat))
            for fam in sorted(familias.keys()):
                tipos = familias[fam]
                for typ, info in sorted(tipos.items()):
                    label = typ or fam or u"(sem nome)"
                    sap   = info["sap_code"]
                    sap_s = u"  SAP:{}".format(sap) if sap else u"  [sem SAP]"
                    print(u"        {:>4}x  {}{}".format(
                        info["count"], label, sap_s))

print("")
print(SEP2)
print("  TOP {} MAIS FREQUENTES (GLOBAL)".format(MAX_POPUP_TOP))
print(SEP2)
for i, ((cat, fam, typ), cnt) in enumerate(
        sorted(totais.items(), key=lambda x: -x[1])[:MAX_POPUP_TOP], start=1):
    label = typ or fam or u"(sem nome)"
    print("  {:>2}. {:>5}x  {}  [{}]".format(i, cnt, label, cat))

# ============================================================
# RELATORIO DE CONFORMIDADE NO CONSOLE
# ============================================================
print("")
print(SEP)
print("  RELATORIO DE CONFORMIDADE")
print(SEP)
_conf_sap = conformidade["sem_codigo_sap"]
_conf_pav = conformidade["sem_pavimento"]
_conf_amb = conformidade["sem_ambiente"]
print("  Sem Codigo SAP : {}  ({:.1f}%)".format(
    _conf_sap["count"], _conf_sap["percentual"]))
if _conf_sap["por_categoria"]:
    _top_cat = sorted(_conf_sap["por_categoria"].items(),
                      key=lambda x: -x[1])[:5]
    print("    Top categorias sem SAP:")
    for _c, _n in _top_cat:
        print("      {:>5}x  {}".format(_n, _c))
print("  Sem Pavimento  : {}  ({:.1f}%)".format(
    _conf_pav["count"], _conf_pav["percentual"]))
print("  Sem Ambiente   : {}  ({:.1f}%)".format(
    _conf_amb["count"], _conf_amb["percentual"]))
if conformidade["erros_extracao"] > 0:
    print("  Erros (excecao): {}".format(conformidade["erros_extracao"]))
print("")
print(SEP2)
print("  ARQUIVOS GERADOS")
print(SEP2)
print("  [Historico]  {}".format(RUN_DIR))
print("    inventario_por_pavimento.csv")
print("    inventario_por_pavimento.json  (hierarquia + flat_items)")
print("    conformidade.json")
print("  [Latest]     {}".format(LATEST_DIR))
print(SEP)


# ============================================================
# POPUP DE RESUMO
# ============================================================
lines = [
    u"Modelo: {}  |  {}".format(model, day),
    u"Filtros: fase_criacao={}  |  excluir_demolidos={}".format(
        _fase_c, _fase_d),
    u"",
    u"=" * 52,
    u"  RESUMO DO INVENTARIO POR PAVIMENTO",
    u"=" * 52,
    u"  Instancias incluidas : {}".format(total_incluido),
    u"  Instancias ignoradas : {}".format(total_ignorado),
    u"  Pavimentos           : {}".format(len(inventario)),
    u"  Tempo de execucao    : {}s".format(_elapsed),
    u"",
    u"-- QUALIDADE ----------------------------------",
    u"  Sem Codigo SAP : {}  ({:.1f}%)".format(
        _conf_sap["count"], _conf_sap["percentual"]),
    u"  Sem Pavimento  : {}  ({:.1f}%)".format(
        _conf_pav["count"], _conf_pav["percentual"]),
    u"  Sem Ambiente   : {}  ({:.1f}%)".format(
        _conf_amb["count"], _conf_amb["percentual"]),
    u"",
    u"-- PAVIMENTOS ({}) -----------------------".format(MAX_POPUP_PAV),
]

for pav in sorted(inventario.keys())[:MAX_POPUP_PAV]:
    ambientes = inventario[pav]
    total_pav = sum(
        info["count"]
        for amb in ambientes.values()
        for cat in amb.values()
        for fam in cat.values()
        for info in fam.values()
    )
    lines.append(u"  {} -- {} instancia(s)".format(pav, total_pav))
    for amb in sorted(ambientes.keys())[:5]:
        categorias = ambientes[amb]
        total_amb  = sum(
            info["count"]
            for cat in categorias.values()
            for fam in cat.values()
            for info in fam.values()
        )
        lines.append(u"    Amb: {}  ({} inst.)".format(amb, total_amb))
    n_amb = len(ambientes)
    if n_amb > 5:
        lines.append(u"    ... e mais {} ambiente(s)".format(n_amb - 5))

if len(inventario) > MAX_POPUP_PAV:
    lines.append(u"  ... e mais {} pavimento(s). Veja o CSV!".format(
        len(inventario) - MAX_POPUP_PAV))

lines += [
    u"",
    u"-- TOP {} ELEMENTOS ---------------------".format(MAX_POPUP_TOP),
]
for i, ((cat, fam, typ), cnt) in enumerate(
        sorted(totais.items(), key=lambda x: -x[1])[:MAX_POPUP_TOP], start=1):
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
    title=u"Inventario por Pavimento / Ambiente",
    warn_icon=False,
)
