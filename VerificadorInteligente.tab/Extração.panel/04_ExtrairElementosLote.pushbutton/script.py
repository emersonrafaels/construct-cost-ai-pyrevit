# -*- coding: utf-8 -*-

# ============================================================
# EXTRACAO EM LOTE -- Multiplos Arquivos .rvt
# ============================================================
# Formulario com duas etapas:
#   1. Selecione o modo:
#        - Extrair Metadados   (sheets_views.csv)
#        - Extrair Completo    (elements.csv + params.csv + json)
#        - Realizar Inventario (inventario_por_folha.csv + json)
#   2. Selecione a pasta com os arquivos .rvt
#
# O script abre cada arquivo, executa a extracao escolhida e
# fecha sem salvar, acumulando os relatorios em:
#   Desktop/RevitScan/<NomeModelo>/<Modo>/<carimbo_hora>/
#   Desktop/RevitScan/<NomeModelo>/latest/<Modo>/
#
# Logica de extracao compartilhada com os botoes individuais:
#   VerificadorInteligente.extension/lib/extracao_lib.py
# ============================================================

from pyrevit import revit, forms, script, output
import os, datetime

from extracao_lib import (
    extrair_metadados,
    extrair_completo,
    extrair_inventario_por_folha,
)
from path_utils import BASE_DIR, ROOT_FOLDER_NAME
from revit_open_utils import DialogSuppressor, open_rvt_detached, close_doc

# Handle da aplicacao Revit (necessario para abrir arquivos)
app = __revit__.Application          # type: ignore  # noqa

# ============================================================
# CONFIG
# ============================================================


# ============================================================
# FORMULARIO -- Passo 1: modo de extracao
# ============================================================
MODOS_LABEL = [
    u"Extrair Metadados",
    u"Extrair Completo",
    u"Realizar Inventario",
]

modo_label = forms.CommandSwitchWindow.show(
    MODOS_LABEL,
    message=u"Selecione o modo de extracao em lote:",
)
if not modo_label:
    script.exit()

if u"Metadados" in modo_label:
    modo = "metadados"
elif u"Completo" in modo_label:
    modo = "completo"
else:
    modo = "inventario"

# ============================================================
# FORMULARIO -- Passo 2: pasta com os arquivos .rvt
# ============================================================
pasta_rvt = forms.pick_folder(title=u"Selecione a pasta com os arquivos .rvt")
if not pasta_rvt:
    script.exit()

# ============================================================
# FORMULARIO -- Passo 3: incluir subpastas?
# ============================================================
incluir_sub = forms.alert(
    u"Incluir subpastas na busca por arquivos .rvt?",
    title=u"Extracao em Lote",
    yes=True,
    no=True,
)

# ============================================================
# LOCALIZA OS ARQUIVOS .rvt
# ============================================================
rvt_files = []
if incluir_sub:
    for root, dirs, files in os.walk(pasta_rvt):
        for f in files:
            if f.lower().endswith(".rvt"):
                rvt_files.append(os.path.join(root, f))
else:
    for f in sorted(os.listdir(pasta_rvt)):
        if f.lower().endswith(".rvt"):
            rvt_files.append(os.path.join(pasta_rvt, f))

if not rvt_files:
    forms.alert(
        u"Nenhum arquivo .rvt encontrado na pasta selecionada.",
        exitscript=True,
    )

# ============================================================
# CONFIRMACAO ANTES DE INICIAR
# ============================================================
confirmacao = forms.alert(
    u"{} arquivo(s) .rvt encontrado(s).\n\nModo: {}\n\nIniciar extracao em lote?".format(
        len(rvt_files), modo_label
    ),
    title=u"Extracao em Lote -- Confirmar",
    yes=True,
    no=True,
)
if not confirmacao:
    script.exit()


# ============================================================
# LOOP PRINCIPAL -- abre, extrai, fecha cada .rvt
# ============================================================
now       = datetime.datetime.now()
run_stamp = now.strftime("%Y-%m-%d_%H%M%S")
day       = now.strftime("%Y-%m-%d")
ts        = now.strftime("%H%M%S")

resultados = []   # [(filename, resumo_str, run_dir)]
erros      = []   # [(filename, mensagem)]

out_panel = output.get_output()
out_panel.print_md(
    u"## Extracao em Lote -- {} arquivo(s)\n"
    u"**Modo:** {}  |  **Inicio:** {}".format(
        len(rvt_files), modo_label, now.strftime("%H:%M:%S")
    )
)
out_panel.print_md(u"---")

with DialogSuppressor(app) as sup:
    for idx, rvt_path in enumerate(rvt_files, start=1):
        filename   = os.path.basename(rvt_path)
        opened_doc = None
        out_panel.print_md(u"**[{}/{}]** `{}`".format(idx, len(rvt_files), filename))

        try:
            opened_doc = open_rvt_detached(app, rvt_path)
            mdl_title  = opened_doc.Title
            mdl_safe   = (mdl_title
                          .replace("/",  "_")
                          .replace("\\", "_")
                          .replace(":",  "_"))

            # ── EXTRAIR METADADOS ──────────────────────────────────
            if modo == "metadados":
                res = extrair_metadados(
                    opened_doc, mdl_title, mdl_safe, run_stamp, day, ts)
                resumo = u"{} folha(s) | {} par(es) folha-vista".format(
                    res["n_sheets"], res["total_pairs"])
                run_dir = res["run_dir"]

            # ── EXTRAIR COMPLETO ───────────────────────────────────
            elif modo == "completo":
                res = extrair_completo(
                    opened_doc, mdl_title, mdl_safe, run_stamp, day, ts)
                resumo = u"{} elemento(s) extraido(s)".format(res["ok"])
                run_dir = res["run_dir"]

            # ── REALIZAR INVENTARIO (por folha) ───────────────────
            else:
                res = extrair_inventario_por_folha(
                    opened_doc, mdl_title, mdl_safe, run_stamp, day, now)
                resumo = u"{} elemento(s) | {} folha(s)".format(
                    res["total_seen"], res["folhas_com_itens"])
                run_dir = res["run_dir"]

            resultados.append((filename, resumo, run_dir))
            out_panel.print_md(u"  &nbsp;&nbsp;OK  {}".format(resumo))

        except Exception as ex:
            erros.append((filename, str(ex)))
            out_panel.print_md(u"  &nbsp;&nbsp;ERRO: {}".format(str(ex)))

        finally:
            close_doc(opened_doc)


# ============================================================
# RELATORIO FINAL NO PAINEL DE OUTPUT
# ============================================================
out_panel.print_md(u"---")
out_panel.print_md(u"## Resultado Final")

for filename, resumo, run_dir in resultados:
    out_panel.print_md(
        u"- **{}** -- {}  \n  Pasta: `{}`".format(filename, resumo, run_dir))

if erros:
    out_panel.print_md(u"### Arquivos com Erro")
    for filename, msg in erros:
        out_panel.print_md(u"- **{}**: {}".format(filename, msg))

out_panel.print_md(
    u"---\n**Concluido:** {}/{} arquivo(s) com sucesso | {} erro(s)".format(
        len(resultados), len(rvt_files), len(erros)
    )
)

# Exibe log de dialogs/warnings suprimidos, se houver.
if sup.log:
    out_panel.print_md(u"### Avisos suprimidos automaticamente ({})".format(len(sup.log)))
    for msg in sup.log[:50]:   # Limita a 50 linhas para nao poluir
        out_panel.print_md(u"- `{}`".format(msg))
    if len(sup.log) > 50:
        out_panel.print_md(u"- ... e mais {} aviso(s) omitido(s).".format(
            len(sup.log) - 50))

# ============================================================
# POPUP DE ENCERRAMENTO
# ============================================================
linhas = [
    u"Extracao em Lote concluida!",
    u"",
    u"  Modo     : {}".format(modo_label),
    u"  Arquivos : {}".format(len(rvt_files)),
    u"  Sucesso  : {}".format(len(resultados)),
    u"  Erros    : {}".format(len(erros)),
    u"",
    u"  Pasta de saida:",
    u"  {}".format(os.path.join(BASE_DIR, ROOT_FOLDER_NAME)),
]

forms.alert(
    u"\n".join(linhas),
    title=u"Extracao em Lote -- Concluida",
    warn_icon=(len(erros) > 0),
)
