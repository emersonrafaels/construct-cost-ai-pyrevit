# -*- coding: utf-8 -*-

# ============================================================
# CONTEXTO RAPIDO PARA QUEM NAO CONHECE REVIT
# ============================================================
# Este script extrai os METADADOS de folhas e vistas do modelo Revit.
#
# No Revit, uma "folha" (ViewSheet) e uma prancha de impressao do projeto.
# Dentro de cada folha existem "vistas" (View) posicionadas via "viewports"
# (janelas que enquadram uma vista dentro da folha).
# Exemplos de tipos de vista: FloorPlan (planta), Section (corte),
# Detail (detalhe), ThreeD (perspectiva), etc.
#
# Este script percorre todas as folhas, coleta as vistas vinculadas a cada
# uma e gera um CSV com as informacoes de cada par folha-vista.
# Util para auditar a documentacao do projeto e verificar organizacao.
#
# Arquivos gerados:
#   sheets_views.csv — um par (folha, vista) por linha
#
# Os arquivos sao salvos em duas pastas:
#   - Pasta com carimbo de data/hora (historico permanente)
#   - Pasta "latest" (sobrescrita a cada execucao)
# ============================================================

# pyrevit: biblioteca que faz a ponte entre Python e o Revit aberto.
from pyrevit import revit

# API oficial do Revit — acesso a folhas, vistas, viewports, etc.
from Autodesk.Revit.DB import *

import csv, os, datetime  # Utilitarios padrao do Python

# "doc" = documento (modelo) atualmente aberto no Revit.
doc = revit.doc

# =========================================
# PARAMS (AJUSTE AQUI)
# =========================================

# Pasta raiz onde os arquivos serao salvos (padrao: Area de Trabalho do Windows).
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"

# Nome do plugin — usado como subpasta para organizar os arquivos.
# Estrutura: revit_dump/NomeModelo/ExtrairMetadados/timestamp/
PLUGIN_NAME = "ExtrairMetadados"
# =========================================

# Captura o momento exato da execucao para organizar as pastas de saida.
now       = datetime.datetime.now()
ts        = now.strftime("%H%M%S")           # Hora em formato HHMMSS
day       = now.strftime("%Y-%m-%d")         # Data em formato YYYY-MM-DD
run_stamp = now.strftime("%Y-%m-%d_%H%M%S") # Carimbo completo data+hora

# doc.Title retorna o nome do arquivo .rvt sem extensao.
model = doc.Title

# Remove caracteres invalidos para nomes de pasta no Windows.
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)

# ---------------------------------------------------------------
# Pasta com carimbo de data/hora — preserva historico permanente.
# Estrutura: revit_dump/NomeModelo/ExtrairMetadados/2026-02-26_143021/
# ---------------------------------------------------------------
RUN_DIR = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)  # Cria a pasta e todos os diretorios intermediarios

# ---------------------------------------------------------------
# Pasta "latest" — sempre sobrescrita pela ultima execucao.
# Estrutura: revit_dump/NomeModelo/latest/ExtrairMetadados/
# ---------------------------------------------------------------
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
if not os.path.exists(LATEST_DIR):
    os.makedirs(LATEST_DIR)

# Caminhos na pasta com carimbo (historico permanente).
out = os.path.join(RUN_DIR, "sheets_views.csv")

# Caminhos na pasta "latest" (sempre sobrescritos).
out_latest = os.path.join(LATEST_DIR, "sheets_views.csv")


# ============================================================
# FUNCOES AUXILIARES
# ============================================================

def u8(x):
    """
    Converte qualquer valor para bytes UTF-8 de forma segura.
    Necessario pelo IronPython 2 ter dois tipos de string (str/unicode).
    O modulo csv do Python 2 exige bytes ao gravar em modo binario.
    """
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
    """
    Abre arquivo CSV em modo binario e injeta BOM UTF-8 no inicio.
    O BOM faz o Excel reconhecer o encoding automaticamente,
    exibindo acentos sem precisar importar manualmente.
    """
    f = open(path, "wb")               # "wb" = escrita binaria
    f.write(u"\ufeff".encode("utf-8")) # \ufeff = Byte Order Mark UTF-8
    return f


def safe_str(x):
    """Retorna a string ou '' se None — evita 'None' literal nos CSVs."""
    return x if x else ""


# ============================================================
# COLETA DE FOLHAS E VISTAS
# ============================================================
# ViewSheet = classe do Revit que representa pranchas de impressao.
# Cada folha pode conter multiplas vistas posicionadas via viewports.
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

# ============================================================
# ABERTURA DOS ARQUIVOS CSV (historico + latest)
# ============================================================
f     = write_bom_csv(out)
f_lat = write_bom_csv(out_latest)

w     = csv.writer(f)
w_lat = csv.writer(f_lat)

# Cabecalho do CSV — um par (folha, vista) por linha.
HEADER = [
    "model",           # Nome do arquivo Revit
    "run_stamp",       # Carimbo de data/hora desta execucao
    "sheet_number",    # Numero da folha (ex: "EL-01", "ARQ-05")
    "sheet_name",      # Nome/titulo da folha
    "sheet_id",        # ID interno da folha no modelo
    "view_id",         # ID interno da vista
    "view_name",       # Nome da vista (ex: "Planta Baixa - Terreo")
    "view_type",       # Tipo da vista: FloorPlan, Section, Detail, ThreeD, etc.
    "view_discipline", # Disciplina da vista: Architectural, Mechanical, Electrical, etc.
    "view_scale",      # Escala da vista (ex: 50 = escala 1:50)
    "is_template",     # 1 se a vista e um template (padrao), 0 se e uma vista normal
]
w.writerow(HEADER)
w_lat.writerow(HEADER)

# Contadores para o output final.
total_pairs = 0  # Total de pares (folha, vista) exportados
skipped     = 0  # Pares ignorados por erro de leitura

for sh in sheets:
    try:
        # GetAllViewports() retorna os IDs de todos os viewports da folha.
        # Um viewport e a "janela" que posiciona uma vista dentro da folha.
        vports = sh.GetAllViewports()

        for vpid in vports:
            try:
                vp = doc.GetElement(vpid)  # Objeto Viewport
                v  = doc.GetElement(vp.ViewId)  # Vista contida no viewport

                # Disciplina da vista (Architectural, Mechanical, Electrical, etc.)
                discipline = ""
                try:
                    discipline = str(v.Discipline)
                except:
                    pass

                # Escala da vista (valor numerico, ex: 50 para 1:50).
                # Vistas sem escala definida (ex: legendas) retornam 0.
                scale = ""
                try:
                    scale = str(v.Scale)
                except:
                    pass

                # IsTemplate indica se a vista e um "template de vista" —
                # um modelo reutilizavel que define configuracoes de visibilidade,
                # nao uma vista real que aparece nas folhas.
                is_template = ""
                try:
                    is_template = "1" if v.IsTemplate else "0"
                except:
                    pass

                row = [
                    u8(model),
                    u8(run_stamp),
                    u8(safe_str(sh.SheetNumber)),
                    u8(safe_str(sh.Name)),
                    u8(sh.Id.ToString()),
                    u8(v.Id.ToString()),
                    u8(safe_str(v.Name)),
                    u8(str(v.ViewType)),
                    u8(discipline),
                    u8(scale),
                    u8(is_template),
                ]
                w.writerow(row)
                w_lat.writerow(row)

                total_pairs += 1
            except:
                skipped += 1
                continue
    except:
        skipped += 1
        continue

f.close()
f_lat.close()


# ============================================================
# OUTPUT — Resumo no console do pyRevit
# ============================================================
print("=" * 55)
print("  EXTRACAO DE METADADOS CONCLUIDA")
print("=" * 55)
print("  Modelo : {}".format(model))
print("  Data   : {}".format(day))
print("  Hora   : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(RUN_DIR))
print("    - sheets_views.csv")
print("")
print("  [Ultima execucao]")
print("    {}".format(LATEST_DIR))
print("    - sheets_views.csv")
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Folhas encontradas  : {}".format(len(sheets)))
print("  Pares folha-vista   : {}".format(total_pairs))
print("  Registros ignorados : {}".format(skipped))
print("=" * 55)
print("PROCESSO REALIZADO COM SUCESSO")