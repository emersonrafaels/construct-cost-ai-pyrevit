# -*- coding: utf-8 -*-

# ============================================================
# CONTEXTO RAPIDO PARA QUEM NAO CONHECE REVIT
# ============================================================
# Este script e um DIAGNOSTICO COMPLETO do modelo Revit aberto.
# Ele demonstra APIs do Revit que NAO foram usadas nos demais scripts:
#
#   - ProjectInfo        : metadados do projeto (cliente, end., autor...)
#   - __revit__.Application : versao do Revit em uso
#   - Level              : pavimentos / niveis com elevacao
#   - FilteredWorksetCollector : worksets (conjuntos de trabalho BIM colaborativo)
#   - doc.GetWarnings()  : alertas/inconsistencias do modelo
#   - RevitLinkInstance  : modelos Revit linkados (RVT externos)
#   - doc.Phases         : fases de construcao do projeto
#   - DesignOption       : opcoes de projeto (variantes de layout)
#   - script.get_output(): painel de saida HTML rico do pyRevit (NOVO)
#   - output.print_table(): tabela renderizada em HTML no painel pyRevit
#   - output.print_md()  : renderiza Markdown no painel pyRevit
#   - output.linkify()   : gera link clicavel para um elemento no Revit
#
# Alem do painel HTML, gera um JSON de "fingerprint" do modelo —
# util para comparar versoes do modelo ao longo do tempo.
# ============================================================

from pyrevit import revit, script   # script: modulo do painel de saida HTML
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FilteredWorksetCollector,       # Coleta worksets (novo)
    WorksetKind,                    # Enum para filtrar por tipo de workset
    Level,                          # Classe de niveis/pavimentos (novo)
    RevitLinkInstance,              # Modelos RVT linkados (novo)
    DesignOption,                   # Opcoes de projeto (novo)
    UnitUtils,                      # Conversao de unidades internas do Revit (novo)
    UnitTypeId,                     # Identificadores de tipo de unidade (Revit 2022+)
)
from collections import Counter, OrderedDict
import csv, os, datetime, json

# --------------------------------------------------------
# doc  = modelo aberto no Revit
# app  = objeto Application do Revit (versao, idioma, etc.)
# __revit__ e injetado automaticamente pelo pyRevit no escopo do script
# --------------------------------------------------------
doc = revit.doc
app = __revit__.Application  # Objeto principal do Revit (novo — nao usado antes)

# =========================================
# PARAMS (AJUSTE AQUI)
# =========================================
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "revit_dump"
PLUGIN_NAME      = "TestAPI"
TOP_N_CATS       = 20  # Quantas categorias exibir no ranking
# =========================================

now       = datetime.datetime.now()
ts        = now.strftime("%H%M%S")
day       = now.strftime("%Y-%m-%d")
run_stamp = now.strftime("%Y-%m-%d_%H%M%S")

model      = doc.Title
model_safe = model.replace("/", "_").replace("\\", "_").replace(":", "_")

ROOT_DIR = os.path.join(BASE_DIR, ROOT_FOLDER_NAME)

# Pasta com carimbo de data/hora (historico permanente).
RUN_DIR = os.path.join(ROOT_DIR, model_safe, PLUGIN_NAME, run_stamp)
if not os.path.exists(RUN_DIR):
    os.makedirs(RUN_DIR)

# Pasta "latest" (sempre sobrescrita).
LATEST_DIR = os.path.join(ROOT_DIR, model_safe, "latest", PLUGIN_NAME)
if not os.path.exists(LATEST_DIR):
    os.makedirs(LATEST_DIR)

# Caminhos de saida
out_cats        = os.path.join(RUN_DIR, "categorias.csv")
out_cats_lat    = os.path.join(LATEST_DIR, "categorias.csv")
out_fingerprint = os.path.join(RUN_DIR, "model_fingerprint.json")
out_fp_lat      = os.path.join(LATEST_DIR, "model_fingerprint.json")


# ============================================================
# FUNCOES AUXILIARES
# ============================================================

def u8(x):
    """
    Garante que o valor seja uma string unicode pura.
    Em IronPython 2, strings da API do Revit podem chegar como bytes (str)
    com encoding UTF-8 ou Latin-1. Retornar unicode evita mojibake no painel
    HTML do pyRevit (ex: 'LuminÃ¡rias' -> 'Luminárias').
    """
    if x is None:
        return u""
    if isinstance(x, unicode):
        return x
    if isinstance(x, (str, bytes)):
        try:
            return x.decode("utf-8")
        except (UnicodeDecodeError, AttributeError):
            try:
                return x.decode("latin-1")
            except (UnicodeDecodeError, AttributeError):
                return u""
    try:
        return unicode(x)
    except:
        return u""


def write_bom_csv(path):
    """Abre CSV em binario com BOM UTF-8 para compatibilidade com Excel."""
    f = open(path, "wb")
    f.write(u"\ufeff".encode("utf-8"))
    return f


def safe_str(x):
    """Retorna unicode string ou '' se None.
    Em IronPython, strings da API do Revit podem chegar como bytes (str),
    causando UnicodeDecodeError no json.dumps. Esta funcao garante unicode.
    """
    if not x:
        return u""
    if isinstance(x, unicode):
        return x
    # byte string: tenta decodificar como utf-8, depois latin-1
    try:
        return x.decode("utf-8")
    except (UnicodeDecodeError, AttributeError):
        try:
            return x.decode("latin-1")
        except (UnicodeDecodeError, AttributeError):
            return unicode(x)


def to_meters(internal_value):
    """
    Converte um valor de comprimento das unidades internas do Revit (pes decimais)
    para metros. Util para exibir elevacoes de niveis de forma legivel.
    A API do Revit armazena distancias em pes internos (feet),
    independente das configuracas de unidade do usuario.
    """
    try:
        return round(
            UnitUtils.ConvertFromInternalUnits(
                internal_value, UnitTypeId.Meters
            ), 3
        )
    except:
        return internal_value


# ============================================================
# INICIALIZA O PAINEL DE SAIDA RICO DO PYREVIT (NOVO)
# ============================================================
# script.get_output() retorna um objeto que renderiza HTML no painel
# lateral do pyRevit — muito mais rico que print() no console.
# Permite tabelas, Markdown, links clicaveis para elementos, etc.
output = script.get_output()
output.set_height(800)                        # Altura do painel em pixels
output.set_title("TestAPI — Diagnostico do Modelo")  # Titulo da aba

output.print_md("# Diagnostico do Modelo Revit")
output.print_md("**Modelo:** `{}`  \n**Execucao:** `{}`".format(model, run_stamp))


# ============================================================
# 1. INFORMACOES DO PROJETO (ProjectInfo — NOVO)
# ============================================================
# doc.ProjectInformation retorna o objeto ProjectInfo com metadados
# definidos em Revit > Gerenciar > Informacoes do Projeto.
output.print_md("---\n## 1. Informacoes do Projeto")

proj = doc.ProjectInformation

proj_fields = OrderedDict([
    ("Nome do Projeto",  safe_str(proj.Name)),
    ("Numero",           safe_str(proj.Number)),
    ("Cliente",          safe_str(proj.ClientName)),
    ("Endereco",         safe_str(proj.Address)),
    ("Status",           safe_str(proj.Status)),
    ("Autor",            safe_str(proj.Author)),
    ("Data da Emissao",  safe_str(proj.IssueDate)),
])

proj_table = [[k, v] for k, v in proj_fields.items()]
output.print_table(proj_table, columns=["Campo", "Valor"])


# ============================================================
# 2. VERSAO DO REVIT (Application — NOVO)
# ============================================================
# __revit__.Application expoe informacoes do executavel do Revit,
# nao do documento. Util para logs de compatibilidade.
output.print_md("---\n## 2. Versao do Revit")

app_info = OrderedDict([
    ("Versao",    str(app.VersionNumber)),
    ("Nome",      str(app.VersionName)),
    ("Build",     str(app.VersionBuild)),
    ("Idioma",    str(app.Language)),
])

app_table = [[k, v] for k, v in app_info.items()]
output.print_table(app_table, columns=["Campo", "Valor"])


# ============================================================
# 3. NIVEIS / PAVIMENTOS (Level — NOVO)
# ============================================================
# Level representa pavimentos do edificio (Terreo, 1o Andar, Cobertura...).
# A elevacao e armazenada em pes internos e convertida para metros aqui.
output.print_md("---\n## 3. Niveis / Pavimentos")

levels = sorted(
    FilteredElementCollector(doc).OfClass(Level).ToElements(),
    key=lambda lv: lv.Elevation  # Ordena do mais baixo para o mais alto
)

levels_table = []
for lv in levels:
    elev_m = to_meters(lv.Elevation)
    levels_table.append([
        u8(lv.Name),
        "{} m".format(elev_m),
        u8(lv.Id.ToString()),
    ])

output.print_table(levels_table, columns=["Nome", "Elevacao", "ID"])
output.print_md("_Total: {} niveis_".format(len(levels)))


# ============================================================
# 4. WORKSETS (FilteredWorksetCollector — NOVO)
# ============================================================
# Worksets sao conjuntos de trabalho usados em modelos COLABORATIVOS
# (arquivo central / modelos locais). Se o modelo nao e workshared,
# essa coleta retorna lista vazia.
output.print_md("---\n## 4. Worksets (Conjuntos de Trabalho)")

worksets_info = []
try:
    if doc.IsWorkshared:
        wsets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
        for ws in wsets:
            # Conta quantos elementos pertencem a este workset.
            # WorksetId e usado como filtro.
            ws_elements = FilteredElementCollector(doc) \
                .WhereElementIsNotElementType() \
                .ToElements()
            count = sum(
                1 for el in ws_elements
                if el.WorksetId and el.WorksetId.IntegerValue == ws.Id.IntegerValue
            )
            worksets_info.append([
                u8(ws.Name),
                u8(ws.Owner) if ws.Owner else "(aberto)",
                str(count),
                "Aberto" if ws.IsOpen else "Fechado",
                u8(str(ws.Id)),
            ])
        output.print_table(worksets_info, columns=["Workset", "Proprietario", "Elementos", "Estado", "ID"])
    else:
        output.print_md("_Modelo nao e workshared (sem worksets de usuario)._")
except Exception as e:
    output.print_md("_Erro ao ler worksets: {}_".format(str(e)))


# ============================================================
# 5. ALERTAS DO MODELO (doc.GetWarnings() — NOVO)
# ============================================================
# O Revit registra inconsistencias autodetectadas como "Warnings" (alertas).
# Exemplos: paredes sobrepostas, elementos duplicados, rooms nao fechados.
# GetWarnings() retorna todos os alertas ativos no modelo.
output.print_md("---\n## 5. Alertas do Modelo (Warnings)")

warnings = []
try:
    warnings = list(doc.GetWarnings())
except:
    pass

if warnings:
    warn_counter = Counter()
    for w in warnings:
        try:
            desc = str(w.GetDescriptionText())
            # Trunca descricoes longas para caber na tabela.
            warn_counter[desc[:80]] += 1
        except:
            warn_counter["(erro ao ler descricao)"] += 1

    warn_table = [[desc, str(cnt)] for desc, cnt in warn_counter.most_common(15)]
    output.print_table(warn_table, columns=["Descricao (top 15)", "Ocorrencias"])
    output.print_md("_Total de alertas: {}_".format(len(warnings)))
else:
    output.print_md("_Nenhum alerta encontrado no modelo._")


# ============================================================
# 6. MODELOS LINKADOS (RevitLinkInstance — NOVO)
# ============================================================
# Um modelo Revit pode referenciar outros arquivos .rvt via "links".
# Por exemplo: estrutura + arquitetura + instalacoes como arquivos separados.
# RevitLinkInstance representa cada instancia de link no modelo hospedeiro.
output.print_md("---\n## 6. Modelos Linkados (RVT Links)")

links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()

if links:
    links_table = []
    for lk in links:
        try:
            lk_type = doc.GetElement(lk.GetTypeId())
            status  = str(lk_type.GetLinkedFileStatus()) if lk_type else "N/D"
            name    = u8(lk.Name)
            links_table.append([name, status, u8(lk.Id.ToString())])
        except:
            links_table.append([u8(lk.Name), "Erro", u8(lk.Id.ToString())])
    output.print_table(links_table, columns=["Nome do Link", "Status", "ID"])
else:
    output.print_md("_Nenhum modelo linkado encontrado._")


# ============================================================
# 7. FASES (doc.Phases — NOVO)
# ============================================================
# Fases representam etapas cronologicas do projeto: demolicao, nova construcao, etc.
# Cada elemento do Revit pode ser associado a uma fase de criacao/demolicao.
output.print_md("---\n## 7. Fases do Projeto")

phases = doc.Phases  # Retorna uma colecao de Phase objects
phases_table = []
for i, ph in enumerate(phases):
    phases_table.append([str(i + 1), u8(ph.Name), u8(ph.Id.ToString())])

output.print_table(phases_table, columns=["Ordem", "Nome", "ID"])


# ============================================================
# 8. OPCOES DE PROJETO (DesignOption — NOVO)
# ============================================================
# Design Options permitem modelar variantes de layout no mesmo arquivo.
# Exemplo: duas opcoes de escada (escada A / escada B) coexistindo no modelo.
output.print_md("---\n## 8. Opcoes de Projeto (Design Options)")

d_opts = FilteredElementCollector(doc).OfClass(DesignOption).ToElements()

if d_opts:
    opts_table = [[u8(d.Name), u8(d.Id.ToString())] for d in d_opts]
    output.print_table(opts_table, columns=["Opcao", "ID"])
else:
    output.print_md("_Nenhuma Design Option encontrada._")


# ============================================================
# 9. CATEGORIAS DE ELEMENTOS (Counter — ja conhecido, mas com linkify NOVO)
# ============================================================
# Esta secao usa output.linkify() — gera um link clicavel no painel HTML
# que, ao ser clicado, seleciona o elemento no Revit.
output.print_md("---\n## 9. Top {} Categorias de Elementos".format(TOP_N_CATS))

all_elements = list(
    FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
)

cats_map = {}  # cat_name -> list de element Ids (para poder usar linkify)
for el in all_elements:
    try:
        if el.Category:
            cname = el.Category.Name
            if cname not in cats_map:
                cats_map[cname] = []
            cats_map[cname].append(el.Id)
    except:
        pass

counter    = Counter({k: len(v) for k, v in cats_map.items()})
cats_table = []
for cat, qty in counter.most_common(TOP_N_CATS):
    # output.linkify(element_id) retorna um link HTML que seleciona o
    # elemento no Revit ao ser clicado no painel de saida.
    # Aqui usamos o primeiro elemento da categoria para ilustrar linkify.
    sample_id   = cats_map[cat][0] if cats_map[cat] else None
    sample_link = output.linkify(sample_id) if sample_id else "—"
    cats_table.append([u8(cat), str(qty), sample_link])

output.print_table(cats_table, columns=["Categoria", "Quantidade", "Exemplo (clique)"])
output.print_md("_Total de categorias unicas: {}  |  Total de elementos: {}_".format(
    len(counter), len(all_elements)
))


# ============================================================
# 10. CSV — Ranking de categorias (dual-save)
# ============================================================
f     = write_bom_csv(out_cats)
f_lat = write_bom_csv(out_cats_lat)
w     = csv.writer(f)
w_lat = csv.writer(f_lat)

HEADER = ["run_stamp", "model", "categoria", "quantidade"]
w.writerow(HEADER)
w_lat.writerow(HEADER)

for cat, qty in counter.most_common():
    # Codifica para UTF-8 bytes pois o CSV e aberto em modo binario
    row = [u8(run_stamp).encode("utf-8"), u8(model).encode("utf-8"),
           u8(cat).encode("utf-8"), str(qty)]
    w.writerow(row)
    w_lat.writerow(row)

f.close()
f_lat.close()


# ============================================================
# 11. JSON FINGERPRINT — Relatorio multi-secao (dual-save)
# ============================================================
# O "fingerprint" registra o estado do modelo num dado momento.
# Salvar fingerprints sequenciais permite detectar alteracoes entre versoes.

def normalize_for_json(obj):
    """
    Percorre recursivamente um objeto e garante que todas as strings
    sejam unicode puro — necessario para contornar bug do json do IronPython 2.7
    que nao consegue serializar byte strings com chars nao-ASCII.
    """
    if isinstance(obj, dict):
        return {normalize_for_json(k): normalize_for_json(v) for k, v in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [normalize_for_json(v) for v in obj]
    elif isinstance(obj, unicode):
        return obj  # ja e unicode, sem tratamento adicional
    elif isinstance(obj, (str, bytes)):
        try:
            return obj.decode("utf-8")
        except (UnicodeDecodeError, AttributeError):
            try:
                return obj.decode("latin-1")
            except (UnicodeDecodeError, AttributeError):
                return u""
    else:
        return obj


def safe_phases_list(phases_obj):
    result = []
    try:
        for ph in phases_obj:
            result.append(safe_str(ph.Name))
    except:
        pass
    return result

fingerprint = {
    "run_stamp":   run_stamp,
    "model":       safe_str(model),
    "revit_version": {
        "number": safe_str(str(app.VersionNumber)),
        "name":   safe_str(str(app.VersionName)),
        "build":  safe_str(str(app.VersionBuild)),
    },
    "project_info": {
        "name":    safe_str(proj.Name),
        "number":  safe_str(proj.Number),
        "client":  safe_str(proj.ClientName),
        "address": safe_str(proj.Address),
        "status":  safe_str(proj.Status),
        "author":  safe_str(proj.Author),
    },
    "stats": {
        "total_elements":         len(all_elements),
        "unique_categories":      len(counter),
        "levels_count":           len(levels),
        "links_count":            len(list(links)),
        "phases_count":           len(list(phases)),
        "warnings_count":         len(warnings),
        "design_options_count":   len(list(d_opts)),
        "is_workshared":          bool(doc.IsWorkshared),
        "worksets_count":         len(worksets_info),
    },
    "levels": [
        {"name": safe_str(lv.Name), "elevation_m": to_meters(lv.Elevation)}
        for lv in levels
    ],
    "phases": safe_phases_list(phases),
    "top_categories": [
        {"category": safe_str(cat), "count": qty}
        for cat, qty in counter.most_common(TOP_N_CATS)
    ],
}

for path in [out_fingerprint, out_fp_lat]:
    try:
        # Normaliza todo o dict para unicode puro antes de serializar.
        # Contorna bug do json do IronPython 2.7 com byte strings nao-ASCII.
        safe_fp = normalize_for_json(fingerprint)
        json_bytes = json.dumps(safe_fp, indent=2, ensure_ascii=True)
        if isinstance(json_bytes, unicode):
            json_bytes = json_bytes.encode("utf-8")
        with open(path, "wb") as jf:
            jf.write(json_bytes)
    except Exception as ex_json:
        with open(path, "wb") as jf:
            jf.write(u"{{\"error\": \"{}\"}}".format(str(ex_json)).encode("utf-8"))


# ============================================================
# OUTPUT FINAL — console padrao (complementar ao painel HTML)
# ============================================================
print("=" * 55)
print("  DIAGNOSTICO DO MODELO CONCLUIDO")
print("=" * 55)
print("  Modelo  : {}".format(model))
print("  Revit   : {} ({})".format(app.VersionNumber, app.VersionName))
print("  Data    : {}".format(day))
print("  Hora    : {}".format(ts))
print("")
print("  [Historico]")
print("    {}".format(RUN_DIR))
print("    - categorias.csv")
print("    - model_fingerprint.json")
print("")
print("  [Ultima execucao]")
print("    {}".format(LATEST_DIR))
print("")
print("-" * 55)
print("  TOTAIS")
print("-" * 55)
print("  Elementos       : {}".format(len(all_elements)))
print("  Categorias      : {}".format(len(counter)))
print("  Niveis          : {}".format(len(levels)))
print("  Links RVT       : {}".format(len(list(links))))
print("  Fases           : {}".format(len(list(phases))))
print("  Alertas         : {}".format(len(warnings)))
print("  Design Options  : {}".format(len(list(d_opts))))
print("  Workshared      : {}".format("Sim" if doc.IsWorkshared else "Nao"))
print("=" * 55)
print("PROCESSO REALIZADO COM SUCESSO")
