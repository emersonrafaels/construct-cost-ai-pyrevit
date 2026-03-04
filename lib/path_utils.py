# -*- coding: utf-8 -*-
"""
path_utils.py
=============
Utilitarios de caminho e I/O compartilhados por todos os scripts do
VerificadorInteligente. Sem dependencias de Revit API — Python puro.

Exportacoes publicas:
  BASE_DIR, ROOT_FOLDER_NAME   -- constantes de pasta raiz
  make_run_stamp(now)          -- "YYYY-MM-DD_HHMMSS"
  make_day(now)                -- "YYYY-MM-DD"
  make_ts(now)                 -- "HHMMSS"
  make_model_safe(title)       -- sanitiza nome do modelo para nome de pasta
  build_run_dir(...)           -- pasta timestampada de uma execucao
  build_latest_dir(...)        -- pasta "latest" (sobrescrita a cada run)
  ensure_dirs(*paths)          -- mkdir -p para cada path dado
  u8(x)                        -- qualquer valor -> string CSV-safe
  safe_str(x)                  -- None -> ""
  write_bom_csv(path)          -- abre CSV com BOM UTF-8
  write_json_file(path, data)  -- salva dict como JSON com BOM UTF-8

Compativel com IronPython 2 / CPython 3
"""

import os
import csv
import json
import sys
import datetime

# ============================================================
# CONSTANTES GLOBAIS
# ============================================================
BASE_DIR         = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT_FOLDER_NAME = "RevitScan"

# ============================================================
# COMPAT Python 2 / 3
# ============================================================
try:
    _unicode = unicode   # IronPython 2
except NameError:
    _unicode = str       # CPython 3


# ============================================================
# TIMESTAMPS
# ============================================================

def make_run_stamp(now=None):
    """Retorna carimbo completo 'YYYY-MM-DD_HHMMSS'."""
    if now is None:
        now = datetime.datetime.now()
    return now.strftime("%Y-%m-%d_%H%M%S")


def make_day(now=None):
    """Retorna data 'YYYY-MM-DD'."""
    if now is None:
        now = datetime.datetime.now()
    return now.strftime("%Y-%m-%d")


def make_ts(now=None):
    """Retorna hora 'HHMMSS'."""
    if now is None:
        now = datetime.datetime.now()
    return now.strftime("%H%M%S")


# ============================================================
# MANIPULACAO DE NOMES DE MODELO
# ============================================================

def make_model_safe(title):
    """
    Sanitiza o titulo do modelo Revit para uso seguro em nomes de pasta.
    Remove / \\ : que sao invalidos no Windows.
    """
    return title.replace("/", "_").replace("\\", "_").replace(":", "_")


# ============================================================
# CONSTRUTORES DE CAMINHOS
# ============================================================

def _root_dir():
    return os.path.join(BASE_DIR, ROOT_FOLDER_NAME)


def build_run_dir(model_safe, plugin_name, run_stamp):
    """
    Constroi a pasta de historico permanente de uma execucao.
    Estrutura: Desktop/RevitScan/<model_safe>/<plugin_name>/<run_stamp>/
    """
    return os.path.join(_root_dir(), model_safe, plugin_name, run_stamp)


def build_latest_dir(model_safe, plugin_name):
    """
    Constroi a pasta 'latest', sobrescrita a cada execucao.
    Estrutura: Desktop/RevitScan/<model_safe>/latest/<plugin_name>/
    """
    return os.path.join(_root_dir(), model_safe, "latest", plugin_name)


def ensure_dirs(*paths):
    """Cria todas as pastas fornecidas (equivalente a mkdir -p)."""
    for p in paths:
        if not os.path.exists(p):
            os.makedirs(p)


# ============================================================
# CONVERSAO DE VALOR PARA STRING
# ============================================================

def u8(x):
    """
    Converte qualquer valor para string compativel com csv.writer.
    Funciona tanto em IronPython 2 (necessita bytes) quanto CPython 3.
    """
    try:
        if x is None:
            return ""
        if sys.version_info[0] >= 3:
            return str(x)
        if isinstance(x, _unicode):
            return x.encode("utf-8")
        return str(x)
    except:
        try:
            if sys.version_info[0] < 3:
                return _unicode(x).encode("utf-8")
            return str(x)
        except:
            return ""


def safe_str(x):
    """Retorna a string ou '' se None — evita 'None' literal nos CSVs."""
    return x if x else ""


# ============================================================
# I/O DE ARQUIVO
# ============================================================

def write_bom_csv(path):
    """
    Abre arquivo CSV para escrita com BOM UTF-8.
    O BOM faz o Excel reconhecer o encoding automaticamente.
    - CPython 3 : usa parametro encoding='utf-8-sig'
    - IronPython 2: modo binario com BOM manual
    """
    if sys.version_info[0] >= 3:
        f = open(path, "w", encoding="utf-8-sig", newline="")
    else:
        f = open(path, "wb")
        f.write(b"\xef\xbb\xbf")
    return f


def write_json_file(path, data):
    """
    Serializa um dict como JSON com indentacao e BOM UTF-8.
    Compativel com IronPython 2 (sem argumento encoding) e CPython 3.
    """
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    if sys.version_info[0] >= 3:
        with open(path, "w", encoding="utf-8-sig") as fj:
            fj.write(json_str)
    else:
        fj = open(path, "wb")
        fj.write(b"\xef\xbb\xbf")
        if isinstance(json_str, _unicode):
            fj.write(json_str.encode("utf-8"))
        else:
            fj.write(json_str)
        fj.close()
