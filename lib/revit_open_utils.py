# -*- coding: utf-8 -*-
"""
revit_open_utils.py
===================
Utilitarios para abertura e fechamento de arquivos .rvt em batch,
incluindo supressao automatica de dialogs e warnings do Revit.

Exportacoes publicas:
  DialogSuppressor          -- context manager que suprime dialogs/warnings
  open_rvt_detached(app, path)  -- abre .rvt desanexado do central
  close_doc(doc)                -- fecha documento sem salvar (seguro)

Uso tipico:
    from revit_open_utils import DialogSuppressor, open_rvt_detached, close_doc

    app = __revit__.Application
    with DialogSuppressor(app) as sup:
        for rvt in rvt_files:
            doc = open_rvt_detached(app, rvt)
            try:
                # extracao...
            finally:
                close_doc(doc)
        print(sup.log)   # lista de avisos suprimidos

Compativel com Revit 2022-2026  |  IronPython 2 / CPython 3
"""

from Autodesk.Revit.DB import (
    ModelPathUtils,
    OpenOptions,
    DetachFromCentralOption,
    WorksetConfiguration,
    WorksetConfigurationOption,
    FailureProcessingResult,
    FailureSeverity,
)


# ============================================================
# CONTEXT MANAGER — SUPRESSAO DE DIALOGS E WARNINGS
# ============================================================

class DialogSuppressor(object):
    """
    Suprime automaticamente dialogs nativos e warnings da API Revit
    durante operacoes em lote, evitando que modals travem o processo.

    Cobre dois canais:
      - app.DialogBoxShowing   : dialogs de UI (ex: familia desconectada)
      - app.FailuresProcessing : warnings de API  (ex: cota invalida)

    Uso como context manager:
        with DialogSuppressor(app) as sup:
            ...processar arquivos...
        n = len(sup.log)   # total de avisos suprimidos

    Uso manual (para compatibilidade com IronPython sem 'with'):
        sup = DialogSuppressor(app)
        sup.start()
        try:
            ...processar arquivos...
        finally:
            sup.stop()
    """

    def __init__(self, app):
        self._app = app
        self.log  = []   # lista de strings descrevendo cada supressao

    # -- handlers --------------------------------------------------

    def _on_dialog(self, sender, e):
        try:
            dialog_id = ""
            try:
                dialog_id = e.DialogId or ""
            except:
                pass
            self.log.append(u"[dialog] {}".format(dialog_id))
            e.OverrideResult(1)   # 1 = botao padrao (OK / primeira opcao)
        except:
            pass

    def _on_failures(self, sender, e):
        try:
            fa = e.GetFailuresAccessor()
            for msg in fa.GetFailureMessages():
                try:
                    if msg.GetSeverity() == FailureSeverity.Warning:
                        fa.DeleteWarning(msg)
                        self.log.append(
                            u"[warning] {}".format(msg.GetDescriptionText()))
                except:
                    pass
            e.SetProcessingResult(FailureProcessingResult.Continue)
        except:
            pass

    # -- ciclo de vida --------------------------------------------

    def start(self):
        """Registra os handlers nos eventos da Application."""
        # IronPython 2: usa bare except para capturar excecoes .NET
        # (System.MissingMemberException nao e capturada por 'except Exception')
        self._has_dialog_event = False
        if hasattr(self._app, "DialogBoxShowing"):
            try:
                self._app.DialogBoxShowing += self._on_dialog
                self._has_dialog_event = True
            except:
                pass
        self._app.FailuresProcessing += self._on_failures

    def stop(self):
        """Remove os handlers. Sempre chame em um bloco finally."""
        if getattr(self, "_has_dialog_event", False):
            try:
                self._app.DialogBoxShowing -= self._on_dialog
            except Exception:
                pass
        try:
            self._app.FailuresProcessing -= self._on_failures
        except Exception:
            pass

    # -- interface context manager --------------------------------

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
        return False   # propaga excecoes normalmente


# ============================================================
# ABERTURA DE ARQUIVO .rvt
# ============================================================

def open_rvt_detached(app, rvt_path):
    """
    Abre um arquivo .rvt desanexado do servidor central,
    com todos os worksets fechados (abertura rapida).

    Parametros:
      app      -- __revit__.Application
      rvt_path -- caminho absoluto para o arquivo .rvt

    Retorna o Document aberto ou levanta excecao em caso de falha.
    """
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rvt_path)

    opts = OpenOptions()
    opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets

    # Fecha todos os worksets para acelerar a abertura.
    # Em projetos sem worksets o SetOpenWorksetsConfiguration nao existe — ok.
    try:
        wsc = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        opts.SetOpenWorksetsConfiguration(wsc)
    except:
        pass

    return app.OpenDocumentFile(model_path, opts)


# ============================================================
# FECHAMENTO SEGURO DE DOCUMENTO
# ============================================================

def close_doc(doc):
    """
    Fecha o documento sem salvar.
    Engole qualquer excecao para nao interromper o loop do lote.
    """
    if doc is None:
        return
    try:
        doc.Close(False)
    except:
        pass
