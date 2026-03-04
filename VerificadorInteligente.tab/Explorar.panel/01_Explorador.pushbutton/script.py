# -*- coding: utf-8 -*-

# ============================================================
# EXPLORADOR DE ELEMENTOS
# ============================================================
# Abre uma janela interativa (WPF) para navegar pela hierarquia
# completa de elementos do modelo Revit aberto:
#
#   Categorias → Famílias → Tipos → Instâncias
#
# Ao clicar em cada nível o painel seguinte é preenchido.
# Ao selecionar uma instância, seus detalhes aparecem no rodapé:
#   ID, Categoria, Família, Tipo, Nível, Ambiente (Room),
#   Fase Criação, Fase Demolição, Workset.
# ============================================================

from pyrevit import revit, forms
import datetime

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
)

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject
from System.IO import Path as SysPath

# ============================================================
# CONTEXTO DO MODELO
# ============================================================
doc   = revit.doc
app   = __revit__.Application
model = doc.Title
now   = datetime.datetime.now()
day   = now.strftime("%Y-%m-%d")

# ============================================================
# HELPERS
# ============================================================
def _safe(v):
    try:
        return unicode(v) if v else u""
    except NameError:
        return str(v) if v else u""

def _level_name(el):
    try:
        if el.LevelId and el.LevelId != ElementId.InvalidElementId:
            lv = el.Document.GetElement(el.LevelId)
            return lv.Name if lv else u""
    except:
        pass
    return u""

def _room_name(el):
    for attr in ("Room", "Space"):
        try:
            r = getattr(el, attr, None)
            if r:
                return _safe(r.Name)
        except:
            pass
    return u""

def _phase_names(el):
    try:
        pc = el.CreatedPhaseId
        pd = el.DemolishedPhaseId
        c = el.Document.GetElement(pc).Name if pc != ElementId.InvalidElementId else u""
        d = el.Document.GetElement(pd).Name if pd != ElementId.InvalidElementId else u""
        return c, d
    except:
        return u"", u""

def _workset_name(el):
    try:
        wt = el.Document.GetWorksetTable()
        ws = wt.GetWorkset(el.WorksetId)
        return ws.Name if ws else u""
    except:
        return u""

def _fam_type(el):
    fam = u""; typ = u""
    try:
        t = el.Document.GetElement(el.GetTypeId())
        if t:
            fam = _safe(getattr(t, "FamilyName", u""))
            typ = _safe(t.Name)
    except:
        pass
    return fam, typ

# ============================================================
# COLETA DE ELEMENTOS E MONTAGEM DA HIERARQUIA
# ============================================================
print("Coletando elementos...")
elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType())

hier        = {}   # {cat: {fam: {typ: [inst_dict]}}}
total_inst  = 0

for el in elements:
    try:
        cat = _safe(el.Category.Name) if el.Category else u"(sem categoria)"
        fam, typ = _fam_type(el)
        fam  = fam  or u"(sem familia)"
        typ  = typ  or u"(sem tipo)"
        lvl  = _level_name(el)     or u"(sem nivel)"
        room = _room_name(el)      or u"(sem ambiente)"
        pc, pd = _phase_names(el)
        ws = _workset_name(el)

        try:
            eid = str(int(el.Id.IntegerValue))
        except:
            try:
                eid = str(int(el.Id.Value))
            except:
                eid = _safe(el.Id)

        inst = {
            "id":      eid,
            "cat":     cat,
            "fam":     fam,
            "typ":     typ,
            "level":   lvl,
            "room":    room,
            "phase_c": pc,
            "phase_d": pd,
            "workset": ws,
        }
        hier.setdefault(cat, {}).setdefault(fam, {}).setdefault(typ, []).append(inst)
        total_inst += 1
    except:
        continue

print("  {} categorias | {} instancias".format(len(hier), total_inst))

# ============================================================
# JANELA WPF
# ============================================================
class ElementBrowserWindow(forms.WPFWindow):

    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)

        self.lblModelo.Text = u"Modelo: {}".format(model)
        self.lblInfo.Text   = u"Revit {}  |  {}  |  {} inst\u00e2ncias".format(
            app.VersionNumber, day, total_inst)
        self.badgeCat.Text  = u"{} cat.".format(len(hier))
        self.badgeFam.Text  = u"{} fam.".format(
            sum(len(f) for f in hier.values()))
        self.badgeInst.Text = u"{} inst.".format(total_inst)

        self._sel_cat   = None
        self._sel_fam   = None
        self._sel_typ   = None
        self._cat_keys  = []
        self._fam_keys  = []
        self._typ_keys  = []
        self._inst_list = []

        self.lstCategorias.SelectionChanged += self._on_cat_changed
        self.lstFamilias.SelectionChanged   += self._on_fam_changed
        self.lstTipos.SelectionChanged      += self._on_typ_changed
        self.lstInstancias.SelectionChanged += self._on_inst_changed
        self.btnLimpar.Click                += self._on_limpar

        self._fill_cats()
        self._update_breadcrumb()

    # ---- factories ----
    def _make_item(self, name, count_str):
        obj = ExpandoObject()
        obj.Name  = name
        obj.Count = count_str
        return obj

    def _make_inst_item(self, d):
        obj = ExpandoObject()
        obj.Name = u"ID: {}".format(d["id"])
        obj.Sub  = u"N\u00edvel: {}  |  Ambiente: {}".format(d["level"], d["room"])
        return obj

    def _set_source(self, listbox, items):
        col = ObservableCollection[object]()
        for it in items:
            col.Add(it)
        listbox.ItemsSource = col

    # ---- fill ----
    def _fill_cats(self):
        self._cat_keys = sorted(hier.keys())
        items = [
            self._make_item(c, u"{} inst\u00e2ncia(s)".format(
                sum(len(i) for f in hier[c].values() for i in f.values())))
            for c in self._cat_keys
        ]
        self._set_source(self.lstCategorias, items)
        self.cntCat.Text  = _safe(len(self._cat_keys))
        self.hintCat.Text = u"\u2190 clique para filtrar"

    def _fill_fams(self, cat):
        self._fam_keys = sorted(hier.get(cat, {}).keys())
        items = [
            self._make_item(f, u"{} inst\u00e2ncia(s)".format(
                sum(len(i) for i in hier[cat][f].values())))
            for f in self._fam_keys
        ]
        self._set_source(self.lstFamilias, items)
        self.cntFam.Text  = _safe(len(self._fam_keys))
        self.hintFam.Text = u"\u2190 clique para filtrar"

    def _fill_types(self, cat, fam):
        self._typ_keys = sorted(hier.get(cat, {}).get(fam, {}).keys())
        items = [
            self._make_item(t, u"{} inst\u00e2ncia(s)".format(
                len(hier[cat][fam][t])))
            for t in self._typ_keys
        ]
        self._set_source(self.lstTipos, items)
        self.cntTip.Text  = _safe(len(self._typ_keys))
        self.hintTip.Text = u"\u2190 clique para filtrar"

    def _fill_insts(self, cat, fam, typ):
        self._inst_list = hier.get(cat, {}).get(fam, {}).get(typ, [])
        self._set_source(self.lstInstancias,
                         [self._make_inst_item(d) for d in self._inst_list])
        self.cntInst.Text  = _safe(len(self._inst_list))
        self.hintInst.Text = u"\u2190 clique para detalhes"

    # ---- clear ----
    def _clear_panel(self, listbox, cnt, hint, msg):
        self._set_source(listbox, [])
        cnt.Text  = u"0"
        hint.Text = msg

    def _clear_details(self):
        for name in ("dId", "dCat", "dFam", "dTip",
                     "dLevel", "dRoom", "dPhaseC", "dPhaseD", "dWorkset"):
            getattr(self, name).Text = u"\u2014"

    # ---- show ----
    def _show_details(self, d):
        self.dId.Text      = d["id"]
        self.dCat.Text     = d["cat"]
        self.dFam.Text     = d["fam"]
        self.dTip.Text     = d["typ"]
        self.dLevel.Text   = d["level"]
        self.dRoom.Text    = d["room"]
        self.dPhaseC.Text  = d["phase_c"] or u"\u2014"
        self.dPhaseD.Text  = d["phase_d"] or u"\u2014"
        self.dWorkset.Text = d["workset"] or u"\u2014"

    def _update_breadcrumb(self):
        parts = [p for p in (self._sel_cat, self._sel_fam, self._sel_typ) if p]
        self.lblBreadcrumb.Text = (
            u"  \u203a  ".join(parts) if parts else u"Selecione uma categoria")

    # ---- events ----
    def _on_cat_changed(self, sender, args):
        idx = self.lstCategorias.SelectedIndex
        if idx < 0 or idx >= len(self._cat_keys): return
        self._sel_cat = self._cat_keys[idx]
        self._sel_fam = None
        self._sel_typ = None
        self._fill_fams(self._sel_cat)
        self._clear_panel(self.lstTipos,      self.cntTip,  self.hintTip,  u"Selecione uma fam\u00edlia")
        self._clear_panel(self.lstInstancias, self.cntInst, self.hintInst, u"Selecione um tipo")
        self._clear_details()
        self._update_breadcrumb()

    def _on_fam_changed(self, sender, args):
        idx = self.lstFamilias.SelectedIndex
        if idx < 0 or idx >= len(self._fam_keys): return
        self._sel_fam = self._fam_keys[idx]
        self._sel_typ = None
        self._fill_types(self._sel_cat, self._sel_fam)
        self._clear_panel(self.lstInstancias, self.cntInst, self.hintInst, u"Selecione um tipo")
        self._clear_details()
        self._update_breadcrumb()

    def _on_typ_changed(self, sender, args):
        idx = self.lstTipos.SelectedIndex
        if idx < 0 or idx >= len(self._typ_keys): return
        self._sel_typ = self._typ_keys[idx]
        self._fill_insts(self._sel_cat, self._sel_fam, self._sel_typ)
        self._clear_details()
        self._update_breadcrumb()

    def _on_inst_changed(self, sender, args):
        idx = self.lstInstancias.SelectedIndex
        if idx < 0 or idx >= len(self._inst_list): return
        self._show_details(self._inst_list[idx])

    def _on_limpar(self, sender, args):
        self._sel_cat = None
        self._sel_fam = None
        self._sel_typ = None
        self.lstCategorias.SelectedIndex = -1
        self._clear_panel(self.lstFamilias,   self.cntFam,  self.hintFam,  u"Selecione uma categoria")
        self._clear_panel(self.lstTipos,      self.cntTip,  self.hintTip,  u"Selecione uma fam\u00edlia")
        self._clear_panel(self.lstInstancias, self.cntInst, self.hintInst, u"Selecione um tipo")
        self._clear_details()
        self._update_breadcrumb()


# ============================================================
# LANCA A JANELA
# ============================================================
try:
    _xaml = SysPath.Combine(__file__.replace("script.py", ""), "ElementBrowserWindow.xaml")
    browser = ElementBrowserWindow(_xaml)
    browser.ShowDialog()
except Exception as err:
    forms.alert(
        u"Erro ao abrir o Explorador:\n\n{}".format(err),
        title=u"Explorador de Elementos",
        warn_icon=True,
    )