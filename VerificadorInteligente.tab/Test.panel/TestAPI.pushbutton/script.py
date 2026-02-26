# -*- coding: utf-8 -*-

from pyrevit import revit
from Autodesk.Revit.DB import FilteredElementCollector
from collections import Counter

doc = revit.doc

elements = list(
    FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
)

cats = []
for el in elements:
    try:
        if el.Category:
            cats.append(el.Category.Name)
    except:
        pass

counter = Counter(cats)

print("Modelo: {}".format(doc.Title))
print("Total de elementos: {}".format(len(elements)))
print("Total de categorias unicas: {}".format(len(counter)))
print("\nTop 20 categorias:")

for cat, qtd in counter.most_common(20):
    print("{}: {}".format(cat, qtd))