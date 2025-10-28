# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
__title__ = "3lly_RENAME_VIEW"
__doc__ = """Version = 1.0
Date    = 10.13.2025
_____________________________________________________________________
Description:
This is a template file for pyRevit Scripts.
_____________________________________________________________________
How-to:
-> Click on the button
-> ...
_____________________________________________________________________
Last update:
- [16.07.2024] - 1.1 Fixed an issue...
- [15.07.2024] - 1.0 RELEASE
- [10.13.2025] -ADDED A rename function
_____________________________________________________________________
To-Do:
- Describe Next Features
_____________________________________________________________________
Author: Erik Frits"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *

# pyRevit
from pyrevit import revit, forms

# .NET Imports (You often need List import)
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#==================================================
# 1-select views
# Get views - selected in projectbrowser

sel_el_ids = uidoc.Selection.GetElementIds()
sel_elem = [doc.GetElement(el_ids) for el_ids in sel_el_ids]
sel_views = [el for el in sel_elem if issubclass(type(el), View)]

# If None Selected - prompt selectviews from pyrevit.forms.select_views()

if not sel_views:
    sel_views = forms.select_views()

# ensure views not selected
if not sel_views:
    forms.alert("No Views Selected",exitscript=True)

print("DONE!")

# 2. define renaming rules
prefix = "Pre-"
find = "Floor plan"
replace = "Ef-level"
suffix = "-suf"

t=Transaction(doc,'py-rename views')
t.Start()

for view in sel_views:
    Old_name= view.Name
    new_name = prefix + Old_name.replace(find,replace) + suffix
    for i in range(20):
        try:
            view.Name = new_name
            print('{} ----{}'.format(Old_name, new_name))
            break

        except:
         new_name += '*'

t.Commit()




