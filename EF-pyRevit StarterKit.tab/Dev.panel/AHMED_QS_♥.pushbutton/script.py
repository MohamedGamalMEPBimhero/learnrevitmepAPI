# -*- coding: utf-8 -*-
__title__ = "Select by Parameter"
__doc__ = """Version = 1.1
Date    = 20.10.2025
________________________________________________
Description:
Select elements (e.g., pipes) by category and parameter value.

Example:
- Select all Pipes with Diameter = 100 mm
- Select all Walls with Type Name = 'Generic - 200mm'
________________________________________________
Author: Ahmed_QS ♥ | based on Erik Frits template
"""

# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
import clr

clr.AddReference('System')
from System.Collections.Generic import List

# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# ==================================================

# --------------------------------------------------
# Helper: Get parameter value safely
def get_param_value(elem, param_name):
    """Return a parameter's value as string or number."""
    param = elem.LookupParameter(param_name)
    if not param:
        return None
    if param.StorageType == StorageType.Double:
        return param.AsDouble()
    elif param.StorageType == StorageType.String:
        return param.AsString()
    elif param.StorageType == StorageType.ElementId:
        id_val = param.AsElementId()
        return id_val.IntegerValue
    elif param.StorageType == StorageType.Integer:
        return param.AsInteger()
    return None


# --------------------------------------------------
# Core: Filter elements by category + parameter
def select_elements_by_param(category, param_name, target_value, tolerance=0.001, active_view_only=True):
    """Select elements in Revit that match the given parameter value."""

    if active_view_only:
        collector = FilteredElementCollector(doc, doc.ActiveView.Id)
    else:
        collector = FilteredElementCollector(doc)

    elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()

    matched = []
    for elem in elements:
        val = get_param_value(elem, param_name)
        if val is None:
            continue

        # Compare values (handle doubles or strings)
        if isinstance(val, float):
            if abs(val - target_value) < tolerance:
                matched.append(elem)
        elif isinstance(val, str):
            if val.lower() == str(target_value).lower():
                matched.append(elem)
        else:
            if val == target_value:
                matched.append(elem)

    # Highlight elements in Revit
    if matched:
        uidoc.Selection.SetElementIds(List[ElementId]([e.Id for e in matched]))
        TaskDialog.Show("Result", "✅ Found and selected {0} elements matching criteria.".format(len(matched)))
    else:
        TaskDialog.Show("Result", "⚠️ No elements found with that parameter value.")


# ==================================================
# Example Usage (edit freely)
# ==================================================
# Example 1: Select all Pipes with Diameter = 100 mm
# Note: Revit stores diameters internally in feet.
mm_to_ft = 1 / 304.8
desired_diameter = 100 * mm_to_ft

select_elements_by_param(
    category=BuiltInCategory.OST_PipeCurves,
    param_name="Diameter",
    target_value=desired_diameter,
    tolerance=0.0001,
    active_view_only=True
)

# Example 2: (comment out the above, and try selecting walls by name)
# select_elements_by_param(
#     category=BuiltInCategory.OST_Walls,
#     param_name="Type Name",
#     target_value="Generic - 200mm",
#     active_view_only=False
# )

# ==================================================
# Keep default print message
# ==================================================
from Snippets._customprint import kit_button_clicked

kit_button_clicked(btn_name=__title__)

