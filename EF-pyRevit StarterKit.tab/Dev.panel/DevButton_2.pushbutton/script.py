# -*- coding: utf-8 -*-
__title__ = "Constructive.WTC"
__doc__ = """Version = 3.1
Scans current + linked Revit models for MEP fixtures,
checks heights vs. standard values, and exports a UTF-8 CSV report."""

import clr, os, csv, codecs, sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System import Enum

# --------------------------------------------------
# 1️⃣ Safe Revit document getter
# --------------------------------------------------
def get_doc():
    try:
        return __revit__.ActiveUIDocument.Document
    except:
        try:
            return DocumentManager.Instance.CurrentDBDocument
        except:
            return None

doc = get_doc()
if not doc:
    from Autodesk.Revit.UI import TaskDialog
    TaskDialog.Show("pyRevit", "❌ No active Revit document found.\nPlease open a project and run again.")
    sys.exit()

print("🔍 Scanning current and linked models...\n")

# --------------------------------------------------
# 2️⃣ Table of standard heights (mm)
# --------------------------------------------------
standard_heights = {
    "Sink_Drain_Height": 900,
    "Sink_Supply_Height": 650,
    "Lavatory_Drain_Height": 500,
    "Lavatory_Supply_Height": 1100,
    "WM_Supply_Height": 500,
    "DW_Drain_Height": 350,
    "Gas_Outlet_Height": 80,
    "Gas_Heater_Supply_Height": 1100,
    "Electric_Heater_Supply_Height": 1400,
    "Heater_Pipe_Spacing": 100,
}

# --------------------------------------------------
# 3️⃣ Helper functions
# --------------------------------------------------
def mm(feet_val):
    try:
        return round(feet_val * 304.8, 1)
    except:
        return 0.0

def get_ifc_guid(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.IFC_GUID)
        return p.AsString() if p else "N/A"
    except:
        return "N/A"

def get_param_val(elem, name):
    try:
        p = elem.LookupParameter(name)
        if not p or not p.HasValue:
            return None
        if p.StorageType.ToString() == "Double":
            return mm(p.AsDouble())
        elif p.StorageType.ToString() == "String":
            return p.AsString()
        else:
            return p.AsValueString()
    except:
        return None

# --------------------------------------------------
# 4️⃣ Collect elements safely
# --------------------------------------------------
def collect_from_doc(document):
    categories = [
        "OST_PlumbingFixtures",
        "OST_PipeFitting",
        "OST_PipeAccessory",
        "OST_MechanicalEquipment",
        "OST_SpecialityEquipment"
    ]

    elems = []
    for c_name in categories:
        try:
            bic = Enum.Parse(BuiltInCategory, c_name)
            elems.extend(
                FilteredElementCollector(document)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception as ex:
            print("⚠️ Could not collect from category: {}".format(c_name))
    return elems

# --------------------------------------------------
# 5️⃣ Process all elements
# --------------------------------------------------
results = []

def process_elements(document, model_name):
    elements = collect_from_doc(document)
    for e in elements:
        if not e.Category:
            continue
        cat = e.Category.Name
        guid = get_ifc_guid(e)
        elname = e.Name or "Unnamed"

        level_param = e.LookupParameter("Reference Level") or e.LookupParameter("Base Level")
        if level_param:
            lvl = document.GetElement(level_param.AsElementId())
            level_name = lvl.Name
            level_elev = mm(lvl.Elevation)
        else:
            level_name = "N/A"
            level_elev = 0.0

        for param_name, std_value in standard_heights.items():
            val = get_param_val(e, param_name)
            if val is None:
                continue

            relative = val - level_elev
            deviation = relative - std_value
            results.append({
                "Model": model_name,
                "Category": cat,
                "Element Name": elname,
                "IFC_GUID": guid,
                "Parameter": param_name,
                "Level": level_name,
                "Value (mm)": round(relative, 1),
                "Standard (mm)": std_value,
                "Deviation (mm)": round(deviation, 1)
            })

# --------------------------------------------------
# 6️⃣ Process current + linked docs
# --------------------------------------------------
process_elements(doc, doc.Title)
for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
    try:
        ldoc = link.GetLinkDocument()
        if ldoc:
            lname = ldoc.Title or link.Name
            print(u"📎 Checking linked model: {}".format(lname))
            process_elements(ldoc, lname)
    except:
        continue

# --------------------------------------------------
# 7️⃣ Export to CSV safely (UTF-8)
# --------------------------------------------------
proj_path = doc.PathName or os.path.expanduser("~")
proj_dir = os.path.dirname(proj_path)
csv_path = os.path.join(proj_dir, "Revit_Heights_Report.csv")

if not results:
    print("⚠️ No matching elements or parameters found in current or linked models.")
else:
    with codecs.open(csv_path, "w", "utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "Model", "Category", "Element Name", "IFC_GUID",
            "Parameter", "Level", "Value (mm)", "Standard (mm)", "Deviation (mm)"
        ])
        writer.writeheader()
        for r in results:
            writer.writerow(r)

    print("\n📁 Report exported successfully to:\n{}".format(csv_path))

print("\n✅ Height check completed successfully.\n")

# --------------------------------------------------
# 8️⃣ Optional pyRevit message
# --------------------------------------------------
try:
    from Snippets._customprint import kit_button_clicked
    kit_button_clicked(btn_name=__title__)
except:
    pass
