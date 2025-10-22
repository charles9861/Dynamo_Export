# Exporting Data to Excel using Custom Python Node

Before diving in, lets restate our goal to ensure we’re on the same page:

### 🎯 Objective

To:

1. **Extract all available part sizes** from a **Civil 3D Pipe Network**,
   not just the ones already in use.
2. Use that list to **create a pick list** (so you can select a replacement size per pipe).
3. Eventually, **batch-swap parts** (e.g., replace multiple pipes with different sizes at once).

Right now, you’re stuck at step 1 — **getting all possible part sizes** from the Civil 3D parts catalog or the current parts list.

---

### 🔍 Key Clarification

Could you confirm **which source** you want to pull the part sizes from?

There are two possibilities:

1. **Civil 3D Parts Catalog** (the full library — e.g. XML-based structure in `C:\ProgramData\Autodesk\C3D <version>\enu\Pipes Catalog\US Imperial Pipes\...`)

   * This contains *every* available size/type combination.
   * More complex to extract directly — would require reading the catalog XMLs.

2. **The Current Drawing’s Parts List** (the one associated with your network)

   * This is much more accessible via Dynamo.
   * It contains only parts that are available *for use in that drawing*, not the full catalog.

If your goal is to make a practical **size picker for swapping parts**, the **Parts List** route is more manageable.

---

### 🧠 Likely Dynamo Limitation

By default, **Dynamo for Civil 3D** doesn’t expose nodes that list all part sizes from a catalog or parts list directly — you usually get:

* Pipe/Structure objects in the model (with current size info), or
* Limited access to parts list definitions.

You might need to use:

* **Python nodes** accessing the **Autodesk.Civil.DatabaseServices** API.
* Specifically, classes like:

  * `Autodesk.Civil.DatabaseServices.PartsList`
  * `Autodesk.Civil.DatabaseServices.PartSize`
  * `Autodesk.Civil.DatabaseServices.PartFamily`
  * `Autodesk.Civil.DatabaseServices.PipeNetwork`

---

### ✅ What I Can Help You With Next

Here’s what I can do for you now:

1. **Review your Dynamo script** and identify how you’re currently trying to export the pipe data.
2. **Provide a Python node snippet** you can add to your Dynamo script that:

   * Reads the **Parts List** from Civil 3D.
   * Extracts **all part sizes** for each pipe family.
   * Outputs a list that you can export to Excel or use as a pick list in Dynamo.

---

# Part 2

* **Inspect your `.dyn` script** and explain what’s missing / where to plug in the Python node,
  or
* **Jump straight to generating a working Python node** that lists all pipe sizes from the parts list?

(Once you pick, I’ll open and analyze the `.dyn` file or generate the Python code.)

Perfect — that’s exactly the right approach 👌
We’ll start simple and build it up piece by piece so it’s easy to troubleshoot and expand later.

Here’s what I’ll do next:

1. **Open your Dynamo script** (`PipeNetworkExportToExcel_V2.dyn`) to see your current structure — how data flows, where we can slot in a Python node cleanly, and what outputs you already have.
2. Then I’ll write a **Python node** that:

   * Runs inside Dynamo for Civil 3D,
   * Reads the **current Parts List** attached to your active drawing,
   * Extracts **all pipe sizes** (diameters, names, etc.),
   * Outputs them as a list of strings or dictionaries (so you can later export or link them to dropdowns).

Once that works, we can expand it to include things like material, structure type, or even automate the part swapping.

---

Let’s start by inspecting your `.dyn` file.
Give me a moment to read it and see where we’ll integrate the Python node.


Your Dynamo script looks well-developed — it currently has:

* **25 Function nodes** (custom or built-in operations)
* **16 Number input nodes**
* **5 String input nodes**
* **1 Extension node** (likely for Excel or Civil 3D integration)
* **58 connectors** linking everything

This suggests you’re already exporting or processing pipe data systematically (probably through Civil 3D > Excel).

---

### 🧩 Next Step — Add the Custom Python Node

We’ll insert a **Python Script node** that:

* Queries the **active Civil 3D document**,
* Gets the **current Parts List**,
* Extracts **pipe part sizes**,
* Returns them as a list (ready for Excel export or pick list use).

Here’s a minimal version to start with:

```python
# Python Script for Dynamo (Civil 3D)
# Get available pipe part sizes from the current parts list

import clr
clr.AddReference('AcMgd')
clr.AddReference('AcDbMgd')
clr.AddReference('AeccDbMgd')

from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.Civil.ApplicationServices import CivilApplication
from Autodesk.Civil.DatabaseServices import *

doc = Application.DocumentManager.MdiActiveDocument
civilDoc = CivilApplication.ActiveDocument

output = []

try:
    # Loop through all Parts Lists in the current drawing
    for partsListId in civilDoc.Styles.PartLists:
        partsList = doc.Database.TransactionManager.GetObject(partsListId, 0)
        name = partsList.Name
        pipe_families = []

        # Get each part family (e.g. "Concrete Pipe", "HDPE", etc.)
        for familyId in partsList.GetPartFamilies(PartType.Pipe):
            family = doc.Database.TransactionManager.GetObject(familyId, 0)
            family_name = family.Name
            sizes = []

            # Get all part sizes in this family
            for size in family.PartSizes:
                sizes.append(size.Name)

            pipe_families.append({
                "Family": family_name,
                "Sizes": sizes
            })

        output.append({
            "PartsList": name,
            "PipeFamilies": pipe_families
        })

except Exception as e:
    output = str(e)

OUT = output
```

---

### 🧠 How to Use This

1. In Dynamo, add a **Python Script** node.
2. Copy–paste the code above into it.
3. Run it inside a Civil 3D environment (not Sandbox mode).
4. The output should be a nested list/dictionary of:

   * Parts list name → Pipe family → Size names.

Example output:

```
[
  {
    "PartsList": "Storm Sewer",
    "PipeFamilies": [
      {"Family": "Concrete Pipe", "Sizes": ["300mm", "375mm", "450mm"]},
      {"Family": "HDPE Pipe", "Sizes": ["150mm", "200mm"]}
    ]
  }
]
```

---

Would you like me to:

* Modify this script to **output a simple flat list** (just all sizes regardless of family),
  or
* Keep it **grouped by family and parts list** (better for dropdown pickers)?

