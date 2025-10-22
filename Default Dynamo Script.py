# Load the Python Standard and DesignScript Libraries
import sys
import clr

# Add Assemblies for AutoCAD and Civil3D
clr.AddReference('AcMgd')
clr.AddReference('AcCoreMgd')
clr.AddReference('AcDbMgd')
clr.AddReference('AecBaseMgd')
clr.AddReference('AecPropDataMgd')
clr.AddReference('AeccDbMgd')

# Import references from AutoCAD
from Autodesk.AutoCAD.Runtime import *
from Autodesk.AutoCAD.ApplicationServices import *
from Autodesk.AutoCAD.EditorInput import *
from Autodesk.AutoCAD.DatabaseServices import *
from Autodesk.AutoCAD.Geometry import *

# Import references from Civil3D
from Autodesk.Civil.ApplicationServices import *
from Autodesk.Civil.DatabaseServices import *

# The inputs to this node will be stored as a list in the IN variables.
dataEnteringNode = IN

adoc = Application.DocumentManager.MdiActiveDocument
editor = adoc.Editor
civilDoc = CivilApplication.ActiveDocument

output = []

with adoc.LockDocument():
    db = adoc.Database
    tm = db.TransactionManager
    with tm.StartTransaction() as t:
        try:
            # Iterate through all Parts Lists in the current drawing
            for partsListId in civilDoc.Styles.PartLists:
                partsList = t.GetObject(partsListId, OpenMode.ForRead)
                list_name = partsList.Name
                pipe_families = []

                # Get each pipe family within the Parts List
                for familyId in partsList.GetPartFamilies(PartType.Pipe):
                    family = t.GetObject(familyId, OpenMode.ForRead)
                    fam_name = family.Name
                    sizes = []

                    # Get all part sizes in this family
                    for partSize in family.PartSizes:
                        sizes.append(partSize.Name)

                    pipe_families.append({
                        "Family": fam_name,
                        "Sizes": sizes
                    })

                output.append({
                    "PartsList": list_name,
                    "PipeFamilies": pipe_families
                })

        except Exception as e:
            output = str(e)

        # Commit before end transaction (read-only, but good practice)
        t.Commit()

# Assign your output to the OUT variable.
OUT = output
