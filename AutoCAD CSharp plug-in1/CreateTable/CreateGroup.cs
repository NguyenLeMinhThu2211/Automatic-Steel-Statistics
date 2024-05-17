using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_CSharp_plug_in1.RenewBar
{
    internal class CreateGroup
    {
       
    public static void AddEntitiesToGroup()
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Database acCurDb = acDoc.Database;

            
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
            // Create a polyline
            Polyline acPoly = new Polyline();
            acPoly.AddVertexAt(0, new Point2d(1, 1), 0, 0, 0);
            acPoly.AddVertexAt(1, new Point2d(4, 1), 0, 0, 0);

            // Create a text entity
            DBText acText = new DBText();
            acText.Position = new Point3d(1, 2, 0);
            acText.Height = 1;
            acText.TextString = "Sample Text";

            // Open the Block table record Model space for write
            BlockTableRecord acBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

            // Add the polyline and text to the block table record
            acBlkTblRec.AppendEntity(acPoly);
            acTrans.AddNewlyCreatedDBObject(acPoly, true);

            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);

            // Check if the group dictionary exists, if not create it
            DBDictionary grpDict = acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead) as DBDictionary;
                grpDict.UpgradeOpen();

                // Create a new unnamed group
                Group acGroup = new Group("", true);
                grpDict.SetAt(acGroup.Handle.ToString(), acGroup);
                acTrans.AddNewlyCreatedDBObject(acGroup, true);

                // Add entities to the group
                acGroup.Append(acPoly.ObjectId);
                acGroup.Append(acText.ObjectId);

                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }
    }
    public static void GetAttributesFromSelectedBlocks()
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Database acCurDb = acDoc.Database;
        Editor acEd = acDoc.Editor;

        // Prompt the user to select blocks
        PromptSelectionOptions pso = new PromptSelectionOptions();
        pso.MessageForAdding = "\nSelect block references: ";
        pso.AllowDuplicates = false;
        pso.SingleOnly = false;
        pso.SinglePickInSpace = false;

        // Set a filter to select only block references
        TypedValue[] acTypValAr = new TypedValue[1];
        acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 0);
        SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

        PromptSelectionResult acSSPrompt = acEd.GetSelection(pso, acSelFtr);

        if (acSSPrompt.Status == PromptStatus.OK)
        {
            SelectionSet acSSet = acSSPrompt.Value;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Iterate through the selected block references
                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        // Open the block reference for read
                        BlockReference acBlkRef = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;

                        // Check if the block reference has attributes
                        if (acBlkRef.AttributeCollection != null)
                        {
                            foreach (ObjectId attId in acBlkRef.AttributeCollection)
                            {
                                AttributeReference acAttRef = acTrans.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (acAttRef != null)
                                {
                                    // Output the tag and value of the attribute
                                    acEd.WriteMessage($"\nBlock: {acBlkRef.Name}, Tag: {acAttRef.Tag}, Value: {acAttRef.TextString}");
                                }
                            }
                        }
                    }
                }

                // Dispose the transaction
                acTrans.Commit();
            }
        }
        else
        {
            acEd.WriteMessage("\nNo blocks were selected.");
        }
    }


}
}
