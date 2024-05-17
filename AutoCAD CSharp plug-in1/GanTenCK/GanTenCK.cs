using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in1.GanTenCK))]
namespace AutoCAD_CSharp_plug_in1
{
    internal class GanTenCK
    {
        [CommandMethod("GanTenCK", CommandFlags.Modal)]
        public void Ganten_main()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            #region Chọn block thép 
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nChọn block thép dọc: ";
            pso.AllowDuplicates = false;
            pso.SingleOnly = false;
            pso.SinglePickInSpace = false;
            // Set a filter to select only block references
            TypedValue[] acTypValAr = new TypedValue[1];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 0);
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.GetSelection(pso, acSelFtr);
            #endregion

            #region Nhập tên cấu kiện
            PromptStringOptions SLCK_PSO = new PromptStringOptions("\nNhập tên cấu kiện: ");
            SLCK_PSO.AllowSpaces = true;
            string TenCK_string = "";
            PromptResult SLCK_PR = ed.GetString(SLCK_PSO);
            TenCK_string = SLCK_PR.StringResult;
            int bl_count = 0;

            #endregion
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
          
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        bl_count++;
                        if (acSSObj != null)
                        {
                            // Open the block reference for read
                            BlockReference acBlkRef = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                            if (acBlkRef.Layer == "TKT_thepchu")
                            {
                                if (acBlkRef.AttributeCollection != null)
                                {
                                    foreach (ObjectId attId in acBlkRef.AttributeCollection)
                                    {
                                        AttributeReference acAttRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                       
                                        if (acAttRef.Tag == "Tên CK")
                                        {

                                           acAttRef.TextString = TenCK_string;
                                           //acBlkRef.AttributeCollection.AppendAttribute(acAttRef);
                                           //tr.AddNewlyCreatedDBObject(acAttRef, true);

                                        }
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }

    }
}
