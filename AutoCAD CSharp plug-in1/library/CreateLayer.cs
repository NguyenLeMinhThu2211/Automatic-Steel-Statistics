using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_CSharp_plug_in1.library
{
    internal class CreateLayer
    {

        public static void CreateAndAssignALayer()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable LayerTb = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                string sLayerThepchu = "TKT_thepchu";
                string sLayertag = "TKT_Tag_thepchu";
                string sLayerSHtag = "TKT_Tag_SH";
                AssignlayerToDb(acTrans, LayerTb, sLayerThepchu);
                AssignlayerToDb(acTrans, LayerTb, sLayertag);
                AssignlayerToDb(acTrans, LayerTb, sLayerSHtag);
                acTrans.Commit();
            }
        }
        public static void AssignlayerToDb(Transaction acTrans , LayerTable LayerTb, string sLayerName)
        {
            // Check if the layer already exists
            if (!LayerTb.Has(sLayerName))
            {
                using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                {
                    // Assign the layer a name and a color
                    acLyrTblRec.Name = sLayerName;
                    if (sLayerName == "TKT_thepchu")
                    {
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);//red
                    }
                    if (sLayerName == "TKT_Tag_thepchu")
                    {
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);//Yellow
                    }
                    if (sLayerName == "TKT_Tag_SH")
                    {
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);//Green
                    }
                    else 
                    {
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 7);//white
                    }
                        LayerTb.UpgradeOpen();
                    // Append the new layer to the Layer table and the transaction
                    LayerTb.Add(acLyrTblRec);
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                }
            }
        }
    }
}
