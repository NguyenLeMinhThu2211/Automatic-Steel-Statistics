using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in1.clRenewBar))]

namespace AutoCAD_CSharp_plug_in1
{
    internal class clRenewBar
    {
        [CommandMethod("GanThepdoc", CommandFlags.Modal)]
        public void Renew_bar()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityResult per;
            PromptEntityOptions peo1 = new PromptEntityOptions("\nChọn thanh thép: ");
            per = doc.Editor.GetEntity(peo1);
            if (per.Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //string ObjNameText = GenerateRandomName(4);
                    string SignText = "";
                    string Diameter = "";
                    double DiaMeter_double = 0;
                    string Quantity = "";
                    double Quantity_double = 0;
                    System.Drawing.Font fontName = new System.Drawing.Font("Arial", 1);
                    #region Lấy đường kính và số lượng thép
                    // Yêu cầu người dùng chọn một đối tượng Text hoặc MText
                    PromptEntityOptions peoText = new PromptEntityOptions("\nChọn Tag thép: ");
                    peoText.SetRejectMessage("\nChỉ có thể chọn Text hoặc MText.");
                    peoText.AddAllowedClass(typeof(DBText), false);
                    peoText.AddAllowedClass(typeof(MText), false);
                    PromptEntityResult perText = doc.Editor.GetEntity(peoText);

                    if (perText.Status == PromptStatus.OK)
                    {
                        Entity ent = tr.GetObject(perText.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent is DBText)
                        {
                            DBText acText = ent as DBText;
                            string pattern = @"(\d+)∅";
                            Match match1 = Regex.Match(acText.TextString, pattern);
                            Quantity = match1.Groups[1].Value;
                            Quantity_double = Convert.ToDouble(Quantity, CultureInfo.InvariantCulture);
                            // Sử dụng regex để trích xuất số sau "%%C"
                            string pattern2 = @"∅(\d+)";
                            Match match2 = Regex.Match(acText.TextString, pattern2);
                            Diameter = match2.Groups[1].Value;
                            DiaMeter_double = Convert.ToDouble(Diameter, CultureInfo.InvariantCulture);
                            //doc.Editor.WriteMessage("\nNội dung Text: " + match2.Groups[1].Value);
                        }
                        else if (ent is MText)
                        {
                            MText acMText = ent as MText;
                            string pattern = @"(\d+)∅";
                            Match match1 = Regex.Match(acMText.Text, pattern);
                            Quantity = match1.Groups[1].Value;
                            Quantity_double = Convert.ToDouble(Quantity, CultureInfo.InvariantCulture);
                            //doc.Editor.WriteMessage("\nNội dung MText: " + match1.Groups[1].Value);

                            string pattern2 = @"∅(\d+)";
                            Match match2 = Regex.Match(acMText.Text, pattern2);
                            Diameter = match2.Groups[1].Value;
                            DiaMeter_double = Convert.ToDouble(Diameter, CultureInfo.InvariantCulture);
                            //doc.Editor.WriteMessage("\nNội dung Text: " + match2.Groups[1].Value);
                        }
                    }
                    #endregion
                    #region Ký hiệu thép
                    // Yêu cầu người dùng chọn một đối tượng Text hoặc MText
                    PromptEntityOptions peoSignBar = new PromptEntityOptions("\nChọn Ký hiệu thép: ");
                    peoSignBar.SetRejectMessage("\nChỉ có thể chọn Text hoặc MText.");
                    peoSignBar.AddAllowedClass(typeof(DBText), false);
                    peoSignBar.AddAllowedClass(typeof(MText), false);
                    PromptEntityResult perSignBar = doc.Editor.GetEntity(peoSignBar);

                    if (perSignBar.Status == PromptStatus.OK)
                    {
                        Entity ent = tr.GetObject(perSignBar.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent is DBText)
                        {
                            DBText acText = ent as DBText;
                            SignText = acText.TextString;
                        }
                        else if (ent is MText)
                        {
                            MText acMText = ent as MText;
                            SignText = acMText.Text;
                            //doc.Editor.WriteMessage("\nNội dung Text: " + match2.Groups[1].Value);
                        }
                    }
                    #endregion
                    
                    AutoCAD_CSharp_plug_in1.library.CreateLayer.CreateAndAssignALayer();
                    //BlockTableRecord BtrAtt = new BlockTableRecord();
                    CreateNewBarLine createPlineBlk = new CreateNewBarLine();
                    createPlineBlk.CreateBlockBar(db, tr, per, SignText, Quantity, Diameter);
                    tr.Commit();
                }
            }

        }
        
    }
}
