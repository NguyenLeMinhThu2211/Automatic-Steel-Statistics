using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;

using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in1.MainModel))]
namespace AutoCAD_CSharp_plug_in1
{
    internal class MainModel
    {
        [CommandMethod("QTK2", CommandFlags.Modal)]
        
        public void BangTKT()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Yêu cầu người dùng chọn một đường thẳng
            //PromptEntityOptions peo = new PromptEntityOptions("\nChọn đối tượng cần thống kê: ");

            //peo.SetRejectMessage("\nĐối tượng này chưa được gán.");
            //peo.AddAllowedClass(typeof(BlockReference), false);
            //PromptEntityResult per = ed.GetEntity(peo);


            Point3d pt;
            //if (per.Status == PromptStatus.OK)
            //{
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                string ObjNameText = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(4);
                string SignText = "";
                string Diameter = "";
                double DiaMeter_double = 0 ;
                double lengthdouble = 0;
                string Quantity = "";
                double Quantity_double = 0;
                System.Drawing.Font fontName = new System.Drawing.Font("Arial",1) ;
                PromptEntityResult per;
#region Tạo Khung tiêu đề

                BlockTable bt;
                bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectId bID = ObjectId.Null;
                using (BlockTableRecord Btr = new BlockTableRecord())
                {
                    // Tạo một tên ngẫu nhiên cho block
                    string randomBlockName = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(8); // 8 là độ dài của tên
                                                                    // Kiểm tra xem tên đã tồn tại chưa
                    while (bt.Has(randomBlockName))
                    {
                        randomBlockName = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(8); // Sinh tên mới nếu tên đã tồn tại
                    }
                    Btr.Name = randomBlockName;

                    #region Lấy chiều dài
                    PromptEntityOptions peo1 = new PromptEntityOptions("\nChọn thanh thép: ");
                    per = doc.Editor.GetEntity(peo1);
                    if (per.Status == PromptStatus.OK)
                    {
                        Entity entity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                        // Kiểm tra xem đối tượng có thuộc tính Length không
                        PropertyInfo propInfo = entity.GetType().GetProperty("Length");

                        if (propInfo != null)
                        {
                            lengthdouble = (double)propInfo.GetValue(entity);
                            //doc.Editor.WriteMessage($"\nChiều dài của đối tượng là: {length}");
                        }
                        else
                        {
                            doc.Editor.WriteMessage("\nĐối tượng chọn không có thuộc tính chiều dài.");
                        }
                        
                    }
                    #endregion
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
                    #region Tên Cấu Kiện
                    // Yêu cầu người dùng chọn một đối tượng Text hoặc MText
                    PromptEntityOptions peoNameObj = new PromptEntityOptions("\nChọn Tên đối tượng: ");
                    peoNameObj.SetRejectMessage("\nChỉ có thể chọn Text hoặc MText.");
                    peoNameObj.AddAllowedClass(typeof(DBText), false);
                    peoNameObj.AddAllowedClass(typeof(MText), false);
                    PromptEntityResult perNameObj = doc.Editor.GetEntity(peoNameObj);
                    

                    if (perNameObj.Status == PromptStatus.OK)
                    {
                        Entity ent = tr.GetObject(perNameObj.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent is DBText)
                        {
                            DBText acText = ent as DBText;
                           ObjNameText = acText.TextString;
                            string pattern = @"%%U";
                            string cleanedText = Regex.Replace(ObjNameText, pattern, "");
                            ObjNameText= cleanedText;
                           TextStyleTableRecord textStyle = tr.GetObject(acText.TextStyleId, OpenMode.ForRead) as TextStyleTableRecord;
                            fontName = new System.Drawing.Font(textStyle.Font.ToString(),1);
                        }
                        else if (ent is MText)
                        {
                            MText acMText = ent as MText;
                            ObjNameText = acMText.Text;
                            TextStyleTableRecord textStyle = tr.GetObject(acMText.TextStyleId, OpenMode.ForRead) as TextStyleTableRecord;
                            fontName = new System.Drawing.Font(textStyle.Font.ToString(), 1);
                            //doc.Editor.WriteMessage("\nNội dung Text: " + match2.Groups[1].Value);
                        }
                        
                        
                    }
                    #endregion
                    #region Chọn điểm chèn block
                    PromptPointOptions ppo = new PromptPointOptions("\nChọn điểm chèn bảng thống kê trên bản vẽ: ");
                    PromptPointResult ppr = ed.GetPoint(ppo);
                    pt = ppr.Value;
                    Btr.Origin = pt;
                    #endregion
                    #region Vẽ bảng và chèn vào BLOCK

                    Polyline acPoly = new Polyline();

                    acPoly.AddVertexAt(0, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    acPoly.AddVertexAt(1, new Point2d(pt.X + 6000, pt.Y), 0, 0, 0);
                    acPoly.AddVertexAt(2, new Point2d(pt.X + 6000, pt.Y - 300), 0, 0, 0);
                    acPoly.AddVertexAt(3, new Point2d(pt.X, pt.Y - 300), 0, 0, 0);
                    acPoly.AddVertexAt(4, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    Btr.AppendEntity(acPoly);

                    Line line1 = new Line(new Point3d(pt.X, pt.Y - 300, 0), new Point3d(pt.X, pt.Y - 800, 0));
                    line1.Layer = "0";
                    Btr.AppendEntity(line1);
                    Line line2 = new Line(new Point3d(pt.X + 500, pt.Y - 300, 0), new Point3d(pt.X + 500, pt.Y - 800, 0));
                    line2.Layer = "0";
                    Btr.AppendEntity(line2);
                    Line line3 = new Line(new Point3d(pt.X + 1000, pt.Y - 300, 0), new Point3d(pt.X + 1000, pt.Y - 800, 0));
                    line3.Layer = "0";
                    Btr.AppendEntity(line3);
                    Line line4 = new Line(new Point3d(pt.X + 2500, pt.Y - 300, 0), new Point3d(pt.X + 2500, pt.Y - 800, 0));
                    line4.Layer = "0";
                    Btr.AppendEntity(line4);
                    Line line5 = new Line(new Point3d(pt.X + 3000, pt.Y - 300, 0), new Point3d(pt.X + 3000, pt.Y - 800, 0));
                    line5.Layer = "0";
                    Btr.AppendEntity(line5);
                    Line line6 = new Line(new Point3d(pt.X + 3500, pt.Y - 300, 0), new Point3d(pt.X + 3500, pt.Y - 800, 0));
                    line6.Layer = "0";
                    Btr.AppendEntity(line6);
                    Line line7 = new Line(new Point3d(pt.X + 4000, pt.Y - 300, 0), new Point3d(pt.X + 4000, pt.Y - 800, 0));
                    line7.Layer = "0";
                    Btr.AppendEntity(line7);
                    Line line8 = new Line(new Point3d(pt.X + 4500, pt.Y - 550, 0), new Point3d(pt.X + 4500, pt.Y - 800, 0));
                    line8.Layer = "0";
                    Btr.AppendEntity(line8);
                    Line line8a = new Line(new Point3d(pt.X + 4000, pt.Y - 550, 0), new Point3d(pt.X + 5000, pt.Y - 550, 0));
                    line8a.Layer = "0";
                    Btr.AppendEntity(line8a);
                    Line line9 = new Line(new Point3d(pt.X + 5000, pt.Y - 300, 0), new Point3d(pt.X + 5000, pt.Y - 800, 0));
                    line9.Layer = "0";
                    Btr.AppendEntity(line9);
                    Line line10 = new Line(new Point3d(pt.X + 5500, pt.Y - 300, 0), new Point3d(pt.X + 5500, pt.Y - 800, 0));
                    line10.Layer = "0";
                    Btr.AppendEntity(line10);
                    Line line11 = new Line(new Point3d(pt.X + 6000, pt.Y - 300, 0), new Point3d(pt.X + 6000, pt.Y - 800, 0));
                    line11.Layer = "0";
                    Btr.AppendEntity(line11);
                    Line line12 = new Line(new Point3d(pt.X, pt.Y - 800, 0), new Point3d(pt.X + 6000, pt.Y - 800, 0));
                    line12.Layer = "0";
                    Btr.AppendEntity(line12);
                    #endregion
                    #region Thêm các đề mục thống kê thép
                    MText ContentMText = new MText();
                    ContentMText.Location = new Point3d(pt.X + 3000, pt.Y - 150, 0);
                    ContentMText.Contents = "BẢNG THỐNG KÊ THÉP";
                    ContentMText.TextHeight = 75;
                    ContentMText.Layer = "0";
                    ContentMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    ContentMText.Width = 2000;
                    Btr.AppendEntity(ContentMText);

                    MText TenCKMText = new MText();
                    TenCKMText.Location = new Point3d(pt.X + 250, pt.Y - 550, 0);
                    TenCKMText.Contents = "Tên CK";
                    TenCKMText.TextHeight = 75;
                    TenCKMText.Layer = "0";
                    TenCKMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    TenCKMText.Width = 500;
                    Btr.AppendEntity(TenCKMText);

                    MText SHMText = new MText();
                    SHMText.Location = new Point3d(pt.X + 750, pt.Y - 550, 0);
                    SHMText.Contents = "Số hiệu";
                    SHMText.TextHeight = 75;
                    SHMText.Layer = "0";
                    SHMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    SHMText.Width = 500;
                    Btr.AppendEntity(SHMText);

                    MText ShapeMText = new MText();
                    ShapeMText.Location = new Point3d(pt.X + 1750, pt.Y - 550, 0);
                    ShapeMText.Contents = "Hình Dạng Thép";
                    ShapeMText.TextHeight = 75;
                    ShapeMText.Layer = "0";
                    ShapeMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    ShapeMText.Width = 1500;
                    Btr.AppendEntity(ShapeMText);

                    MText DiaMText = new MText();
                    DiaMText.Location = new Point3d(pt.X + 2750, pt.Y - 550, 0);
                    DiaMText.Contents = "Đường kính thép";
                    DiaMText.TextHeight = 75;
                    DiaMText.Layer = "0";
                    DiaMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    DiaMText.Width = 500;
                    Btr.AppendEntity(DiaMText);

                    MText LongMText = new MText();
                    LongMText.Location = new Point3d(pt.X + 3250, pt.Y - 550, 0);
                    LongMText.Contents = "Chiều dài thanh thép (mm)";
                    LongMText.TextHeight = 75;
                    LongMText.Layer = "0";
                    LongMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    LongMText.Width = 500;
                    Btr.AppendEntity(LongMText);

                    MText NumberMText = new MText();
                    NumberMText.Location = new Point3d(pt.X + 3750, pt.Y - 550, 0);
                    NumberMText.Contents = "Số CK";
                    NumberMText.TextHeight = 75;
                    NumberMText.Layer = "0";
                    NumberMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    NumberMText.Width = 500;
                    Btr.AppendEntity(NumberMText);

                    MText QuantityMText = new MText();
                    QuantityMText.Location = new Point3d(pt.X + 4500, pt.Y - 425, 0);
                    QuantityMText.Contents = "Số lượng";
                    QuantityMText.TextHeight = 75;
                    QuantityMText.Layer = "0";
                    QuantityMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    QuantityMText.Width = 1000;
                    Btr.AppendEntity(QuantityMText);

                    MText Quan1CKMText = new MText();
                    Quan1CKMText.Location = new Point3d(pt.X + 4250, pt.Y - 675, 0);
                    Quan1CKMText.Contents = "1CK";
                    Quan1CKMText.TextHeight = 75;
                    Quan1CKMText.Layer = "0";
                    Quan1CKMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    Quan1CKMText.Width = 500;
                    Btr.AppendEntity(Quan1CKMText);

                    MText QuanFullMText = new MText();
                    QuanFullMText.Location = new Point3d(pt.X + 4750, pt.Y - 675, 0);
                    QuanFullMText.Contents = "Toàn bộ";
                    QuanFullMText.TextHeight = 75;
                    QuanFullMText.Layer = "0";
                    QuanFullMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    QuanFullMText.Width = 500;
                    Btr.AppendEntity(QuanFullMText);

                    MText SumOfLengthMtext = new MText();
                    SumOfLengthMtext.Location = new Point3d(pt.X + 5250, pt.Y - 550, 0);
                    SumOfLengthMtext.Contents = "Tổng chiều dài (m)";
                    SumOfLengthMtext.TextHeight = 75;
                    SumOfLengthMtext.Layer = "0";
                    SumOfLengthMtext.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    SumOfLengthMtext.Width = 500;
                    Btr.AppendEntity(SumOfLengthMtext);

                    MText SumOfMassMtext = new MText();
                    SumOfMassMtext.Location = new Point3d(pt.X + 5750, pt.Y - 550, 0);
                    SumOfMassMtext.Contents = "Tổng khối lượng (kG)";
                    SumOfMassMtext.TextHeight = 75;
                    SumOfMassMtext.Layer = "0";
                    SumOfMassMtext.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    SumOfMassMtext.Width = 500;
                    Btr.AppendEntity(SumOfMassMtext);
                    #endregion
                    
                    tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                    bt.Add(Btr);
                    tr.AddNewlyCreatedDBObject(Btr, true);
                    bID = Btr.Id;
                }
               
                #region Chèn Block vào Workspace
                if (bID != ObjectId.Null)
                {
                    using (BlockReference Br = new BlockReference(pt, bID)) //Set location to place block
                    {
                        BlockTableRecord Btr;
                        Btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Btr.AppendEntity(Br);
                        tr.AddNewlyCreatedDBObject(Br, true);
                    }
                }
                #endregion
                #endregion
#region Tạo Block Attribute để lưu dữ liệu thép
                #region Tạo Các thuộc tính (Attribute)
                double scale = 25;
                PromptStringOptions SLCK_PSO = new PromptStringOptions("\nNhập số lượng cấu kiện: ");
                SLCK_PSO.AllowSpaces = true;
                double SLCK=0;
                PromptResult SLCK_PR = ed.GetString(SLCK_PSO);
                SLCK = Convert.ToDouble(SLCK_PR.StringResult, CultureInfo.InvariantCulture);
                if (SLCK == 0) { SLCK = 1;}

                AttributeDefinition attObjName = new AttributeDefinition();
                attObjName.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
                attObjName.Tag = "Tên Ck";
                attObjName.Prompt = "";
                attObjName.Height = 2 * scale;
                //attObjName.Justify = AttachmentPoint.MiddleCenter;

                AttributeDefinition SH = new AttributeDefinition();
                SH.Position = new Point3d(pt.X + 750, pt.Y - 950, 0);
                SH.Tag = "Ký Hiệu Thanh Thép";
                SH.Prompt = "";
                SH.Height =2* scale;

                AttributeDefinition attDiameter = new AttributeDefinition();
                attDiameter.Position = new Point3d(pt.X + 2750, pt.Y - 950, 0);
                attDiameter.Tag = "Đường kính (mm)";
                attDiameter.Prompt = "Enter the value:";
                attDiameter.Height = 2 * scale;

                AttributeDefinition attObjSLCK = new AttributeDefinition();
                attObjSLCK.Position = new Point3d(pt.X + 3750, pt.Y - 950, 0);
                attObjSLCK.Tag = "Số Lượng cấu kiện";
                attObjSLCK.Prompt = "";
                attObjSLCK.Height =2 * scale;

                AttributeDefinition attQuantity = new AttributeDefinition();
                attQuantity.Position = new Point3d(pt.X +4250, pt.Y - 950, pt.Z);
                attQuantity.Tag = "Số lượng trong 1 CK";
                attQuantity.Prompt = "Enter the value:";
                attQuantity.Height = 2 * scale;

                AttributeDefinition attQuantitySUM = new AttributeDefinition();
                attQuantitySUM.Position = new Point3d(pt.X + 4750, pt.Y - 950, pt.Z);
                attQuantitySUM.Tag = "Số lượng trong 1 CK";
                attQuantitySUM.Prompt = "Enter the value:";
                attQuantitySUM.Height = 2* scale;

                AttributeDefinition attBarLength = new AttributeDefinition();
                attBarLength.Position = new Point3d(pt.X+3250, pt.Y-950, pt.Z);
                attBarLength.Tag = "Chiều dài (mm)";
                attBarLength.Prompt = "";
                attBarLength.Height = 2* scale;

                AttributeDefinition attBarSUMLength = new AttributeDefinition();
                attBarSUMLength.Position = new Point3d(pt.X + 5250, pt.Y - 950, pt.Z);
                attBarSUMLength.Tag = "Tổng Chiều dài (mm)";
                attBarSUMLength.Prompt = "";
                attBarSUMLength.Height =2 * scale;

                AttributeDefinition attBarSUMMass = new AttributeDefinition();
                attBarSUMMass.Position = new Point3d(pt.X + 5750, pt.Y - 950, pt.Z);
                attBarSUMMass.Tag = "Tổng Khối lượng (kG)";
                attBarSUMMass.Prompt = "";
                attBarSUMMass.Height = 2* scale;
                #endregion
                #region Tạo tên ngẫu nhiên cho Block Attribute
                string btAttName = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(8); 
                while (bt.Has(btAttName))
                {
                    btAttName = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(8);
                }
                #endregion
                #region Gán các dữ liệu đối tượng vào BlockTable
                BlockTable btAttribute;
                btAttribute = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectId bIDAttribute = ObjectId.Null;
                using (BlockTableRecord BtrAtt = new BlockTableRecord())
                {
                    BtrAtt.Name = btAttName;
                    // set reference location for Block
                    BtrAtt.Origin = pt;
                    // Add the Attribute to the block
                    BtrAtt.AppendEntity(attObjName);
                    BtrAtt.AppendEntity(SH);
                    BtrAtt.AppendEntity(attDiameter);
                    BtrAtt.AppendEntity(attObjSLCK);
                    BtrAtt.AppendEntity(attQuantity);
                    BtrAtt.AppendEntity(attQuantitySUM);
                    BtrAtt.AppendEntity(attBarLength);
                    BtrAtt.AppendEntity(attBarSUMLength);
                    BtrAtt.AppendEntity(attBarSUMMass);
                    // Add Block table reference to Block table
                    tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                    btAttribute.Add(BtrAtt);
                    tr.AddNewlyCreatedDBObject(BtrAtt, true);
                    //Add Shape Object and Line Object
                    ResizePolyline(db,tr, BtrAtt, btAttribute, per, new Point3d(pt.X + 1750, pt.Y - 950, 0));
                    CreateNewBarLine createPolyL = new CreateNewBarLine();
                    //createPolyL.CreateBlcPLine(db, tr, per);

                    InsertRowTable(tr, BtrAtt, new Point3d(pt.X, pt.Y - 800, 0));
                    bID = BtrAtt.Id;
                }
                #endregion
                #region Gán các thuộc tính vào BlockTableRecord
                if (bID != ObjectId.Null)
                {
                    using (BlockReference Br = new BlockReference(new Point3d(pt.X, pt.Y, 0), bID)) //Set location to place block
                    {
                        BlockTableRecord Btr;
                        Btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        //Set Block to BlockReference
                        Btr.AppendEntity(Br);
                        tr.AddNewlyCreatedDBObject(Br, true);
                        //Set Attribute for Block Attribute
                        SetAttToBlock(tr, Br, attObjName, ObjNameText, scale,true);
                        SetAttToBlock(tr, Br, SH, SignText, scale,false);
                        SetAttToBlock(tr, Br, attDiameter, Diameter, scale, false);
                        SetAttToBlock(tr, Br, attObjSLCK, SLCK.ToString(), scale, false);
                        SetAttToBlock(tr, Br, attQuantity, Quantity, scale, false);
                        SetAttToBlock(tr, Br, attQuantitySUM, (Quantity_double * SLCK).ToString(), scale, false);
                            //Round number
                        lengthdouble = Math.Round(lengthdouble / 5) * 5;
                        int lengthint = Convert.ToInt32(Math.Round(lengthdouble));
                        SetAttToBlock(tr, Br, attBarLength, lengthint.ToString(), scale, false);
                        SetAttToBlock(tr, Br, attBarSUMLength, (lengthint * Quantity_double * SLCK).ToString(), scale, false);
                            //Calculate Mass of Rebar
                        double Mass = (lengthint / 1000) * ((Math.PI * DiaMeter_double * DiaMeter_double) / (4 * 1000)) * Quantity_double * SLCK * 7850;
                        Mass = Math.Round(Mass,2);
                        //int Massint = Convert.ToInt32(Math.Round(Mass));
                        SetAttToBlock(tr, Br, attBarSUMMass, Mass.ToString(), scale, false);
                    }
                }
                //--------------------------------------------------------------------------------------
                #endregion
#endregion
               
                tr.Commit();

            }
        }

        
        public static Point3d SetCenterMidle(double scale, AttributeDefinition AttDef, AttributeReference AttRef,bool Rotation90)
        {
            System.Drawing.Font font = new System.Drawing.Font("Arial", 1);
            Size size = TextRenderer.MeasureText(AttRef.TextString, font);
            double sizeWidth = size.Width; double sizeHeight = size.Height;
            double textHeight = AttDef.Height;
            long longtext = AttRef.TextString.LongCount();
            double textWidth = sizeWidth * scale;
            Point3d position;
            if (Rotation90 == true)
            {
               position = new Point3d(AttDef.Position.X + textHeight / 2, AttDef.Position.Y - textWidth / (2), 0);
            }
            else
            {
                position = new Point3d(AttDef.Position.X - textWidth / (2), AttDef.Position.Y - textHeight / 2, 0);
            }    
            return position;
        }
        public void SetAttToBlock(Transaction tr, BlockReference Br, AttributeDefinition AttDef, string TextString, double scale, bool Rotation90)
        {
            AttributeReference attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(AttDef, Br.BlockTransform);
            attRef.TextString = TextString;
            Point3d position = SetCenterMidle(scale, AttDef, attRef, Rotation90);
            attRef.Position = position;
            Br.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
            if (Rotation90 == true)
            {
            attRef.Rotation = Math.PI / 2;
            }
        }
        public void ResizePolyline(Database db, Transaction tr,BlockTableRecord Btr, BlockTable Bt, PromptEntityResult per, Point3d Point)
        {
            // Assuming 'polylineId' is the ObjectId of your polyline
            //Polyline polyline = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Polyline;
            Polyline Sign_polyline = new Polyline();
            double SumX = 0; double SumY = 0; double X_Location = 0; ; double Y_Location = 0;
            long countPointX = 1; long countPointY = 1;
            Polyline polyline = new Polyline();
            //Chuyển line thành Polyline
            Entity entity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Entity;
            if (entity is Line line)
            {
                Point3dCollection stretchPoints = new Point3dCollection();
                entity.GetStretchPoints(stretchPoints);
                Point3d StartPoint = stretchPoints[0];
                Point3d EndPoint = stretchPoints[1];
                polyline.AddVertexAt(0, new Point2d(StartPoint.X, StartPoint.Y), 0, 0, 0);
                polyline.AddVertexAt(1, new Point2d(EndPoint.X, EndPoint.Y), 0, 0, 0);

            }
            else
            {
                polyline = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Polyline;
            }
            if (polyline != null)
            {
                Sign_polyline.CopyFrom(polyline);
                // Calculate the bounding box of the polyline
                Extents3d bounds = Sign_polyline.GeometricExtents;
                double currentXLength = bounds.MaxPoint.X - bounds.MinPoint.X;
                double currentYLength = bounds.MaxPoint.Y - bounds.MinPoint.Y;

                // Define the new dimensions
                double newXLength = currentXLength > 800 ? 800 : currentXLength;
                double newYLength = currentYLength > 50 ? 50 : currentYLength;

                // Calculate the scale factors
                double xScaleFactor = currentXLength == 0 ? 1.0 : newXLength / currentXLength;
                double yScaleFactor = currentYLength == 0 ? 1.0 : newYLength / currentYLength;

                // Resize the polyline by scaling each vertex
                for (int i = 0; i < Sign_polyline.NumberOfVertices; i++)
                {
                    Point2d oldPoint = polyline.GetPoint2dAt(i);
                    Point2d newPoint = new Point2d(oldPoint.X * xScaleFactor, oldPoint.Y * yScaleFactor);
                    Sign_polyline.SetPointAt(i, newPoint);
                    if (i == 0)
                    { 
                    SumX = SumX + newPoint.X;
                    SumY = SumY + newPoint.Y;
                    }
                    if (i != 0 && polyline.GetPoint2dAt(i).X != polyline.GetPoint2dAt(i - 1).X)
                    {
                        SumX = SumX + newPoint.X;
                        countPointX = countPointX + 1;
                    }
                    if (i != 0 && polyline.GetPoint2dAt(i).Y != polyline.GetPoint2dAt(i - 1).Y)
                    {
                        SumY = SumY + newPoint.Y;
                        countPointY = countPointY + 1;
                    }
                }
                X_Location = SumX / countPointX;
                Y_Location = SumY / countPointY;
                Point2d points = new Point2d(0, 0);
                points = Sign_polyline.GetPoint2dAt(0);
                Point3d P3D = new Point3d(X_Location, Y_Location, 0);
                Vector3d acVec3d = P3D.GetVectorTo(new Point3d(Point.X, Point.Y, 0));
                Matrix3d moveMatrix = Matrix3d.Displacement(acVec3d);
                Sign_polyline.TransformBy(moveMatrix);
                Btr.AppendEntity(Sign_polyline);
                tr.AddNewlyCreatedDBObject(Sign_polyline, true);
            }
        }

        public void InsertRowTable(Transaction tr, BlockTableRecord Btr, Point3d pt)
        {

            Line line1 = new Line(new Point3d(pt.X, pt.Y , 0), new Point3d(pt.X, pt.Y - 250, 0));
            line1.Layer = "0";
            Btr.AppendEntity(line1);
            tr.AddNewlyCreatedDBObject(line1, true);
            Line line2 = new Line(new Point3d(pt.X + 500, pt.Y, 0), new Point3d(pt.X + 500, pt.Y - 250, 0));
            line2.Layer = "0";
            Btr.AppendEntity(line2);
            tr.AddNewlyCreatedDBObject(line2, true);
            Line line3 = new Line(new Point3d(pt.X + 1000, pt.Y , 0), new Point3d(pt.X + 1000, pt.Y - 250, 0));
            line3.Layer = "0";
            Btr.AppendEntity(line3);
            tr.AddNewlyCreatedDBObject(line3, true);
            Line line4 = new Line(new Point3d(pt.X + 2500, pt.Y, 0), new Point3d(pt.X + 2500, pt.Y - 250, 0));
            line4.Layer = "0";
            Btr.AppendEntity(line4);
            tr.AddNewlyCreatedDBObject(line4, true);
            Line line5 = new Line(new Point3d(pt.X + 3000, pt.Y, 0), new Point3d(pt.X + 3000, pt.Y - 250, 0));
            line5.Layer = "0";
            Btr.AppendEntity(line5);
            tr.AddNewlyCreatedDBObject(line5, true);
            Line line6 = new Line(new Point3d(pt.X + 3500, pt.Y, 0), new Point3d(pt.X + 3500, pt.Y - 250, 0));
            line6.Layer = "0";
            Btr.AppendEntity(line6);
            tr.AddNewlyCreatedDBObject(line6, true);
            Line line7 = new Line(new Point3d(pt.X + 4000, pt.Y, 0), new Point3d(pt.X + 4000, pt.Y - 250, 0));
            line7.Layer = "0";
            Btr.AppendEntity(line7);
            tr.AddNewlyCreatedDBObject(line7, true);
            Line line8 = new Line(new Point3d(pt.X + 4500, pt.Y, 0), new Point3d(pt.X + 4500, pt.Y - 250, 0));
            line8.Layer = "0";
            Btr.AppendEntity(line8);
            tr.AddNewlyCreatedDBObject(line8, true);
            Line line9 = new Line(new Point3d(pt.X + 5000, pt.Y, 0), new Point3d(pt.X + 5000, pt.Y - 250, 0));
            line9.Layer = "0";
            Btr.AppendEntity(line9);
            tr.AddNewlyCreatedDBObject(line9, true);
            Line line10 = new Line(new Point3d(pt.X + 5500, pt.Y, 0), new Point3d(pt.X + 5500, pt.Y - 250, 0));
            line10.Layer = "0";
            Btr.AppendEntity(line10);
            tr.AddNewlyCreatedDBObject(line10, true);
            Line line11 = new Line(new Point3d(pt.X + 6000, pt.Y , 0), new Point3d(pt.X + 6000, pt.Y - 250, 0));
            line11.Layer = "0";
            Btr.AppendEntity(line11);
            tr.AddNewlyCreatedDBObject(line11, true);
            Line line12 = new Line(new Point3d(pt.X, pt.Y-250 ,0), new Point3d(pt.X + 6000, pt.Y - 250, 0));
            line12.Layer = "0";
            Btr.AppendEntity(line12);
            tr.AddNewlyCreatedDBObject(line12, true);

        }
    }
}
