using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using static System.Net.Mime.MediaTypeNames;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in1.QuetTKT.QuetTKT))]

namespace AutoCAD_CSharp_plug_in1.QuetTKT
{
    internal class QuetTKT
    {
        [CommandMethod("QuetTKT", CommandFlags.Modal)]
        public void BangTKT()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            #region Chọn block thép 
            // Prompt the user to select a block reference
            // Prompt the user to select blocks
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

            #region Chọn điểm chèn block
            PromptPointOptions ppo = new PromptPointOptions("\nChọn điểm chèn bảng thống kê trên bản vẽ: ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            Point3d pt = ppr.Value;
            #endregion

            #region Nhập số lượng cấu kiện
            PromptStringOptions SLCK_PSO = new PromptStringOptions("\nNhập số lượng cấu kiện: ");
            SLCK_PSO.AllowSpaces = true;
            double SLCK = 0;
            PromptResult SLCK_PR = ed.GetString(SLCK_PSO);
            SLCK = Convert.ToDouble(SLCK_PR.StringResult, CultureInfo.InvariantCulture);
            #endregion

            int i_block = 0;int i_blockGH = 0;
            string[] DS_SoHieu = new string[2];
            string[] DS_TenCauKien = new string[2];

            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;
                #region Đếm để tìm dòng cuối cùng
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            // Open the block reference for read
                            BlockReference acBlkRef = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                            if (acBlkRef.Layer == "TKT_thepchu")
                            {
                                i_blockGH++;
                            }
                        }
                    }
                    DS_TenCauKien = new string[i_blockGH];
                    DS_TenCauKien = LayDanhSachDL_Attribute(tr, acSSet, "Tên CK");
                    DS_SoHieu = new string[i_blockGH];
                    DS_SoHieu = LayDanhSachDL_Attribute(tr, acSSet, "Số hiệu");
                }
                #endregion

                foreach (string caukien in DS_TenCauKien)
                {
                    foreach (string SoHieu in DS_SoHieu)
                    {
                        foreach (SelectedObject acSSObj in acSSet)
                        {
                       
                            if (acSSObj != null)
                            {
                                using (Transaction tr = db.TransactionManager.StartTransaction())
                                {
                                    BlockReference acBlkRef = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                                    string Tenck_tam = LayDL_Attribute(tr, acBlkRef, "Tên CK");
                                    if (caukien == Tenck_tam)
                                    {
                                        if (acBlkRef.Layer == "TKT_thepchu")
                                        {
                                            string[] Data = LayCacThongSo(tr, acBlkRef);
                                            string ObjNameText = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(4);
                                            string SignText = "";
                                            string Diameter = "";
                                            double DiaMeter_double = 0;
                                            double lengthdouble = 0;
                                            string Quantity = "";
                                            double Quantity_double = 0;
                                            string A_str = ""; string B_str = ""; string C_str = "";
                                            double i_double = pt.Y - 950 - i_block * 250;
                                            lengthdouble = Convert.ToDouble(Data[4], CultureInfo.InvariantCulture);  //chiều dài thanh thép
                                            Quantity = Data[2]; // Số lượng thanh thép
                                            Quantity_double = Convert.ToDouble(Quantity, CultureInfo.InvariantCulture);
                                            Diameter = Data[3]; // đường kính thép
                                            DiaMeter_double = Convert.ToDouble(Diameter, CultureInfo.InvariantCulture);
                                            SignText = Data[1]; // ký hiệu thanh thép
                                            ObjNameText = Data[0]; //Tên cấu kiện
                                            A_str = Data[5]; B_str = Data[6]; C_str = Data[7];
                                            if (SignText == SoHieu)
                                            {
                                                #region Tạo Các thuộc tính (Attribute)
                                                double scale = 25;

                                                if (SLCK == 0) { SLCK = 1; }

                                                AttributeDefinition attObjName = new AttributeDefinition();
                                                attObjName.Position = new Point3d(pt.X + 250, i_double, 0);
                                                attObjName.Tag = "Tên Ck";
                                                attObjName.Prompt = "";
                                                attObjName.Height = 2 * scale;

                                                AttributeDefinition SH = new AttributeDefinition();
                                                SH.Position = new Point3d(pt.X + 750, i_double, 0);
                                                SH.Tag = "Ký Hiệu Thanh Thép";
                                                SH.Prompt = "";
                                                SH.Height = 2 * scale;

                                                AttributeDefinition attDiameter = new AttributeDefinition();
                                                attDiameter.Position = new Point3d(pt.X + 2750, i_double, 0);
                                                attDiameter.Tag = "Đường kính (mm)";
                                                attDiameter.Prompt = "Enter the value:";
                                                attDiameter.Height = 2 * scale;

                                                AttributeDefinition attObjSLCK = new AttributeDefinition();
                                                attObjSLCK.Position = new Point3d(pt.X + 3750, i_double, 0);
                                                attObjSLCK.Tag = "Số Lượng cấu kiện";
                                                attObjSLCK.Prompt = "";
                                                attObjSLCK.Height = 2 * scale;

                                                AttributeDefinition attQuantity = new AttributeDefinition();
                                                attQuantity.Position = new Point3d(pt.X + 4250, i_double, pt.Z);
                                                attQuantity.Tag = "Số lượng trong 1 CK";
                                                attQuantity.Prompt = "Enter the value:";
                                                attQuantity.Height = 2 * scale;

                                                AttributeDefinition attQuantitySUM = new AttributeDefinition();
                                                attQuantitySUM.Position = new Point3d(pt.X + 4750, i_double, pt.Z);
                                                attQuantitySUM.Tag = "Số lượng trong 1 CK";
                                                attQuantitySUM.Prompt = "Enter the value:";
                                                attQuantitySUM.Height = 2 * scale;

                                                AttributeDefinition attBarLength = new AttributeDefinition();
                                                attBarLength.Position = new Point3d(pt.X + 3250, i_double, pt.Z);
                                                attBarLength.Tag = "Chiều dài (mm)";
                                                attBarLength.Prompt = "";
                                                attBarLength.Height = 2 * scale;

                                                AttributeDefinition attBarSUMLength = new AttributeDefinition();
                                                attBarSUMLength.Position = new Point3d(pt.X + 5250, i_double, pt.Z);
                                                attBarSUMLength.Tag = "Tổng Chiều dài (mm)";
                                                attBarSUMLength.Prompt = "";
                                                attBarSUMLength.Height = 2 * scale;

                                                AttributeDefinition attBarSUMMass = new AttributeDefinition();
                                                attBarSUMMass.Position = new Point3d(pt.X + 5750, i_double, pt.Z);
                                                attBarSUMMass.Tag = "Tổng Khối lượng (kG)";
                                                attBarSUMMass.Prompt = "";
                                                attBarSUMMass.Height = 2 * scale;
                                                #endregion
                                                #region Khai báo các biến của Block
                                                BlockTable bt;
                                                bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                                ObjectId bID = ObjectId.Null;
                                                #endregion
                                                #region Tạo khung tiêu đề
                                                if (i_block == 0)
                                                {
                                                    TaoKhungtieude(tr, bt, db, pt);
                                                }
                                                #endregion
                                                #region Tạo tên ngẫu nhiên cho Block Attribute
                                                string btAttName = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(8);
                                                while (bt.Has(btAttName))
                                                {
                                                    btAttName = AutoCAD_CSharp_plug_in1.library.RandomName.GenerateRandomName(8);
                                                }
                                                #endregion
                                                #region Thêm Block vào bản vẽ CAD
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
                                                    //ResizePolyline(db, tr, BtrAtt, btAttribute, per, new Point3d(pt.X + 1750, pt.Y - 950, 0));
                                                    CreateNewBarLine createPolyL = new CreateNewBarLine();
                                                    //createPolyL.CreateBlcPLine(db, tr, per);
                                                    if (i_blockGH == i_block + 1)
                                                    {
                                                        InsertRowTableEnd(tr, BtrAtt, new Point3d(pt.X, pt.Y - 800 - i_block * 250, 0));
                                                    }
                                                    else
                                                    {
                                                        InsertRowTable(tr, BtrAtt, new Point3d(pt.X, pt.Y - 800 - i_block * 250, 0));
                                                    }
                                                    bID = BtrAtt.Id;
                                                    TaoHinhDangThep(tr, BtrAtt, db, new Point3d(pt.X + 1250, pt.Y - 1000 - 250 * i_block, 0), A_str, B_str, C_str);

                                                }
                                                #endregion
                                                #region Chèn block vào môi trường model Space
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
                                                        SetAttToBlock(tr, Br, attObjName, ObjNameText, scale, true);
                                                        SetAttToBlock(tr, Br, SH, SignText, scale, false);
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
                                                        double Mass = (lengthdouble / 1000) * ((Math.PI * DiaMeter_double * DiaMeter_double) / (4 * 1000)) * Quantity_double * SLCK * 7850;
                                                        Mass = Math.Round(Mass, 2);
                                                        //int Massint = Convert.ToInt32(Math.Round(Mass));
                                                        SetAttToBlock(tr, Br, attBarSUMMass, Mass.ToString(), scale, false);
                                                    }
                                                }
                                                //--------------------------------------------------------------------------------------
                                                #endregion
                                                i_block++;
                                            }
                                        }
                                        //acBlkRef.ResetBlock();
                                        tr.Commit();
                                    }
                                }
                            }
                        }
                        
                    }
                }
            }
        }
        public void InsertRowTable(Transaction tr, BlockTableRecord Btr, Point3d pt)
        {

            Line line1 = new Line(new Point3d(pt.X, pt.Y, 0), new Point3d(pt.X, pt.Y - 250, 0));
            line1.Layer = "0";
            Btr.AppendEntity(line1);
            tr.AddNewlyCreatedDBObject(line1, true);
            Line line2 = new Line(new Point3d(pt.X + 500, pt.Y, 0), new Point3d(pt.X + 500, pt.Y - 250, 0));
            line2.Layer = "0";
            Btr.AppendEntity(line2);
            tr.AddNewlyCreatedDBObject(line2, true);
            Line line3 = new Line(new Point3d(pt.X + 1000, pt.Y, 0), new Point3d(pt.X + 1000, pt.Y - 250, 0));
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
            Line line11 = new Line(new Point3d(pt.X + 6000, pt.Y, 0), new Point3d(pt.X + 6000, pt.Y - 250, 0));
            line11.Layer = "0";
            Btr.AppendEntity(line11);
            tr.AddNewlyCreatedDBObject(line11, true);
            Line line12 = new Line(new Point3d(pt.X+500, pt.Y - 250, 0), new Point3d(pt.X + 6000, pt.Y - 250, 0));
            line12.Layer = "0";
            Btr.AppendEntity(line12);
            tr.AddNewlyCreatedDBObject(line12, true);

        }

        public void InsertRowTableEnd(Transaction tr, BlockTableRecord Btr, Point3d pt)
        {

            Line line1 = new Line(new Point3d(pt.X, pt.Y, 0), new Point3d(pt.X, pt.Y - 250, 0));
            line1.Layer = "0";
            Btr.AppendEntity(line1);
            tr.AddNewlyCreatedDBObject(line1, true);
            Line line2 = new Line(new Point3d(pt.X + 500, pt.Y, 0), new Point3d(pt.X + 500, pt.Y - 250, 0));
            line2.Layer = "0";
            Btr.AppendEntity(line2);
            tr.AddNewlyCreatedDBObject(line2, true);
            Line line3 = new Line(new Point3d(pt.X + 1000, pt.Y, 0), new Point3d(pt.X + 1000, pt.Y - 250, 0));
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
            Line line11 = new Line(new Point3d(pt.X + 6000, pt.Y, 0), new Point3d(pt.X + 6000, pt.Y - 250, 0));
            line11.Layer = "0";
            Btr.AppendEntity(line11);
            tr.AddNewlyCreatedDBObject(line11, true);
            Line line12 = new Line(new Point3d(pt.X, pt.Y - 250, 0), new Point3d(pt.X + 6000, pt.Y - 250, 0));
            line12.Layer = "0";
            Btr.AppendEntity(line12);
            tr.AddNewlyCreatedDBObject(line12, true);

        }

        public void SetAttToBlock(Transaction tr, BlockReference Br, AttributeDefinition AttDef, string TextString, double scale, bool Rotation90)
        {
            AttributeReference attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(AttDef, Br.BlockTransform);
            attRef.TextString = TextString;
            Point3d position = SetCenterMidle(scale, AttDef, attRef, Rotation90);
            //Point3d position = AttDef.Position;
            attRef.Position = position;
            Br.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
            if (Rotation90 == true)
            {
                attRef.Rotation = Math.PI / 2;
            }
        }

        public static Point3d SetCenterMidle(double scale, AttributeDefinition AttDef, AttributeReference AttRef, bool Rotation90)
        {
            int longtext = AttRef.TextString.Length;
            double textHeight = AttDef.Height;
            double textWidth = longtext * textHeight * 38/50;
            Point3d position;
            if (Rotation90 == true)
            {
                position = new Point3d(AttDef.Position.X /*+ textHeight / 2*/, AttDef.Position.Y - textWidth / (2), 0);
            }
            else
            {
                position = new Point3d(AttDef.Position.X - textWidth / (2), AttDef.Position.Y/* - textHeight / 2*/, 0);
            }
            return position;
        }
        
        public void TaoKhungtieude(Transaction tr, BlockTable bt, Database db , Point3d pt)
        {
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
                Btr.Origin = pt;
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
                    // Tạo text tên bảng thống kê
                    MText ContentMText = new MText();
                    ContentMText.Location = new Point3d(pt.X + 3000, pt.Y - 150, 0);
                    ContentMText.Contents = "BẢNG THỐNG KÊ THÉP";
                    ContentMText.TextHeight = 75;
                    ContentMText.Layer = "0";
                    ContentMText.Attachment = AttachmentPoint.MiddleCenter; // Căn giữa MText
                    ContentMText.Width = 2000;
                    Btr.AppendEntity(ContentMText);
                    tr.AddNewlyCreatedDBObject(ContentMText, true);
                }
            }
            #endregion
        }

        public static string[] LayCacThongSo(Transaction tr, BlockReference acBlkRef)
        {
            string[] Dulieu = new string[8];
            int att_count = 0;

            if (acBlkRef.AttributeCollection != null)
            {
                foreach (ObjectId attId in acBlkRef.AttributeCollection)
                {
                    AttributeReference acAttRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (acAttRef != null)
                    {
                        
                        Dulieu[att_count] = acAttRef.TextString;
                        att_count++;
                    }
                }
            }
         return Dulieu;
        }

        public void TaoHinhDangThep(Transaction tr, BlockTableRecord Btr, Database db, Point3d pt, string A, string B, string C)
        {
            Polyline Shape_polyline = new Polyline();
            double A_doub = Convert.ToDouble(A, CultureInfo.InvariantCulture);
            double B_doub = Convert.ToDouble(B, CultureInfo.InvariantCulture);
            double C_doub = Convert.ToDouble(C, CultureInfo.InvariantCulture);
            if (A != "0")
            { 
            Shape_polyline.AddVertexAt(0, new Point2d(pt.X,pt.Y), 0, 0, 0);
            Shape_polyline.AddVertexAt(1, new Point2d(pt.X, pt.Y+100), 0, 0, 0);
            Shape_polyline.AddVertexAt(2, new Point2d(pt.X+1000, pt.Y + 100), 0, 0, 0);
                if (C_doub != 0)
                { Shape_polyline.AddVertexAt(3, new Point2d(pt.X + 1000, pt.Y + 100 - 100), 0, 0, 0); }
            }
            else
            {
                Shape_polyline.AddVertexAt(0, new Point2d(pt.X, pt.Y + 100), 0, 0, 0);
               
                Shape_polyline.AddVertexAt(1, new Point2d(pt.X + 1000, pt.Y + 100), 0, 0, 0);
                if (C_doub != 0)
                { Shape_polyline.AddVertexAt(2, new Point2d(pt.X + 1000, pt.Y + 100 - 100), 0, 0, 0); }
            }
              
            
            Btr.AppendEntity(Shape_polyline);
            tr.AddNewlyCreatedDBObject(Shape_polyline, true);

            
            //A_text.TextString = A;
            // Open the TextStyleTable for read
            TextStyleTable acTextStyleTbl = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
            if (acTextStyleTbl.Has("Standard"))
            {
                // Get the TextStyleTableRecord for the "Standard" text style
                TextStyleTableRecord acTextStyleTblRec = tr.GetObject(acTextStyleTbl["Standard"], OpenMode.ForWrite) as TextStyleTableRecord;
                if (A!= "0")
                { CreateDbText(tr, Btr, acTextStyleTblRec, A, new Point3d(pt.X -10,pt.Y,0), Math.PI / 2); }
                if (B != "0")
                { CreateDbText(tr, Btr, acTextStyleTblRec, B, new Point3d(pt.X + 425, pt.Y+110, 0), 0); }
                if (C != "0")
                { CreateDbText(tr, Btr, acTextStyleTblRec, C, new Point3d(pt.X +1010, pt.Y+100, 0), -Math.PI / 2); }    
                    
            }
        }

        public void CreateDbText(Transaction tr, BlockTableRecord Btr,TextStyleTableRecord acTextStyleTblRec,string text, Point3d pt, double rotation )
        {
            DBText A_text = new DBText();
            A_text.TextString = text;
            A_text.Position = pt;
            A_text.Rotation = rotation;
            A_text.Height= 50;
            
            // Replace "A" with the actual text you want to set
            A_text.SetDatabaseDefaults(); // This sets the default properties for the text
            
            // Set the text object's TextStyleId to the "Standard" text style
            A_text.TextStyleId = acTextStyleTblRec.ObjectId;
           
            // Make the text annotative
            A_text.Annotative = AnnotativeStates.True;
           
            Btr.AppendEntity(A_text);
            tr.AddNewlyCreatedDBObject(A_text, true);
        }

        public static string[] LayDanhSachDL_Attribute(Transaction tr, SelectionSet acSSet, string AttributString)
        {
            // Dictionary to hold grouped attributes
            Dictionary<string, string> groupedAttributes = new Dictionary<string, string>();
            // Start a transaction

            foreach (SelectedObject acSSObj in acSSet)
            {
                if (acSSObj != null)
                {
                    // Open the block reference for read
                    BlockReference acBlkRef = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                    string Tenck = LayDL_Attribute(tr, acBlkRef, AttributString);
                    if (!groupedAttributes.ContainsKey(Tenck))
                    {
                        groupedAttributes.Add(Tenck, Tenck);

                    }
                }
            }
            string[] att_list = new string[groupedAttributes.Count];
            long a = 0;
            foreach (string key in groupedAttributes.Values)
            {
                att_list[a] = groupedAttributes[key];
                a = a + 1;
            }
            Array.Sort(att_list);
            return att_list;
        }

        public static string LayDL_Attribute(Transaction tr, BlockReference acBlkRef, string AttributString)
        {
            string Tenck = "";
            if (acBlkRef.Layer == "TKT_thepchu")
            {
                if (acBlkRef.AttributeCollection != null)
                {
                    foreach (ObjectId attId in acBlkRef.AttributeCollection)
                    {
                        AttributeReference acAttRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;

                        if (acAttRef.Tag == AttributString)
                        {
                            Tenck = acAttRef.TextString;
                        }
                    }
                }
            }

            return Tenck;
        }

    }
}
