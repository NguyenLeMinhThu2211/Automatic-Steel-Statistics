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
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace AutoCAD_CSharp_plug_in1
{
    internal class CreateTagBlk
    {
        public void CreateTag(Document doc, Database db,Transaction tr, Point3d pt, double scale, string SH, string SL, string DK)
        {
            //pt là biến vị trí click chuột
            //scale là tỷ lệ của tag
          
            #region Tạo Các thuộc tính (Attribute)
            AttributeDefinition SoHieu_Att = new AttributeDefinition();
            SoHieu_Att.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
            SoHieu_Att.Tag = "Số hiệu";
            SoHieu_Att.Prompt = "";
            SoHieu_Att.Height = 2 * scale;
            AttributeDefinition NumberBar_Att = new AttributeDefinition();
            NumberBar_Att.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
            NumberBar_Att.Tag = "Số lượng";
            NumberBar_Att.Prompt = "";
            NumberBar_Att.Height = 2 * scale;
            AttributeDefinition DiacenterBar_Att = new AttributeDefinition();
            DiacenterBar_Att.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
            DiacenterBar_Att.Tag = "Đường kính";
            DiacenterBar_Att.Prompt = "";
            DiacenterBar_Att.Height = 2 * scale;
            #endregion
            #region Tạo tên ngẫu nhiên cho Block Attribute
            BlockTable bt;
            bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            ObjectId bID = ObjectId.Null;

            string btAttName = "KyHieuThepDoc_tkt";
            //string btAttName = GenerateRandomName(8);
            if (!bt.Has(btAttName))
            {
                #endregion
                #region add Attribute to blocktable

                ObjectId bIDAttribute = ObjectId.Null;
                using (BlockTableRecord BtrAtt = new BlockTableRecord())
                {
                    BtrAtt.Name = btAttName;
                    // set reference location for Block
                    BtrAtt.Origin = pt;
                    // Add the Attribute to the block
                    BtrAtt.AppendEntity(SoHieu_Att);
                    BtrAtt.AppendEntity(NumberBar_Att);
                    BtrAtt.AppendEntity(DiacenterBar_Att);
                    
                    // Add Block table reference to Block table
                    tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                    bt.Add(BtrAtt);
                    tr.AddNewlyCreatedDBObject(BtrAtt, true);
                    bID = BtrAtt.Id;
                    Point3d KH_phi_local = new Point3d(pt.X - 400, pt.Y + scale, 0);
                    CreateAtext(tr, BtrAtt, KH_phi_local, 2 * scale, "%%c");
                    CreateCircle(tr, BtrAtt, new Point3d(pt.X - 600 - scale * 2, pt.Y + scale, 0));
                }
                #endregion
                
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
                        Point3d SH_local = new Point3d(pt.X - 690, pt.Y, 0);
                        SetAttToBlock(tr, Br, SoHieu_Att, SH, scale, false, SH_local);
                        double SL_double = Convert.ToDouble(SL);
                        Point3d SL_local = new Point3d();
                        if (SL_double < 10)
                        {
                            SL_local = new Point3d(pt.X - 475, pt.Y + scale, 0);
                        }
                        else
                        {
                            SL_local = new Point3d(pt.X - 550, pt.Y + scale, 0);
                        }    
                        SetAttToBlock(tr, Br, NumberBar_Att, SL, scale, false, SL_local);
                        Point3d DK_local = new Point3d(pt.X - 300, pt.Y + scale, 0);
                        SetAttToBlock(tr, Br, DiacenterBar_Att, DK, scale, false, DK_local);
                        //CreateTagline(doc, db, tr, Btr, pt, scale);
                    }
                }
            }
            else 
            {
                
                BlockTableRecord acBlkDef = tr.GetObject(bt["KyHieuThepDoc_tkt"], OpenMode.ForRead) as BlockTableRecord;
                if (acBlkDef.Id != ObjectId.Null)
                {
                    using (BlockReference Br = new BlockReference(new Point3d(pt.X, pt.Y, 0), acBlkDef.Id)) //Set location to place block
                    {
                        BlockTableRecord Btr;
                        Btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        //Set Block to BlockReference
                        Btr.AppendEntity(Br);
                        tr.AddNewlyCreatedDBObject(Br, true);
                        //Set Attribute for Block Attribute
                        Point3d SH_local = new Point3d(pt.X - 750, pt.Y, 0);
                        SetAttToBlock(tr, Br, SoHieu_Att, SH, scale, false, SH_local);
                        double SL_double = Convert.ToDouble(SL);
                        Point3d SL_local = new Point3d();
                        if (SL_double < 10)
                        {
                            SL_local = new Point3d(pt.X - 500, pt.Y + scale, 0);
                        }
                        else
                        {
                            SL_local = new Point3d(pt.X - 575, pt.Y + scale, 0);
                        }
                        SetAttToBlock(tr, Br, NumberBar_Att, SL, scale, false, SL_local);
                        Point3d DK_local = new Point3d(pt.X - 300, pt.Y + scale, 0);
                        SetAttToBlock(tr, Br, DiacenterBar_Att, DK, scale, false, DK_local);
                        //CreateTagline(doc, db, tr, Btr, pt, scale);
                    }
                }
            }
        }
        public void SetAttToBlock(Transaction tr, BlockReference Br, AttributeDefinition AttDef, string TextString, double scale, bool Rotation90, Point3d pt)
        {
            AttributeReference attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(AttDef, Br.BlockTransform);
            attRef.TextString = TextString;
            //attRef.Layer = "0";
            Point3d position =pt;
            attRef.Position = position;
            Br.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
            if (Rotation90 == true)
            {
                attRef.Rotation = Math.PI / 2;
            }
        }
        public void SetAttToBlock_Left(Transaction tr, BlockReference Br, AttributeDefinition AttDef, string TextString, double scale, bool Rotation90, Point3d pt)
        {
            AttributeReference attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(AttDef, Br.BlockTransform);
            attRef.TextString = TextString;
            //attRef.Layer = "0";double doubleValue = Convert.ToDouble(stringValue);
            double TextTodouble = Convert.ToDouble(TextString);
            Point3d position = new Point3d();
            if (TextTodouble<10)
            { position = new Point3d(pt.X, pt.Y, pt.Z); }  
            else
            { position = new Point3d(pt.X - scale * 2, pt.Y, pt.Z); }    
            
            attRef.Position = position;
            Br.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
            if (Rotation90 == true)
            {
                attRef.Rotation = Math.PI / 2;
            }
        }
        public void CreateAtext(Transaction tr, BlockTableRecord Btr, Point3d pt,double size, string noidung)
        {
            DBText dbtext = new DBText();
            dbtext.TextString = noidung;
            dbtext.Layer = "0";
            dbtext.Position = pt;
            dbtext.Height = size;
            Btr.AppendEntity(dbtext);
            tr.AddNewlyCreatedDBObject(dbtext, true);
        }
        public void CreateCircle(Transaction tr, BlockTableRecord Btr, Point3d pt)
        {
            Circle acCirc = new Circle();
            acCirc.Center = pt;
            acCirc.Radius = 100;
            Btr.AppendEntity(acCirc);
            tr.AddNewlyCreatedDBObject(acCirc, true);
        }
        //    public void CreateTagline( Document doc, Autodesk.AutoCAD.DatabaseServices.Database db , Transaction tr, BlockTableRecord Btr, Point3d pt1, double scale)
        //{
        //    PromptEntityResult per;
        //    Polyline acPoly = new Polyline();
        //    PromptEntityOptions peo1 = new PromptEntityOptions("\nChọn thanh thép: ");
        //    per = doc.Editor.GetEntity(peo1);
        //    Curve curve = null;
            
        //    if (per.Status == PromptStatus.OK)
        //    {
                
        //        Entity entity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
        //        PropertyInfo propInfo = entity.GetType().GetProperty("Length");

        //        if (propInfo != null)
        //        {
        //            //Create the bar
        //            CreateNewBarLine createPolyL = new CreateNewBarLine();
        //            createPolyL.CreatePLine(db, tr, per);
        //            //create Tag line
        //            curve = entity as Curve;
        //            Point3d closestPoint = curve.GetClosestPointTo(pt1, false);
        //            acPoly.AddVertexAt(0, new Point2d(closestPoint.X, closestPoint.Y), 0, 0, 0);
        //            acPoly.AddVertexAt(1, new Point2d(pt1.X, pt1.Y+scale), 0, 0, 0);
        //            acPoly.AddVertexAt(2, new Point2d(pt1.X - 600, pt1.Y+scale), 0, 0, 0);
        //            Btr.AppendEntity(acPoly);
        //            tr.AddNewlyCreatedDBObject(acPoly, true);
        //        }
        //        else
        //        {
        //            doc.Editor.WriteMessage("\nĐối tượng chọn không có thuộc tính chiều dài.");
        //        }
        //    }
        //}
    }
}
