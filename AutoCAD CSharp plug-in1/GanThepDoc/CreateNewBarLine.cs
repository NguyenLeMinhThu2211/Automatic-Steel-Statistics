using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
using RandomName = AutoCAD_CSharp_plug_in1.library.RandomName;
namespace AutoCAD_CSharp_plug_in1
{
    internal class CreateNewBarLine
    {
        public void CreateBlockBar(Database db,Transaction tr, PromptEntityResult per, string SH, string SL, string DK)
        {
            Entity entity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Entity;

            #region Lấy chiều dài thanh thép
            PropertyInfo propInfo = entity.GetType().GetProperty("Length");
            double lengthdouble = (double)propInfo.GetValue(entity);
            lengthdouble = Math.Round(lengthdouble / 5) * 5;
            #endregion

            #region Copy Polyline
            Polyline clonedPline =new Polyline();
            if (entity is Line line)
            {
                Point3dCollection stretchPoints = new Point3dCollection();
                entity.GetStretchPoints(stretchPoints);
                clonedPline.AddVertexAt(0, new Point2d(stretchPoints[0].X, stretchPoints[0].Y), 0, 0, 0);
                clonedPline.AddVertexAt(1, new Point2d(stretchPoints[1].X, stretchPoints[1].Y), 0, 0, 0);
            }
            else
            {
                clonedPline = entity.Clone() as Polyline;
            }
            clonedPline.Layer = "TKT_thepchu";
            #endregion

            #region Xóa các đoạn móc thép
            double set_segment = 100; //đặt đoạn móc giới hạn
            //Kiểm tra tại vị trí cuối của Pline
            int Lastpoint = clonedPline.NumberOfVertices - 1;
            Point2d pt1 = clonedPline.GetPoint2dAt(Lastpoint);
            Point2d pt2 = clonedPline.GetPoint2dAt(Lastpoint - 1);
            double segmentLength = pt1.GetDistanceTo(pt2);
            // nếu đoạn này < đoạn móc giới hạn thì xóa
            if (segmentLength < set_segment)
            {
                clonedPline.RemoveVertexAt(Lastpoint);
            }
            //Kiểm tra tại vị trí đầu của Pline
            pt1 = clonedPline.GetPoint2dAt(1);
            pt2 = clonedPline.GetPoint2dAt(0);
            segmentLength = pt1.GetDistanceTo(pt2);

            // nếu đoạn này < đoạn móc giới hạn thì xóa
            if (segmentLength < set_segment)
            {
                clonedPline.RemoveVertexAt(0);
            }
            #endregion

            #region Tìm các đoạn A B C
            double X_check = 0;double Y_check = 0;
            double A_check = 0;double B_check = 0; double C_check = 0;
            for (int i=0; i < clonedPline.NumberOfVertices-1; i++ )
            {
                Y_check = clonedPline.GetPoint2dAt(i).Y - clonedPline.GetPoint2dAt(i+1).Y;
                if (Math.Abs(Y_check) >0.1)
                {
                    switch (i)
                    {
                        case 0: A_check = Math.Round(Math.Abs(Y_check)/5)*5; break;
                        case 1: C_check = Math.Round(Math.Abs(Y_check) / 5) * 5; break;
                        case 2: C_check = Math.Round(Math.Abs(Y_check) / 5) * 5; break;
                    }
                }
                X_check = clonedPline.GetPoint2dAt(i).X - clonedPline.GetPoint2dAt(i + 1).X;
                if (Math.Abs(X_check) >0.1 )
                {
                        B_check = Math.Round(Math.Abs(X_check) / 5) * 5;
                }

            }
            #endregion

            #region Tạo Các thuộc tính (Attribute)
            AttributeDefinition Tenck_Att = new AttributeDefinition();
            Tenck_Att.Tag = "Tên CK";
            Tenck_Att.Prompt = "";
            Tenck_Att.Invisible = true;

            AttributeDefinition SoHieu_Att = new AttributeDefinition();
            //SoHieu_Att.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
            SoHieu_Att.Tag = "Số hiệu";
            SoHieu_Att.Prompt = "";
            SoHieu_Att.Invisible = true;
            AttributeDefinition NumberBar_Att = new AttributeDefinition();
            //NumberBar_Att.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
            NumberBar_Att.Tag = "Số lượng";
            NumberBar_Att.Prompt = "";
            NumberBar_Att.Invisible= true;
            AttributeDefinition DiacenterBar_Att = new AttributeDefinition();
            //DiacenterBar_Att.Position = new Point3d(pt.X + 250, pt.Y - 950, 0);
            DiacenterBar_Att.Tag = "Đường kính";
            DiacenterBar_Att.Prompt = "";
            DiacenterBar_Att.Invisible = true;

            AttributeDefinition attBarLength = new AttributeDefinition();
            attBarLength.Tag = "Chiều dài (mm)";
            attBarLength.Prompt = "";
            attBarLength.Invisible = true;

            AttributeDefinition A_attDef = new AttributeDefinition();
            A_attDef.Tag = "A";
            A_attDef.Prompt = "";
            A_attDef.Invisible = true;
            AttributeDefinition B_attDef = new AttributeDefinition();
            B_attDef.Tag = "B";
            B_attDef.Prompt = "";
            B_attDef.Invisible = true;
            AttributeDefinition C_attDef = new AttributeDefinition();
            C_attDef.Tag = "C";
            C_attDef.Prompt = "";
            C_attDef.Invisible = true;

            #endregion

            string btAttName = RandomName.GenerateRandomName(8); //Tạo tên ngẫu nhiên cho Block

            #region khai báo các biến của Block
            BlockTable bt;
            bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            ObjectId bID = ObjectId.Null;
            #endregion

            #region Thêm block vào môi trường Autocad
            Point3d pt = new Point3d(pt1.X, pt1.Y, 0); //Khai báo điểm đặt block trùng với điểm đặt của Polyline
            using (BlockTableRecord BtrAtt = new BlockTableRecord())
            {
                BtrAtt.Name = btAttName;
                BtrAtt.Origin = pt;
                //Thêm thuộc tính vào block
                BtrAtt.AppendEntity(Tenck_Att);
                BtrAtt.AppendEntity(SoHieu_Att);
                BtrAtt.AppendEntity(NumberBar_Att);
                BtrAtt.AppendEntity(DiacenterBar_Att);
                BtrAtt.AppendEntity(attBarLength);
                BtrAtt.AppendEntity(A_attDef);
                BtrAtt.AppendEntity(B_attDef);
                BtrAtt.AppendEntity(C_attDef);

                //Thêm polyline mới vào block
                BtrAtt.AppendEntity(clonedPline);
                //Thêm Block vào Autocad
                tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                bt.Add(BtrAtt);
                tr.AddNewlyCreatedDBObject(BtrAtt, true);
                bID = BtrAtt.Id;
            }
            #endregion

            #region Chèn block vào Model Space
            if (bID != ObjectId.Null)
            {
                using (BlockReference Br = new BlockReference(new Point3d(pt.X, pt.Y, 0), bID)) //Set location to place block
                {
                    Br.Layer = "TKT_thepchu"; //Set layer cho block
                    BlockTableRecord Btr;
                    Btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                   
                    Btr.AppendEntity(Br);
                    tr.AddNewlyCreatedDBObject(Br, true);
                    //Đặt các thông số cho block attribute
                    SetAttToBlock(tr, Br, Tenck_Att, "", false, pt);
                    SetAttToBlock(tr, Br, SoHieu_Att, SH, false, pt);
                    SetAttToBlock(tr, Br, NumberBar_Att, SL, false, pt);
                    SetAttToBlock(tr, Br, DiacenterBar_Att, DK, false, pt);
                    SetAttToBlock(tr, Br, attBarLength, lengthdouble.ToString(), false, pt);
                    SetAttToBlock(tr, Br, A_attDef, A_check.ToString(), false, pt);
                    SetAttToBlock(tr, Br, B_attDef, B_check.ToString(), false, pt);
                    SetAttToBlock(tr, Br, C_attDef, C_check.ToString(), false, pt);
                }
            }
            #endregion

        }
        private void SetAttToBlock(Transaction tr, BlockReference Br, AttributeDefinition AttDef, string TextString, bool Rotation90, Point3d pt)
        {
            AttributeReference attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(AttDef, Br.BlockTransform);
            attRef.TextString = TextString;
            //attRef.Layer = "0";
            Point3d position = pt;
            attRef.Position = position;
            Br.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
            if (Rotation90 == true)
            {
                attRef.Rotation = Math.PI / 2;
            }
        }

    }
}
