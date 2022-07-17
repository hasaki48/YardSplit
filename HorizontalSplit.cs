using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using PARTITF;
using System.IO;

namespace YardSplit
{
    internal class HorizontalSplit
    {
        // 计算顶部高程和底部高程，分割流程函数
        internal static void HorizontalSplitProcessControl(INFITF.Application CATIA)
        {
            string yard_path = $"F:\\SourceCode\\CSharp\\YardSplit\\Yard.CATPart";
            string yard_dir = $"F:\\LaZui\\Yard_CATIA";
            Directory.CreateDirectory(yard_dir);
            int plane_count = 0;
            bool is_finish = false;
            while (is_finish == false)
            {
                plane_count++;
                // 打开单个坝段文件
                CATIA.Documents.Open(yard_path);
                double high_height = Parameters.TOP_HEIGHT - Parameters.LAYER_HEIGHT * (plane_count - 1);
                double low_height = Parameters.TOP_HEIGHT - Parameters.LAYER_HEIGHT * plane_count;
                if (Parameters.BOTTOM_HEIGHT >= low_height)
                {
                    is_finish = true;
                }
                HorizontalSplitLoft(CATIA, high_height, low_height);

                string save_path = $"F:\\LaZui\\Yard_CATIA\\Yard{high_height}-{low_height}.CATPart";
                CATIA.ActiveDocument.SaveAs(save_path);
                CATIA.ActiveDocument.Close();
            }
        }

        // 根据顶部高程和底部高程水平分割实体，功能函数
        private static void HorizontalSplitLoft(INFITF.Application CATIA, double high_height, double low_height)
        {
            Part part = (CATIA.ActiveDocument as PartDocument).Part;
            ShapeFactory shapeFactory = part.ShapeFactory as ShapeFactory;

            HybridShapeFill high_height_fill = SetOffsetPlaneAndFill(CATIA, part, high_height, 1);
            HybridShapeFill low_height_fill = SetOffsetPlaneAndFill(CATIA, part, low_height, 2);

            part.InWorkObject = part.Bodies.Item("地形");
            Reference emp_ref = part.CreateReferenceFromName("");
            Split first_split = shapeFactory.AddNewSplit(emp_ref, CatSplitSide.catNegativeSide);
            Reference first_split_ref = part.CreateReferenceFromObject(high_height_fill);
            first_split.Surface = first_split_ref;
            Common.Hide(CATIA, high_height_fill);
            part.UpdateObject(first_split);
            part.Update();

            Reference new_emp_ref = part.CreateReferenceFromName("");
            Split sec_split = shapeFactory.AddNewSplit(new_emp_ref, CatSplitSide.catPositiveSide);
            Reference sec_split_reference = part.CreateReferenceFromObject(low_height_fill);
            sec_split.Surface = sec_split_reference;
            Common.Hide(CATIA, low_height_fill);
            part.UpdateObject(sec_split);
            part.Update();
        }

        private static HybridShapeFill SetOffsetPlaneAndFill(INFITF.Application CATIA, Part part, double height, int plane_count)
        //private static void SetOffsetPlaneAndFill(INFITF.Application CATIA, Part part, double height, int plane_count)
        {
            #region 设置偏移平面
            HybridShapeFactory hybrid_shape_factory = part.HybridShapeFactory as HybridShapeFactory;
            Reference xy_plane = part.CreateReferenceFromObject((Reference)part.OriginElements.PlaneXY);
            HybridShapePlaneOffset offset = hybrid_shape_factory.AddNewPlaneOffset(xy_plane, height, false);
            Body body = part.Bodies.Item("地形");
            body.InsertHybridShape(offset);
            Common.Hide(CATIA, offset);
            part.InWorkObject = offset;
            part.Update();
            #endregion

            #region 进入草稿模式并重命名
            Sketch sketch = body.Sketches.Add((Reference)body.HybridShapes.GetItem($"平面.{plane_count+2}"));
            (body.HybridShapes.GetItem($"平面.{plane_count+2}") as HybridShape).set_Name($"{height}m水平面");
            sketch.set_Name($"{height}m水平面草图");
            sketch.SetAbsoluteAxisData(new object[] { 0, 0, height, 1, 0, 0, 0, 1, 0 });
            part.InWorkObject = sketch;
            part.Update();
            #endregion

            #region 设置座标轴
            Factory2D factory = sketch.OpenEdition();
            Axis2D axis = (Axis2D)sketch.GeometricElements.GetItem("绝对轴");
            (axis.GetItem("横向") as Line2D).ReportName = 1;
            (axis.GetItem("纵向") as Line2D).ReportName = 2;
            part.Update();
            #endregion

            #region 画圆
            Circle2D circle = factory.CreateClosedCircle(282070.781250, 3659790.250000, 445.763505);
            sketch.CloseEdition();
            Common.Hide(CATIA, sketch);
            part.Update();
            #endregion

            #region 圆填充平面
            Reference ske_ref = part.CreateReferenceFromObject(sketch);
            HybridShapeFill hori_plane = hybrid_shape_factory.AddNewFill();
            hori_plane.AddBound(ske_ref);
            hori_plane.Continuity = 0;
            body.InsertHybridShape(hori_plane);
            part.InWorkObject = hori_plane;
            part.Update();
            #endregion

            return hori_plane;
        }
    }
}
