using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rhino.Geometry;
using Rhino.Collections;
using Rhino.Geometry.Collections;
using SolidWorks.Interop.swconst;
using System.Threading;

namespace solidConnect2
{
    class PanelPart
    {
        private SldWorks swApp;
        string template;
        private List<string> log = new List<string>();
        private ModelDoc2 swDoc;
        private FeatureManager swFeatureManager = default(FeatureManager);
        private ModelDocExtension swModelDocExt = default(ModelDocExtension);
        SketchManager swSketchManager =default(SketchManager);
        SelectionMgr swSelMgr = default(SelectionMgr);
       
        public PanelPart(SldWorks app, string temp,bool show)
        {
            swApp = app;
            template = temp;
            swDoc = SWHelpers.newPart(template, swApp);
            //swApp.Visible = show;
            swModelDocExt = (ModelDocExtension)swDoc.Extension;
            swFeatureManager = (FeatureManager)swDoc.FeatureManager;
            swSketchManager = (SketchManager)swDoc.SketchManager;
            swSelMgr = swDoc.SelectionManager;
        }
        public void sheetmetalFromSurfaces(List<string[]> sNames,Brep panel,Point3d faceP,string filename)
        {
            Feature swFeat = default(Feature);
            bool status;
            for (int i=0;i<sNames.Count;i++)
            {
                status = swModelDocExt.SelectByID2(sNames[i][0], "SURFACEBODY", 0, 0, 0, true, 1, null, 0);
            }
            swFeat = swFeatureManager.InsertSewRefSurface(true, false, false, 0.0000025, 0.0001);
            
            status = swModelDocExt.SelectByID2("Surface-Knit1", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            //select face
            status = swDoc.Extension.SelectByID2("", "FACE", faceP.X, faceP.Y,faceP.Z, true, 2, null, 0);
            //get the folds a plane on each
            List<Point3d> foldSelP = getFoldEdges(panel);
            //select the folds
            for(int p=0;p< foldSelP.Count; p++)
            {
                status = swDoc.Extension.SelectByID2("", "EDGE", foldSelP[p].X, foldSelP[p].Y, foldSelP[p].Z, true, 1, null, 0);
            }
            //convert to sheet metal
            status = swFeatureManager.InsertConvertToSheetMetal2(0.002, false, false, 0.002, 0.002, 0, 0.5, 0, 0.5, false);
            swDoc.ClearSelection2(true);
            //open flat pattern mode
            status = swDoc.Extension.SelectByID2("Flat-Pattern", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            swDoc.ClearSelection2(true);
            //bend states see http://help.solidworks.com/2018/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swSMBendState_e.html
            var bendresult = swDoc.SetBendState((int)swSMBendState_e.swSMBendStateFlattened);
            status = swDoc.EditRebuild3();
            //save the file
            saveClose(@"C:\Users\r.hudson\Documents\WORK\projects\Frontis3D\P2\", filename + ".sldprt");
            
            //Export sheet metal to a single drawing file
            //exportToDXF();
        }
        public void exportToDXF()
        {
            int options = 1;  //include flat-pattern geometry
//Bit #1: 1 to export flat-pattern geometry; 0 to not
//Bit #2: 1 to include hidden edges; 0 to not
//Bit #3: 1 to export bend lines; 0 to not
//Bit #4: 1 to include sketches; 0 to not
//Bit #5: 1 to merge coplanar faces; 0 to not
//Bit #6: 1 to export library features; 0 to not
//Bit #7: 1 to export forming tools; 0 to not
//Bit #8: 0
//Bit #9: 0
//Bit #10: 0
//Bit #11: 0
//Bit #12: 1 to export bounding box; 0 to not
//For example, if you want to export:
//flat - pattern geometry, bend lines, and sketches, then Bits 1, 3, and 4 are 1, the bitmask is 0001101, and you need to set SheetMetalOptions = 2 ^ 0 + 2 ^ 2 + 2 ^ 3 = 1 + 4 + 8 = 13.
                          PartDoc swPart;
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swPart = (PartDoc)swDoc;
            object varAlignment;
            double[] dataAlignment = new double[12];
            dataAlignment[0] = 0.0;
            dataAlignment[1] = 0.0;
            dataAlignment[2] = 0.0;
            dataAlignment[3] = 1.0;
            dataAlignment[4] = 0.0;
            dataAlignment[5] = 0.0;
            dataAlignment[6] = 0.0;
            dataAlignment[7] = 1.0;
            dataAlignment[8] = 0.0;
            dataAlignment[9] = 0.0;
            dataAlignment[10] = 0.0;
            dataAlignment[11] = 1.0;

            varAlignment = dataAlignment;
            string docpath = swModel.GetPathName();
            bool exported = swPart.ExportToDWG2(@"C:\Users\r.hudson\Documents\WORK\projects\Frontis3D\P2\testflatpanel.dwg", docpath, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, options, null);
        }

        public List<Point3d> getFoldEdges(Brep panel)
        {
            List<Point3d> foldps = new List<Point3d>();
            for (int b = 0; b < panel.Edges.Count; b++)
            {
                var afs = panel.Edges[b].AdjacentFaces();

                if (afs.Length == 2)
                {
                    Line ed = new Line(panel.Curves3D[b].PointAtStart, panel.Curves3D[b].PointAtEnd);
                    Point3d selP = ed.PointAt(0.5);
                    foldps.Add(selP);
                    
                }
            }
            return foldps;
        }
        public string[] surface(Plane rhinoOrigin, Rhino.Geometry.Curve curve3d)
        {
            RefPlane surfPlane = default(RefPlane);
            Feature swFeat = default(Feature);
            
            var solidWPlane = plane3Points(rhinoOrigin.Origin, rhinoOrigin.PointAt(0.4, 0), rhinoOrigin.PointAt(0, 0.4),ref surfPlane);
            Point3d solidOrigin = getPlaneOrigin(surfPlane);
            Point3dList corners = getPlaneCorners(surfPlane);
            Vector3d solidXAxis = corners[1] - corners[0];
            Vector3d solidYAxis = corners[1] - corners[2];
            Plane solidPlane = new Plane(solidOrigin, solidXAxis, solidYAxis);
            //loop curve and get points in solidplane
            Polyline pl = new Polyline();
            Point3dList pointsInPlane = new Point3dList();
            
            if (curve3d.TryGetPolyline(out pl))
            {
                
                Point3d map = new Point3d();
                for (int i = 0; i < pl.Count; i++)
                {
                    solidPlane.RemapToPlaneSpace(pl[i], out map);
                    pointsInPlane.Add(map);
                }
            }
            //select the plane
            swDoc.Extension.SelectByID2(solidWPlane[0], solidWPlane[1], rhinoOrigin.OriginX, rhinoOrigin.OriginY, rhinoOrigin.OriginZ, false, 0, null, 0);
            SketchSegment swSketchSegment = default(SketchSegment);

            //add the points and sketch the boundary
            swSketchManager.InsertSketch(true);
             for (int i = 0; i < pointsInPlane.Count - 1; i++)
                {
                    if (i == pointsInPlane.Count - 2)
                        swSketchSegment = (SketchSegment)swSketchManager.CreateLine(pointsInPlane[i].X, pointsInPlane[i].Y, 0, pointsInPlane[0].X, pointsInPlane[0].Y, 0);
                    else swSketchSegment = (SketchSegment)swSketchManager.CreateLine(pointsInPlane[i].X, pointsInPlane[i].Y, 0, pointsInPlane[i + 1].X, pointsInPlane[i + 1].Y, 0);
                 
            }
            
            swDoc.ClearSelection2(true);
            swSketchManager.InsertSketch(true);
            bool s = swDoc.InsertPlanarRefSurface();
            //get the name and type of the surface
            var nametype = getfeatureNameType(swSelMgr.GetSelectedObject6(1, -1));
            
            swDoc.ClearSelection2(true);
            return nametype;
        }
        public string[] getfeatureNameType(Feature swFeat)
        {
            
            
            string fType = "";
            string name = swFeat.GetNameForSelection(out fType);
            string[] featNameType = { name, fType };
            return featNameType;
        }
        public Point3d getPlaneOrigin(RefPlane rp)
        {
            MathTransform swXform = default(MathTransform);
            swXform = rp.Transform;
            var transArray =swXform.ArrayData;
            Point3d origin = new Point3d(swXform.ArrayData[9], swXform.ArrayData[10], swXform.ArrayData[11]);
            //        Debug.Print "    Origin = (" & -1# * swXform.ArrayData(9) * 1000# & ", " & -1# * swXform.ArrayData(10) * 1000# & ", " & -1# * swXform.ArrayData(11) * 1000# & ") mm"
            //Debug.Print "    Rot1   = (" & swXform.ArrayData(0) & ", " & swXform.ArrayData(1) & ", " & swXform.ArrayData(2) & ")"
            //Debug.Print "    Rot2   = (" & swXform.ArrayData(3) & ", " & swXform.ArrayData(4) & ", " & swXform.ArrayData(5) & ")"
            //Debug.Print "    Rot3   = (" & swXform.ArrayData(6) & ", " & swXform.ArrayData(7) & ", " & swXform.ArrayData(8) & ")"
            //Debug.Print "    Trans  = (" & swXform.ArrayData(9) * 1000# & ", " & swXform.ArrayData(10) * 1000# & ", " & swXform.ArrayData(11) * 1000# & ") mm"
            //Debug.Print "    Scale  = " & swXform.ArrayData(12)
            return origin;
        }
        public Point3dList getPlaneCorners(RefPlane rp)
        {
            Point3dList corners = new Point3dList();
            Object[] cpObj;
            cpObj = (Object[])rp.CornerPoints;
            MathPoint[] vMathPoints = new MathPoint[4];
            Double[] vArrayData;
            for (int i = 0; i <= cpObj.GetUpperBound(0); i++)
            {
                vMathPoints[i] = (MathPoint)cpObj[i];
                vArrayData = (Double[])(vMathPoints[i].ArrayData);
                Point3d p = new Point3d(vArrayData[0], vArrayData[1], vArrayData[2]);
                corners.Add(p);
            }
            //plotPlaneCorners(corners);
            return corners;
        }
        public string[] plane3Points(Point3d origin, Point3d xpt, Point3d ypt, ref RefPlane swRefPlane)
        {
           
            SketchPoint swSketchPoint = default(SketchPoint);
            SelectionMgr swSelMgr = default(SelectionMgr);
            swSelMgr = swDoc.SelectionManager;
            Feature swFeat = default(Feature);
            Sketch swSketch = default(Sketch);
            //RefPlane swRefPlane = default(RefPlane);
            bool status = false;

            //Open new part document
            swSketchManager.Insert3DSketch(true);
            swSketchPoint = (SketchPoint)swSketchManager.CreatePoint(origin.X, origin.Y, origin.Z);
            
            swSketch = (Sketch)swSketchManager.ActiveSketch;

            swSketchPoint = (SketchPoint)swSketchManager.CreatePoint(xpt.X,xpt.Y,xpt.Z);
            
            swSketch = (Sketch)swSketchManager.ActiveSketch;

            swSketchPoint = (SketchPoint)swSketchManager.CreatePoint(ypt.X, ypt.Y, ypt.Z);
            
            swSketch = (Sketch)swSketchManager.ActiveSketch;

            
            swSketchManager.Insert3DSketch(true);
            swDoc.ClearSelection2(true);
            swDoc.ViewZoomtofit2();
            status = swModelDocExt.SelectByID2("", "EXTSKETCHPOINT", origin.X, origin.Y, origin.Z, true, 0, null, 0);
            status = swModelDocExt.SelectByID2("", "EXTSKETCHPOINT", xpt.X, xpt.Y, xpt.Z, true, 1, null, 0);
            status = swModelDocExt.SelectByID2("", "EXTSKETCHPOINT", ypt.X, ypt.Y, ypt.Z, true, 2, null, 0);

            swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(4, 0, 4, 0, 4, 0);
            //getPlaneTransform(swRefPlane);
            //getPlaneCorners(swRefPlane);
            //swDoc.ClearSelection2(true);
            //swSketchManager.InsertSketch(true);
            //selectPlane(origin.X, origin.Y, origin.Z);
            //swDoc.ClearSelection2(true);
            //sketchRectangle();
            //swDoc.ClearSelection2(true);
            //swSketchManager.InsertSketch(true);

            selectPlane(origin.X, origin.Y, origin.Z);
            swFeat = swSelMgr.GetSelectedObject6(1, -1);
            string fType = "";
            string name = swFeat.GetNameForSelection(out fType);
            string[] featNameType = { name, fType };
            return featNameType;
            //polyline2dTest();

        }
        private void selectPlane(double x,double y,double z)
        {
            swDoc.Extension.SelectByID2("", "PLANE", x,y,z, false, 0, null, 0);
        }
        private void sketchRectangle()
        {
            swDoc.SketchRectangle(0, 4, 0, 1, 0, 0, true);
        }
        private void extrusion()
        {
            swDoc.FeatureManager.FeatureExtrusion(true, false, false, 0, 0, 0.09, 0, false, false, false, false, 0.0, 0.0, false, false, false, false, true, false, false);
        }
        private void saveCloseExitApp(string filepath,string name)
        {
            swDoc.SaveAs(filepath+"\\"+name);
            swApp.CloseDoc(name);
            swApp.ExitApp();
        }
        private void saveClose(string filepath, string name)
        {
            swDoc.SaveAs(filepath + "\\" + name);
            swApp.CloseDoc(name);
            
        }
        public void extrusionTest(string filepath, string name)
        {
            selectPlane(0,0,0);
            sketchRectangle();
            extrusion();
            saveCloseExitApp(filepath, name);
        }
        public void plotPlaneCorners(Point3dList corners)
        {
            SketchPoint sketchPt;

            swSketchManager.Insert3DSketch(true);
            for (int i = 0; i < corners.Count; i++)
            {
                sketchPt = swSketchManager.CreatePoint(corners[i].X, corners[i].Y, corners[i].Z);
            }
            swSketchManager.Insert3DSketch(true);
        }
        public void polyline3d(BrepCurveList curves)
        {
            SketchSegment swSketchSegment = default(SketchSegment);
            
            swSketchManager.Insert3DSketch(true);
            for(int i = 0; i < curves.Count; i++)
            {
                swSketchSegment = (SketchSegment)swSketchManager.CreateLine(curves[i].PointAtStart.X, curves[i].PointAtStart.Y, curves[i].PointAtStart.Z,curves[i].PointAtEnd.X, curves[i].PointAtEnd.Y, curves[i].PointAtEnd.Z);
            }
            swSketchManager.Insert3DSketch(true);
        }
        public void polyline3d(Rhino.Geometry.Curve c)
        {
            SketchSegment swSketchSegment = default(SketchSegment);
            swSketchManager.Insert3DSketch(true);
            swDoc.ClearSelection2(true);
            NurbsCurve nc = c.ToNurbsCurve();
            for (int i = 0; i < nc.Points.Count-1; i++)
            {
                
                swSketchSegment = (SketchSegment)swSketchManager.CreateLine(nc.Points[i].Location.X, nc.Points[i].Location.Y, nc.Points[i].Location.Z,
                    nc.Points[i+1].Location.X, nc.Points[i + 1].Location.Y, nc.Points[i + 1].Location.Z);
            }
            swDoc.ClearSelection2(true);
            swSketchManager.Insert3DSketch(true);
            swDoc.ClearSelection2(true);
        }
        public void polyline2dTest()
        {
            SketchSegment swSketchSegment = default(SketchSegment);
            swSketchManager.InsertSketch(true);
            swSketchSegment = (SketchSegment)swSketchManager.CreateLine(-0.68849, 0.30798, 0, -0.48088, 0.42774, 0);
            swSketchSegment = (SketchSegment)swSketchManager.CreateLine(-0.48088, 0.42774, 0, -0.4327, 0.24715, 0);
            swSketchSegment = (SketchSegment)swSketchManager.CreateLine(-0.4327, 0.24715, 0, -0.61852, 0.15244, 0);
            swSketchSegment = (SketchSegment)swSketchManager.CreateLine(-0.61852, 0.15244, 0, -0.68849, 0.30798, 0);
            swDoc.ClearSelection2(true);
            swSketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            //boolstatus = Part.Extension.SelectByID2("Sketch4", "SKETCH", 0, 0, 0, False, 1, Nothing, 0)
            //Part.ClearSelection2 True
            //boolstatus = Part.Extension.SelectByID2("Sketch4", "SKETCH", 0, 0, 0, False, 1, Nothing, 0)
            //boolstatus = Part.InsertPlanarRefSurface()
            //Part.ClearSelection2 True
        }
        public void polyline2d(Plane pl, Rhino.Geometry.Curve c,string[] nameType)
        {
            //swDoc.ClearSelection2(true);
            //swSketchManager.InsertSketch(true);
            //selectPlane(origin.X, origin.Y, origin.Z);
            //swDoc.ClearSelection2(true);
            //sketchRectangle();
            //swDoc.ClearSelection2(true);
            //swSketchManager.InsertSketch(true);
            
            //selectPlane(pl.Origin.X, pl.Origin.Y, pl.Origin.Z);
            swDoc.Extension.SelectByID2(nameType[0], nameType[1], pl.OriginX, pl.OriginY, pl.OriginZ, false, 0, null, 0);
            SketchSegment swSketchSegment = default(SketchSegment);
            

            swSketchManager.InsertSketch(true);
            Polyline pc = new Polyline();
            if (c.TryGetPolyline(out pc))
            {
                //swSketchSegment = (SketchSegment)swSketchManager.CreateLine(6467, 298, 0,6467, 0, 0);
                //swSketchSegment = (SketchSegment)swSketchManager.CreateLine(6467, 0, 0, 10340, 0, 0);
                //swSketchSegment = (SketchSegment)swSketchManager.CreateLine(10340, 0, 0, 10340, 298, 0);
                //swSketchSegment = (SketchSegment)swSketchManager.CreateLine(10340, 298, 0, 6467, 298, 0);
                
                for (int i = 0; i < pc.Count - 1; i++)
                {
                    if(i== pc.Count - 2)
                        swSketchSegment = (SketchSegment)swSketchManager.CreateLine(pc[i].X, pc[i].Y, 0,pc[0].X, pc[0].Y, 0);
                    else swSketchSegment = (SketchSegment)swSketchManager.CreateLine(pc[i].X, pc[i].Y, 0,pc[i + 1].X, pc[i + 1].Y, 0);
                }
            }
            swDoc.ClearSelection2(true);
            swSketchManager.InsertSketch(true);
            bool s = swDoc.InsertPlanarRefSurface();
            swDoc.ClearSelection2(true);
        }
    }
}
