using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Rhino.Geometry;
using Rhino.Collections;

namespace solidConnect2
{
    class DrawingSW
    {
        private SldWorks swApp;
        string template;
        private List<string> log = new List<string>();
        private DrawingDoc swDrawing;
        private ModelDoc2 activeDoc=default(ModelDoc2);
        public DrawingSW(SldWorks app, string temp)
        {
            template = temp;
            swApp = app;

            swDrawing = SWHelpers.newDrawing(template, swApp);
            //swAssem.sa
        }
        public DrawingSW(SldWorks app, string temp, string part)
        {
            template = temp;
            swApp = app;
            
            
            View dwgView = default(View);
            swDrawing = SWHelpers.newDrawing(template, swApp);
            activeDoc = swApp.ActiveDoc;
            
            string[] viewNames = { "Flat pattern","*Front","*Top","*Right","*Isometric" };
                //loop the predefined views
               
                double xpos = 0.105;
                double ypos = 0.1;
                int rows = 3;
                
                int rNum = 0;
                int cNum = 0;
                for (int i = 0; i < viewNames.Length; i++)
                {
                    
                dwgView = (View)swDrawing.CreateDrawViewFromModelView3(part, viewNames[i], xpos * cNum + xpos / 2, ypos * rNum + ypos / 2, 0.0);
                

                dwgView.ScaleDecimal = dwgView.ScaleDecimal / 2;
                if (viewNames[i].Contains("Flat"))
                {
                    dwgView.InsertBendTable(false, xpos, ypos, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, "A", "D:\\Programs\\SOLIDWORKS\\lang\\english\\bendtable-standard.sldbndtbt");
                }

                rNum++;
                    if (rNum == rows)
                    {
                        rNum = 1;
                        cNum++;
                    }
                    insertModelAnnotations(dwgView);
                    activeDoc.ClearSelection2(true);
                    //BoundingDims(dwgView, "XY");
                    // setAutoDims(dwgView);
                }
                activeDoc.EditRebuild3();

                //save the drawing
                string name = Path.GetFileNameWithoutExtension(part);
                string folder = Path.GetDirectoryName(part);
                SWHelpers.saveCloseDoc(swApp, folder, name + ".slddrw");
                //close the part
                swApp.CloseDoc(part);
            
        }
        private void addFlatPatternView(string part,View dwgView, double xpos, double ypos)
        {
            int longwarnings = 0;
            int errors = 0;
            bool status;
            ModelDoc2 temppart = default(ModelDoc2);
            //open part
            temppart = swApp.OpenDoc6(part, 1, 32, "", ref errors, ref longwarnings);
            //define a view it name and select it
            dwgView = swDrawing.CreateViewport3(0.01, 0.01, 0, 0.05);
            string viewname = dwgView.GetName2();
            status = activeDoc.Extension.SelectByID2(viewname, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            dwgView = (View)swDrawing.CreateUnfoldedViewAt3(xpos , ypos, 0, false);
            dwgView.InsertBendTable(false, xpos, ypos, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, "A", "D:\\Programs\\SOLIDWORKS\\lang\\english\\bendtable-standard.sldbndtbt");

        }

        public DrawingSW(SldWorks app, string temp,string[] parts)
        {
            template = temp;
            swApp = app;
            ModelDoc2 temppart = default(ModelDoc2);
            int longwarnings = 0;
            int errors = 0;
            View dwgView = default(View);
            foreach (string part in parts)
            {
                //open the part
                temppart = swApp.OpenDoc6(part, 1, 32, "", ref errors, ref longwarnings);
                swDrawing = SWHelpers.newDrawing(template, swApp);
                activeDoc = swApp.ActiveDoc;
                //insertFlatPattern(20, 0, 0, part);

                swDrawing.GenerateViewPaletteViews(part);
                object[] viewNames = (object[])swDrawing.GetDrawingPaletteViewNames();
                //loop the predefined views
                string viewPaletteName="";
                double xpos = 0.105;
                double ypos = 0.1;
                int rows = 3;
                int cols = 4;
                int rNum = 0;
                int cNum = 0;
                for (int i = 0; i < viewNames.Length; i++)
                {
                    viewPaletteName = (string)viewNames[i];
                    dwgView = (View)swDrawing.DropDrawingViewFromPalette2(viewPaletteName, xpos*cNum+xpos/2, ypos*rNum + ypos / 2, 0.0);
                    dwgView.ScaleDecimal= dwgView.ScaleDecimal/2;
                    if(viewPaletteName.Contains("Flat"))
                    {
                        dwgView.InsertBendTable(false,xpos, ypos, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, "A", "D:\\Programs\\SOLIDWORKS\\lang\\english\\bendtable-standard.sldbndtbt");
                    }
                    rNum++;
                    if (rNum == rows)
                    {
                        rNum = 1;  
                        cNum++;
                    }
                    insertModelAnnotations(dwgView);
                    activeDoc.ClearSelection2(true);
                    //BoundingDims(dwgView, "XY");
                    // setAutoDims(dwgView);
                }
                activeDoc.EditRebuild3();
                
                //save the drawing
                string name = Path.GetFileNameWithoutExtension(part);
                string folder = Path.GetDirectoryName(part);
                SWHelpers.saveCloseDoc(swApp, folder, name + ".slddrw");
                //close the part
                swApp.CloseDoc(part);
            }
        }
        private void insertModelAnnotations(View view )
        {
            object[] annotations;
            SelectData swSelData;
            SelectionMgr swSelmgr;
            swSelmgr = (SelectionMgr)activeDoc.SelectionManager;
            swSelData = swSelmgr.CreateSelectData();
            object selAnnot;
            Annotation swAnnotation;
            bool status;
            // Select and activate the view
            
            string name = view.GetName2();
            status = activeDoc.Extension.SelectByID2(name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            status = swDrawing.ActivateView(name);
            annotations = (object[])swDrawing.InsertModelAnnotations3((int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel, (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing, true, false, false, false);
            if (annotations == null) return;
            // Select and mark each annotation
            
            int mark = 0;

            foreach (object annot in annotations)
            {
                selAnnot = annot;
                swAnnotation = (Annotation)selAnnot;
                status = swAnnotation.Select3(true, swSelData);
                swSelData.Mark = mark;
                mark = mark + 1;
            }
        }

        private void setAutoDims(View view)
        {
            //autodims
            object polylines;
            var edges = view.GetPolylines6(0, out polylines);
            foreach (Edge e in (Array)edges)
            {
                if (e != null)
                {
                    Vertex v = e.GetStartVertex();
                    if (v != null)
                    {

                        addAutoDims(v);
                        break;
                    }
                }
            }
        }
        private void BoundingDims(View view,string planeOrientation)
        {
            Point3d p1 = new Point3d();
            Point3d p2 = new Point3d();
            Point3d p3 = new Point3d();
            BoundingBox bb = getBoundingBox(view);
            switch (planeOrientation)
            {
                case "XY":
                    p1 = new Point3d(bb.Min.X, bb.Min.Y,0);
                    p2 = new Point3d(bb.Max.X, bb.Min.Y, 0);
                    p3 = new Point3d(bb.Max.X, bb.Max.Y, 0);
                    break;
                case "XZ":
                    p1 = new Point3d(bb.Min.X, 0, bb.Min.Z);
                    p2 = new Point3d(bb.Min.X, 0, bb.Max.Z);
                    p3 = new Point3d(bb.Max.X, 0, bb.Min.Z);
                    break;
                case "YZ":
                    p1 = new Point3d(0, bb.Min.Y, bb.Min.Z);
                    p2 = new Point3d(0, bb.Min.Y, bb.Max.Z);
                    p3 = new Point3d(0, bb.Max.Y, bb.Min.Z);
                    break;
            }
            SketchManager swSketchManager = default(SketchManager);
            swSketchManager = (SketchManager)activeDoc.SketchManager;
            SketchPoint sp1 = default(SketchPoint);
            SketchPoint sp2 = default(SketchPoint);
            SketchPoint sp3 = default(SketchPoint);
            activeDoc.ClearSelection2(true);
            sp1 = (SketchPoint)swSketchManager.CreatePoint(p1.X, p1.Y, p1.Z);
            sp2 = (SketchPoint)swSketchManager.CreatePoint(p2.X, p2.Y, p2.Z);
            sp3 = (SketchPoint)swSketchManager.CreatePoint(p3.X, p3.Y, p3.Z);

            bool status = activeDoc.Extension.SelectByID2("", "SKETCHPOINT", p1.X, p1.Y, p1.Z, true, 3, null, 0);
            status = activeDoc.Extension.SelectByID2("", "SKETCHPOINT", p2.X, p2.Y, p2.Z, true, 3, null, 0);
            activeDoc.AddDimension2((p1.X+p2.X)/2, (p1.Y + p2.Y) / 2, (p1.Z + p2.Z) / 2);
            activeDoc.ClearSelection2(true);

            status = activeDoc.Extension.SelectByID2("", "SKETCHPOINT", p1.X, p1.Y, p1.Z, true, 3, null, 0);
            status = activeDoc.Extension.SelectByID2("", "SKETCHPOINT", p3.X, p3.Y, p3.Z, true, 2, null, 0);
            activeDoc.AddDimension2((p1.X + p3.X) / 2, (p1.Y + p3.Y) / 2, (p1.Z + p3.Z) / 2);
        }
        private Point3dList edgeExtremes(View view)
        {
            object polylines;
            var edges = view.GetPolylines6(0, out polylines);
            Point3dList pts = new Point3dList();
            Point3d temp = new Point3d();
            foreach (Edge e in (Array)edges)
            {
                if (e != null)
                {
                    Vertex sv = e.GetStartVertex();
                    Vertex ev = e.GetEndVertex();
                    if(sv!=null&&ev!=null)
                    {
                        var spoint = sv.GetPoint();
                        var epoint = ev.GetPoint();

                    }
                }
            }
            return pts;
        }

        private BoundingBox getBoundingBox(View view)
        {
            object polylines;
            var edges = view.GetPolylines6(0, out polylines);
            Point3dList pts = new Point3dList();
            Point3d temp = new Point3d();
            foreach (Edge e in (Array)edges)
            {
                if (e != null)
                {
                    Vertex sv = e.GetStartVertex();
                    if (sv != null)
                    {
                        var point = sv.GetPoint();
                        temp = new Point3d((double)point[0], (double)point[1], (double)point[2]);
                        pts.Add(temp);
                    }
                    Vertex ev = e.GetEndVertex();
                    if (ev != null)
                    {
                        var point = ev.GetPoint();
                        temp = new Point3d((double)point[0], (double)point[1], (double)point[2]);
                        pts.Add(temp);
                    }
                }
            }
            BoundingBox bb = new BoundingBox(pts);
            return bb;
        }
        private void addAutoDims(Vertex v)
        {
            int selmark=0;
            selmark = (int)swAutodimMark_e.swAutodimMarkHorizontalDatum;
            selmark = (int)swAutodimMark_e.swAutodimMarkVerticalDatum;
            selmark = (int)swAutodimMark_e.swAutodimMarkOriginDatum;
                var point = v.GetPoint();
            
            bool status = activeDoc.Extension.SelectByID2("", "VERTEX", (double)point[0], (double)point[1], (double)point[2], true, selmark, null, 0);
            int success = swDrawing.AutoDimension((int)swAutodimEntities_e.swAutodimEntitiesBasedOnPreselect, 
                (int)swAutodimScheme_e.swAutodimSchemeBaseline, 
                (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementAbove, 
                (int)swAutodimScheme_e.swAutodimSchemeBaseline, 
                (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight);
        }
        private void insertFlatPattern(int scale,double xpos, double ypos,string part)
        {
            View dwgView = default(View);
            dwgView = swDrawing.CreateFlatPatternViewFromModelView3(part, "", 0, 0, 0, false, false);
            dwgView.ScaleDecimal = dwgView.ScaleDecimal * 1.0/scale;

            activeDoc.EditRebuild3();
        }
    }
    
}
