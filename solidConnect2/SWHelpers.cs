using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using System.Diagnostics;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;

using Microsoft.VisualBasic;
using System.IO;

namespace solidConnect2
{
    class SWHelpers
    {
        public static ModelDoc2 newPart(string partTemplateFile, SldWorks swApp)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            swModel = (ModelDoc2)swApp.NewDocument(partTemplateFile, 0, 0, 0);
            return swModel;
        }
        public static AssemblyDoc newAssembly(string assemTemplateFile, SldWorks swApp)
        {
            AssemblyDoc assemb = default(AssemblyDoc);
            assemb = (AssemblyDoc)swApp.NewDocument(assemTemplateFile, 0, 0, 0);
            return assemb;
        }
        public static DrawingDoc newDrawing(string drawingTemplateFile , SldWorks swApp)
        {
            DrawingDoc drawDoc = default(DrawingDoc);
            //drawDoc = (DrawingDoc)swApp.NewDocument(drawingTemplateFile,(int)swDwgPaperSizes_e.swDwgPaperA3size, 0,0);
            drawDoc = swApp.NewDrawing2((int)swDwgTemplates_e.swDwgTemplateA3size, "", (int)swDwgPaperSizes_e.swDwgPaperA3size, 0.42, 0.297);
            return drawDoc;
        }
        public static void saveCloseDoc(SldWorks swApp, string filepath, string name )
        {
            ModelDoc2 doc = (ModelDoc2)swApp.ActiveDoc;
            doc.SaveAs(filepath + "\\" + name);
            swApp.CloseDoc(filepath + "\\" + name);
        }
        public static string[] getPartsFromFolder(string folder)
        {
            List<string> parts = new List<string>();
            string[] files = Directory.GetFiles(folder);
            for (int i = 0; i < files.Length; i++)
            {
                string ext = Path.GetExtension(files[i]);
                if (ext == ".sldprt"|| ext == ".SLDPRT")
                {
                    if(!files[i].Contains("~"))parts.Add(files[i]);
                }
            }
            return parts.ToArray();
       }
        public static void swSamples(int i,string template)
        {
            SldWorks swApp = new SldWorks();
            swApp.Visible = true;
            switch(i)
            {
                case 0:
                    offsetSurface(swApp, template);
                    break;
                case 1:
                    makeSolidTest(swApp, template);
                    break;
                case 2:
                    tempExtrudeSurface(swApp, template);
                    break;
                case 3:
                    refplaneCreate(swApp, template);
                    break;
                case 4:
                    refPlaneByPoints(swApp, template);
                    break;
                case 5:
                    sketch3d(swApp, template);
                    break; 
                case 6:
                    sketch3dPoints(swApp, template);
                    break;
                case 7:
                    modelannotation(swApp);
                    break;
            }
        }
        private static void modelannotation(SldWorks swApp)
        {
            ModelDoc2 swModel;
            ModelDocExtension swModelDocExt;
            DrawingDoc swDrawing;
            SelectionMgr swSelmgr;
            View swView;
            object[] annotations;
            object selAnnot;
            Annotation swAnnotation;
            SelectData swSelData;
            int mark;
            string retval;
            bool status;

            retval = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
            swModel = (ModelDoc2)swApp.NewDocument(retval, 0, 0, 0);
            swDrawing = (DrawingDoc)swModel;
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swSelmgr = (SelectionMgr)swModel.SelectionManager;

            // Create drawing from assembly
            //"C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2018\\samples\\tutorial\\api\\wrench.sldasm"
            //@"C:\Users\r.hudson\Documents\WORK\projects\Frontis3D\P2\part0.sldprt"
            swView = (View)swDrawing.CreateDrawViewFromModelView3("C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2018\\samples\\tutorial\\api\\wrench.sldasm", "*Front", 0.1314541543147, 0.1407887187817, 0);

            // Select and activate the view
            status = swModelDocExt.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            status = swDrawing.ActivateView("Drawing View1");

            swModel.ClearSelection2(true);

            // Insert the annotations marked for the drawing
            annotations = (object[])swDrawing.InsertModelAnnotations3((int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel, (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing, true, false, false, false);

            // Select and mark each annotation
            swSelData = swSelmgr.CreateSelectData();
            mark = 0;

            foreach (object annot in annotations)
            {
                selAnnot = annot;
                swAnnotation = (Annotation)selAnnot;
                status = swAnnotation.Select3(true, swSelData);
                swSelData.Mark = mark;
                mark = mark + 1;
            }

        }
        private static void refPlaneByPoints(SldWorks swApp, string templateFile)
        {

            ModelDoc2 swDoc = SWHelpers.newPart(templateFile, swApp);
            FeatureManager swFeatureManager = default(FeatureManager);
            SketchManager swSketchManager = default(SketchManager);
            SelectionMgr swSelMgr = default(SelectionMgr);
            SelectData swSelData = default(SelectData);
            swFeatureManager = (FeatureManager)swDoc.FeatureManager;
            swSketchManager = (SketchManager)swDoc.SketchManager;
            swSelMgr = (SelectionMgr)swDoc.SelectionManager;
            swSelData = (SelectData)swSelMgr.CreateSelectData();
            RefPlane swRefPlane = default(RefPlane);
            
            swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swSketchManager.Insert3DSketch(true);
            SketchPoint sk1 = swSketchManager.CreatePoint(0, 0, 0);
            SketchPoint sk2 = swSketchManager.CreatePoint(1, 1, 0);
            SketchPoint sk3 = swSketchManager.CreatePoint(0, 1, 1);
            
            
            swSketchManager.Insert3DSketch(true);
            bool status = swSketchManager.CreateSketchPlane(9, 9, 0);
        }
        
        private static void offsetSurface(SldWorks swApp, string templateFile)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            FeatureManager swFeatureManager = default(FeatureManager);
            SketchSegment sketchSegment = default(SketchSegment);
            SketchManager swSketchManager = default(SketchManager);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            SelectionMgr swSelectionManager = default(SelectionMgr);
            Edge swEdge = default(Edge);
            Body2 swBody = default(Body2);
            Body2 newBody1 = default(Body2);
            Body2 newBody2 = default(Body2);
            object pointArray = null;
            double[] points = new double[12];
            bool status = false;

            swModel = (ModelDoc2)swApp.NewDocument(templateFile, 0, 0, 0);
            swFeatureManager = (FeatureManager)swModel.FeatureManager;
            swSketchManager = (SketchManager)swModel.SketchManager;
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swSelectionManager = (SelectionMgr)swModel.SelectionManager;

            //Create extruded surface body
            points[0] = -0.0720746414289124;
            points[1] = -0.0283600245263074;
            points[2] = 0;
            points[3] = -0.0514715593755;
            points[4] = -0.00345025084396866;
            points[5] = 0;
            points[6] = 0;
            points[7] = 0;
            points[8] = 0;
            points[9] = 0.0872558597840225;
            points[10] = 0.0521037067517796;
            points[11] = 0;
            pointArray = points;
            sketchSegment = (SketchSegment)swSketchManager.CreateSpline((pointArray));
            swSketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);
            status = swModelDocExt.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 4, null, 0);
            swFeatureManager.FeatureExtruRefSurface2(true, false, false, 0, 0, 0.0508, 0.00254, false, false, false,
            false, 0.0174532925199433, 0.0174532925199433, false, false, false, false, false, false, false,
            false);
            swSelectionManager.EnableContourSelection = false;

            //Offset selected edge and create two temporary bodies
            status = swModelDocExt.SelectByID2("", "EDGE", -0.00623752003605205, 0.000329492391927033, 0.050581684437077, false, 0, null, 0);
            swEdge = (Edge)swSelectionManager.GetSelectedObject6(1, -1);
            swBody = (Body2)swEdge.GetBody();
            swBody = (Body2)swBody.Copy();
            //Using a copy of the selected surface body, create two new temporary bodies
            newBody1 = (Body2)swBody.MakeOffset(0.01, false);
            newBody2 = (Body2)swBody.MakeOffset(0.01, true);
            //Display and color the new temporary body blue
            newBody1.Display3(swModel, 0, (int)swTempBodySelectOptions_e.swTempBodySelectOptionNone);
            //Display and color the new temporary body red
            newBody2.Display3(swModel, 1, (int)swTempBodySelectOptions_e.swTempBodySelectOptionNone);


        }
        private static void makeSolidTest(SldWorks swApp,string templateFile)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            Modeler swModeler = default(Modeler);
            Feature swFeat = default(Feature);
            double[] nConeParam = new double[9];
            object vConeArr = null;
            Body2 swConeBody = default(Body2);
            double[] nBoxParam = new double[9];
            object vBoxArr = null;
            Body2 swBoxBody = default(Body2);
            object[] vNewBodyArr = null;
            object vNewBody = null;
            PartDoc swNewPart = default(PartDoc);
            Body2 swNewBody = default(Body2);
            FaultEntity swFaultEnt = default(FaultEntity);
            int nRetVal = 0;
            int nCount = 0;

            // Form cone
            // Face center
            nConeParam[0] = 0.0;
            nConeParam[1] = 0.1;
            nConeParam[2] = 0.0;
            // Axis
            nConeParam[3] = 0.0;
            nConeParam[4] = 0.0;
            nConeParam[5] = 1.0;
            // Base radius
            nConeParam[6] = 0.2;
            // Top radius
            nConeParam[7] = 0.1;
            // Height
            nConeParam[8] = 0.3;
            vConeArr = nConeParam;

            // Form box
            // Face center
            nBoxParam[0] = 0.0;
            nBoxParam[1] = 0.1;
            nBoxParam[2] = 0.2;
            // Axis
            nBoxParam[3] = 0.0;
            nBoxParam[4] = 0.0;
            nBoxParam[5] = 1.0;
            // Width
            nBoxParam[6] = 0.3;
            // Length
            nBoxParam[7] = 0.25;
            //Height
            nBoxParam[8] = 0.4;
            vBoxArr = nBoxParam;

            swModeler = (Modeler)swApp.GetModeler();
            swConeBody = (Body2)swModeler.CreateBodyFromCone((vConeArr));
            swBoxBody = (Body2)swModeler.CreateBodyFromBox((vBoxArr));
            swFaultEnt = (FaultEntity)swConeBody.Check3;
            nCount = swFaultEnt.Count;
            if (nCount != 0)
            {
                Debug.Print("Faulty cone!");
                return;
            }
            swFaultEnt = (FaultEntity)swBoxBody.Check3;
            nCount = swFaultEnt.Count;
            if (nCount != 0)
            {
                Debug.Print("Faulty box!");
                return;
            }
            vNewBodyArr = (object[])swConeBody.Operations2((int)swBodyOperationType_e.SWBODYADD, swBoxBody, out nRetVal);

            swNewPart = (PartDoc)swApp.NewDocument(templateFile, 0, 0, 0);
            foreach (object vNewBody_loopVariable in vNewBodyArr)
            {
                vNewBody = vNewBody_loopVariable;
                swNewBody = (Body2)vNewBody;
                // Create solid body feature
                swFeat = (Feature)swNewPart.CreateFeatureFromBody3(swNewBody, false, (int)swCreateFeatureBodyOpts_e.swCreateFeatureBodyCheck + (int)swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify);
            }

            swModel = (ModelDoc2)swNewPart;
            swModel.ViewZoomtofit2();
            swModel.Save();
            swModel.Close();
        }

        private static void tempExtrudeSurface(SldWorks swApp, string templateFile)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            FeatureManager swFeatureManager = default(FeatureManager);
            SketchManager swSketchManager = default(SketchManager);
            SelectionMgr swSelectionManager = default(SelectionMgr);
            SketchSegment sketchSegment = default(SketchSegment);
            Modeler swModeler = default(Modeler);
            MathUtility swMath = default(MathUtility);
            Body2 profileBody = default(Body2);
            Body2 extrudedBody = default(Body2);
            MathVector dirVector = default(MathVector);
            SolidWorks.Interop.sldworks.Surface planeSurf = default(SolidWorks.Interop.sldworks.Surface);
            SolidWorks.Interop.sldworks.Curve[] trimCurves = new SolidWorks.Interop.sldworks.Curve[4];
            double[] points = new double[12];
            object pointArray = null;
            double halfWidth = 0;
            double halfLength = 0;
            double[] startArr = new double[3];
            double[] endArr = new double[3];
            double[] ptArr = new double[3];
            double[] dirArr = new double[3];
            double slotWidth = 0;
            double slotLength = 0;
            double slotDepth = 0;
            bool slotThruAll = false;
            bool status = false;

            swModeler = (Modeler)swApp.GetModeler();
            swMath = (MathUtility)swApp.GetMathUtility();
            swModel = (ModelDoc2)swApp.NewDocument(templateFile, 0, 0, 0);
            swFeatureManager = (FeatureManager)swModel.FeatureManager;
            swSketchManager = (SketchManager)swModel.SketchManager;
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swSelectionManager = (SelectionMgr)swModel.SelectionManager;

            //Create and select extruded surface body
            points[0] = -0.0720746414289124;
            points[1] = -0.0283600245263074;
            points[2] = 0;
            points[3] = -0.0514715593755;
            points[4] = -0.00345025084396866;
            points[5] = 0;
            points[6] = 0;
            points[7] = 0;
            points[8] = 0;
            points[9] = 0.0872558597840225;
            points[10] = 0.0521037067517796;
            points[11] = 0;
            pointArray = points;
            sketchSegment = (SketchSegment)swSketchManager.CreateSpline((pointArray));
            swSketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);
            status = swModelDocExt.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 4, null, 0);
            swFeatureManager.FeatureExtruRefSurface2(true, false, false, 0, 0, 0.0508, 0.00254, false, false, false,
            false, 0.0174532925199433, 0.0174532925199433, false, false, false, false, false, false, false,
            false);
            swSelectionManager.EnableContourSelection = false;
            status = swModelDocExt.SelectByID2("Surface-Extrude1", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);

            slotDepth = 0.01;
            slotWidth = 0.04;
            slotLength = 0.09;
            slotThruAll = false;
            halfWidth = slotWidth / 2;
            halfLength = slotLength / 2;
            ptArr[0] = 0.0;
            ptArr[1] = 0.0;
            ptArr[2] = 0.0;
            dirArr[0] = 0.0;
            dirArr[1] = 0.0;
            dirArr[2] = 1.0;
            startArr[0] = 1.0;
            startArr[1] = 0.0;
            startArr[2] = 0.0;
            planeSurf = (SolidWorks.Interop.sldworks.Surface)swModeler.CreatePlanarSurface2((ptArr), (dirArr), (startArr));

            ptArr[0] = -halfLength;
            ptArr[1] = halfWidth;
            ptArr[2] = 0.0;
            dirArr[0] = 1.0;
            dirArr[1] = 0.0;
            dirArr[2] = 0.0;
            trimCurves[0] = (SolidWorks.Interop.sldworks.Curve)swModeler.CreateLine((ptArr), (dirArr));
            trimCurves[0] = (SolidWorks.Interop.sldworks.Curve)trimCurves[0].CreateTrimmedCurve2(-halfLength, halfWidth, 0.0, halfLength, halfWidth, 0.0);

            ptArr[0] = halfLength;
            ptArr[1] = 0.0;
            ptArr[2] = 0.0;
            startArr[0] = halfLength;
            startArr[1] = halfWidth;
            startArr[2] = 0.0;
            endArr[0] = halfLength;
            endArr[1] = -halfWidth;
            endArr[2] = 0.0;
            dirArr[0] = 0.0;
            dirArr[1] = 0.0;
            dirArr[2] = -1.0;
            trimCurves[1] = (SolidWorks.Interop.sldworks.Curve)swModeler.CreateArc((ptArr), (dirArr), halfWidth, (startArr), (endArr));
            trimCurves[1] = (SolidWorks.Interop.sldworks.Curve)trimCurves[1].CreateTrimmedCurve2(halfLength, halfWidth, 0.0, halfLength, -halfWidth, 0.0);

            ptArr[0] = halfLength;
            ptArr[1] = -halfWidth;
            ptArr[2] = 0.0;
            dirArr[0] = -1.0;
            dirArr[1] = 0.0;
            dirArr[2] = 0.0;
            trimCurves[2] = (SolidWorks.Interop.sldworks.Curve)swModeler.CreateLine((ptArr), (dirArr));
            trimCurves[2] = (SolidWorks.Interop.sldworks.Curve)trimCurves[2].CreateTrimmedCurve2(halfLength, -halfWidth, 0.0, -halfLength, -halfWidth, 0.0);

            ptArr[0] = -halfLength;
            ptArr[1] = 0.0;
            ptArr[2] = 0.0;
            startArr[0] = -halfLength;
            startArr[1] = -halfWidth;
            startArr[2] = 0.0;
            endArr[0] = -halfLength;
            endArr[1] = halfWidth;
            endArr[2] = 0.0;
            dirArr[0] = 0.0;
            dirArr[1] = 0.0;
            dirArr[2] = -1.0;
            trimCurves[3] = (SolidWorks.Interop.sldworks.Curve)swModeler.CreateArc((ptArr), (dirArr), halfWidth, (startArr), (endArr));
            trimCurves[3] = (SolidWorks.Interop.sldworks.Curve)trimCurves[3].CreateTrimmedCurve2(-halfLength, -halfWidth, 0.0, -halfLength, halfWidth, 0.0);
            profileBody = (Body2)planeSurf.CreateTrimmedSheet((trimCurves));

            dirArr[0] = 0.0;
            dirArr[1] = 0.0;
            dirArr[2] = -1.0;
            dirVector = (MathVector)swMath.CreateVector((dirArr));
            extrudedBody = (Body2)swModeler.CreateExtrudedBody(profileBody, dirVector, slotDepth);
            //extrudedBody.Display3(swModel, Information.RGB(1, 0, 0), (int)swTempBodySelectOptions_e.swTempBodySelectOptionNone);

        }

        private static void connect2(SldWorks swApp, string templateFile)
        {
            //ModelDoc2 swDoc = swApp.IActiveDoc2;
            ModelDoc2 swDoc = SWHelpers.newPart(templateFile, swApp);

            swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            swDoc.InsertSketch2(false);
            

            swDoc.SketchRectangle(0, 0.04, 0, 0.01, 0, 0, true);

            swDoc.FeatureManager.FeatureExtrusion(true, false, false, 0, 0, 0.09, 0, false, false, false, false, 0.0, 0.0, false, false, false, false, true, false, false);
            swDoc.SaveAs(@"C:\Users\r.hudson\Documents\WORK\projects\Frontis3D\P2\parttest1.sldprt");
            swApp.CloseDoc("parttest1.sldprt");
            swApp.ExitApp();

        }
        private static void sketch3d(SldWorks swApp, string templateFile)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            SketchManager swSketchManager = default(SketchManager);
            SketchSegment swSketchSegment = default(SketchSegment);
            Sketch swSketch = default(Sketch);
            bool status = false;

            //Open new part document
            swModel = SWHelpers.newPart(templateFile, swApp);
            //Insert 3D sketch of  two lines
            swSketchManager = (SketchManager)swModel.SketchManager;
            swSketchManager.Insert3DSketch(true);
            swSketchSegment = (SketchSegment)swSketchManager.CreateCenterLine(-0.082642, 0.005659, 0.0, -0.049926, 0.045073, 0.0);
            swSketch = (Sketch)swSketchManager.ActiveSketch;
            //status = swSketch.SetWorkingPlaneOrientation(0, 0, 0, 0, 1, 0, 0, 0, 1, 1,
            //0, 0);
            swSketchSegment = (SketchSegment)swSketchManager.CreateCenterLine(-0.049926, 0.045073, 0.0, -0.049926, -0.022634, -1.065874);
            swSketch = (Sketch)swSketchManager.ActiveSketch;
           
            swModel.ClearSelection2(true);
            swSketchManager.InsertSketch(true);
            swModel.ViewZoomtofit2();
            //Insert 2D sketch of a circle
            swModel.ActivateSelectedFeature();
            swModel.ClearSelection2(true);
            swSketchManager.InsertSketch(true);
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            status = swModelDocExt.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.ClearSelection2(true);
            swSketchSegment = (SketchSegment)swSketchManager.CreateCircle(-0.056401, 0.005985, 0.0, -0.054697, -0.005141, 0.0);
            swModel.ClearSelection2(true);
            swSketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);

            //Insert a 3D sketch plane
            swSketchManager.Insert3DSketch(true);
            status = swModelDocExt.SelectByID2("Line1@3DSketch1", "EXTSKETCHSEGMENT", -0.0565609614209999, 0.0370796232466087, 0, true, 0, null, 0);
            status = swModelDocExt.SelectByID2("Point4@3DSketch1", "EXTSKETCHPOINT", -0.0564010297276809, 0.00598490302365917, 0, true, 0, null, 0);
            status = swSketchManager.CreateSketchPlane(9, 9, 0);
            status = swModelDocExt.SelectByID2("Plane1", "SKETCHSURFACES", 0, 0, 0, false, 0, null, 0);
            swModel.ActivateSelectedFeature();
            swModel.ClearSelection2(true);
            swSketchManager.InsertSketch(true);

        }
        private static void refplaneCreate(SldWorks swApp, string templateFile)

        {

            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            FeatureManager swFeatureManager = default(FeatureManager);
            Feature swFeature = default(Feature);
            RefPlane swRefPlane = default(RefPlane);
            SelectionMgr swSelMgr = default(SelectionMgr);
            RefPlaneFeatureData swRefPlaneFeatureData = default(RefPlaneFeatureData);

            int fileerror = 0;

            int filewarning = 0;

            bool boolstatus = false;

            int planeType = 0;


            //D:\\Programs\\SOLIDWORKS\\lang\\english\\Tutorial
            swApp.OpenDoc6("C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2018\\samples\\tutorial\\api\\plate.sldprt", (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref fileerror, ref filewarning);

            swModel = (ModelDoc2)swApp.ActiveDoc;

            swModelDocExt = (ModelDocExtension)swModel.Extension;

            swFeatureManager = (FeatureManager)swModel.FeatureManager;

            swSelMgr = (SelectionMgr)swModel.SelectionManager;



            // Create a constraint-based reference plane

            //boolstatus = swModelDocExt.SelectByID2("", "FACE", 0.028424218552, 0.07057725774359, 0, true, 0, null, 0);
            //
            boolstatus = swModelDocExt.SelectByID2("", "EDGE", 0.05976462601598, 0.0718389621656, 0.0001242036435087, true, 1, null, 0);
            boolstatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(16, 0.8, 4, 0, 0, 0);



            // Get type of the just-created reference plane

            boolstatus = swModelDocExt.SelectByID2("Plane1", "PLANE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);

            swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);

            swRefPlaneFeatureData = (RefPlaneFeatureData)swFeature.GetDefinition();



            planeType = swRefPlaneFeatureData.Type2;

            Debug.Print("Type of reference plane using IRefPlaneFeatureData::Type2: ");

            switch (planeType)

            {

                case 0:

                    Debug.Print(" Invalid");

                    break;

                case 1:

                    Debug.Print(" Undefined");

                    break;

                case 2:

                    Debug.Print(" Line Point");

                    break;

                case 3:

                    Debug.Print(" Three Points");

                    break;

                case 4:

                    Debug.Print(" Line Line");

                    break;

                case 5:

                    Debug.Print(" Distance");

                    break;

                case 6:

                    Debug.Print(" Parallel");

                    break;

                case 7:

                    Debug.Print(" Angle");

                    break;

                case 8:

                    Debug.Print(" Normal");

                    break;

                case 9:

                    Debug.Print(" On Surface");

                    break;

                case 10:

                    Debug.Print(" Standard");

                    break;

                case 11:

                    Debug.Print(" Constraint-based");

                    break;

            }

            Debug.Print("");



            planeType = swRefPlaneFeatureData.Type;

            Debug.Print("Type of reference plane using IRefPlaneFeatureData::Type: ");

            switch (planeType)

            {

                case 0:

                    Debug.Print(" Invalid");

                    break;

                case 1:

                    Debug.Print(" Undefined");

                    break;

                case 2:

                    Debug.Print(" Line Point");

                    break;

                case 3:

                    Debug.Print(" Three Points");

                    break;

                case 4:

                    Debug.Print(" Line Line");

                    break;

                case 5:

                    Debug.Print(" Distance");

                    break;

                case 6:

                    Debug.Print(" Parallel");

                    break;

                case 7:

                    Debug.Print(" Angle");

                    break;

                case 8:

                    Debug.Print(" Normal");

                    break;

                case 9:

                    Debug.Print(" On Surface");

                    break;

                case 10:

                    Debug.Print(" Standard");

                    break;

                case 11:

                    Debug.Print(" Constraint-based");

                    break;

            }

            Debug.Print("");



            swModel.ClearSelection2(true);

        }
        private static void sketch3dPoints(SldWorks swApp, string templateFile)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            SketchManager swSketchManager = default(SketchManager);
            SketchSegment swSketchSegment = default(SketchSegment);
            SketchPoint swSketchPoint = default(SketchPoint);
            FeatureManager swFeatureManager = default(FeatureManager);
            Sketch swSketch = default(Sketch);
            RefPlane swRefPlane = default(RefPlane);
            bool status = false;

            //Open new part document
            swModel = SWHelpers.newPart(templateFile, swApp);
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swFeatureManager = (FeatureManager)swModel.FeatureManager;
            //Insert 3D sketch of  3 pts
            swSketchManager = (SketchManager)swModel.SketchManager;
            swSketchManager.Insert3DSketch(true);
            swSketchPoint = (SketchPoint)swSketchManager.CreatePoint(-0.82642, 0.5659,0);
            swSketch = (Sketch)swSketchManager.ActiveSketch;

            swSketchPoint = (SketchPoint)swSketchManager.CreatePoint(-0.49926, 0.45073, 0);
            swSketch = (Sketch)swSketchManager.ActiveSketch;

            swSketchPoint = (SketchPoint)swSketchManager.CreatePoint(0.49926, -0.45073, 0.5);
            swSketch = (Sketch)swSketchManager.ActiveSketch;

            swModel.ClearSelection2(true);
            swSketchManager.InsertSketch(true);

            ////Insert 2D sketch of a circle
            //swModel.ActivateSelectedFeature();
            //swModel.ClearSelection2(true);
            //swSketchManager.InsertSketch(true);
            //swModelDocExt = (ModelDocExtension)swModel.Extension;
            //status = swModelDocExt.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            //swModel.ClearSelection2(true);
            //swSketchSegment = (SketchSegment)swSketchManager.CreateCircle(-0.056401, 0.005985, 0.0, -0.054697, -0.005141, 0.0);
            //swModel.ClearSelection2(true);
            //swSketchManager.InsertSketch(true);
            //swModel.ClearSelection2(true);

            ////Insert a 3D sketch plane
            //swSketchManager.Insert3DSketch(true);
            swModel.ViewZoomtofit2();
            status = swModelDocExt.SelectByID2("Point1@3DSketch1", "EXTSKETCHPOINT", -0.82642, 0.5659,0, true, 0, null, 0);
            status = swModelDocExt.SelectByID2("Point2@3DSketch1", "EXTSKETCHPOINT", -0.49926, 0.45073, 0, true, 1, null, 0);
            status = swModelDocExt.SelectByID2("Point3@3DSketch1", "EXTSKETCHPOINT", 0.49926, -0.45073, 0.5, true, 2, null, 0);
            swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(4, 0, 4, 0, 4,0);
            
            //

        }
    }
}
