using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using SolidWorks.Interop.sldworks;

namespace solidConnect2
{
    public class solidConnect : GH_Component
    {
        BackgroundWorker bw = new BackgroundWorker();
        GH_Structure<IGH_Goo> panels = new GH_Structure<IGH_Goo>();
        GH_Structure<IGH_Goo> facepoints = new GH_Structure<IGH_Goo>();
        StreamWriter sw;
        
        bool makeParts;
        string format = "MMMddddHH:mmyyyy";
        string folder;
        bool showSolidApp;
        IGH_DataAccess Daccess;
        string startTime;
        private string templateFile;

        public solidConnect()
          : base("solidConnect", "solidConnect",
              "Description",
              "Extra", "RH")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Make parts", "", "", GH_ParamAccess.item, false);
            pManager.AddGenericParameter("BREPs to export", "", "", GH_ParamAccess.tree);
            pManager.AddGenericParameter("points on front face","", "", GH_ParamAccess.tree);
            pManager.AddIntegerParameter("Run code sample", "", "", GH_ParamAccess.item, -1);
            pManager.AddBooleanParameter("Solidworks visible", "", "", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            makeParts = false;
            
            Daccess = DA;
            folder = "";
            templateFile = "D:\\Programs\\SOLIDWORKS\\lang\\english\\Tutorial\\Part.prtdot";
            showSolidApp = false;
            int runSample = -1;
            if (!DA.GetData(0, ref makeParts)) return;
            if (!DA.GetDataTree(1, out panels)) return;
            if (!DA.GetDataTree(2, out facepoints)) return;
            if (!DA.GetData(3, ref runSample)) return;
            if (!DA.GetData(4, ref showSolidApp)) return;
            if (runSample > 0) SWHelpers.swSamples(runSample, templateFile);
            if (!makeParts)
            {
                if (bw.WorkerSupportsCancellation == true && bw.IsBusy)
                {
                    bw.CancelAsync();
                    DateTime endT = DateTime.Now;
                    //sw = new StreamWriter(folder + "\\errorLog" + startTime + ".csv", true);
                    //sw.WriteLine("Process cancelled at:" + endT.ToString(format));
                    MessageBox.Show("Part defintion cancelled");
                    //sw.Close();
                    //updateGHLog("Cancelled");
                }
            }
            if (makeParts)
            {
                bw.WorkerReportsProgress = false;
                bw.WorkerSupportsCancellation = true;
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                if (bw.IsBusy != true)
                {
                    bw.RunWorkerAsync();
                    MessageBox.Show("Part defintion running");
                }

            }


        }
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            if (!worker.CancellationPending)
            {
                SldWorks swApp = new SldWorks();
                swApp.Visible = true;
                DateTime startT = DateTime.Now;
                format = "MMMddddHHmmyyyy";
                startTime = startT.ToString(format);
                //sw = new StreamWriter(folder + "\\errorLog" + startTime + ".csv");
                //sw.WriteLine("Inventor panel definition log on: " + startTime);
                int panelsTotal = 0;
                int errorCount = 0;

                //sw.Close();
                for (int i = 0; i < panels.Branches.Count; i++)
                {
                    for (int j = 0; j < panels.Branches[i].Count; j++)
                    {
                        Brep panel = new Brep();
                        Point3d fp = new Point3d();
                        panels[i][j].CastTo(out panel);
                        facepoints[i][j].CastTo(out fp);
                        PanelPart pp = new PanelPart(swApp, templateFile,showSolidApp);
                        List<string[]> surfaceNameTypes = new List<string[]>();
                        if ((worker.CancellationPending == true))
                        {
                            e.Cancel = true;
                        }
                        else
                        {
                            if (!worker.CancellationPending == true)
                            {
                                //call to first part method
                                //pp.polyline3d(panel.Curves3D);
                                for (int b=0;b<panel.Faces.Count;b++)
                                {
                                    Plane p1 = new Plane();
                                    //panel.Loops[b].Face.FrameAt(0,0,out p1);
                                    panel.Faces[b].TryGetPlane(out p1);
                                    //var planeNameType = pp.plane3Points(p1.Origin, p1.PointAt(0.4, 0), p1.PointAt(0, 0.2));
                                    //pp.polyline2d(p1, panel.Loops[b].To2dCurve(),planeNameType);
                                    var nametype = pp.surface(p1, panel.Loops[b].To3dCurve());

                                    surfaceNameTypes.Add(nametype);

                                }
                                pp.sheetmetalFromSurfaces(surfaceNameTypes,panel,fp,"part"+j);
                               // pp.polyline3d(panel.Curves3D);
                                //pp.extrusionTest(@"C:\Users\r.hudson\Documents\WORK\projects\Frontis3D\P2", "testpart.SLDPRT");
                            }
                        }
                        panelsTotal++;
                    }

                }
                if (!worker.CancellationPending == true)
                {
                    //finish off log data after all panels
                    DateTime endT = DateTime.Now;
                    TimeSpan runtime = endT - startT;
                    //sw = new StreamWriter(folder + "\\errorLog" + startTime + ".csv", true);
                    //sw.WriteLine(panelsTotal.ToString() + " panels defined with " + errorCount.ToString() + " errors");
                    //updateGHLog(panelsTotal.ToString() + " panels defined with " + errorCount.ToString() + " errors");
                    //sw.WriteLine("End process" + endT.ToString(format) + " runtime: " + Math.Round(runtime.TotalSeconds, 1) + " seconds");
                    //updateGHLog("End process" + endT.ToString(format) + " runtime: " + Math.Round(runtime.TotalSeconds, 1) + " seconds");
                    //sw.Close();
                    bw.CancelAsync();
                    makeParts = false;
                    return;
                }
            }

        }
        private void updateGHLog(string text)
        {
            GH_Structure<GH_String> errors = new GH_Structure<GH_String>();
            errors.Append(new GH_String(text));
            //Daccess.SetDataTree(0, errors);
        }





        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                //return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("11b415e2-e8cc-45b4-9fc7-e8e88b05cdde"); }
        }
    }
}
