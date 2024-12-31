using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;
using System.ComponentModel;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;

namespace solidConnect2
{
    public class SolidWorksDrawing : GH_Component
    {
        bool makeDrawings;
      
        string assmName;
        string folder;
        BackgroundWorker bw = new BackgroundWorker();
        private string templateFile;
        public SolidWorksDrawing()
          : base("SolidWorksDrawing", "Nickname",
              "Description",
              "Extra", "RH")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Make drawings", "", "", GH_ParamAccess.item, false);
            pManager.AddTextParameter("Parts folder", "", "", GH_ParamAccess.item);
            
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
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            makeDrawings = false;
            templateFile = "D:\\Programs\\SOLIDWORKS\\lang\\english\\Tutorial\\draw.drwdot";
            
            assmName = "";
            if (!DA.GetData(0, ref makeDrawings)) return;
            if (!DA.GetData(1, ref folder)) return;
            
            
            if (!makeDrawings)
            {
                if (bw.WorkerSupportsCancellation == true && bw.IsBusy)
                {
                    bw.CancelAsync();
                    DateTime endT = DateTime.Now;
                    //sw = new StreamWriter(folder + "\\errorLog" + startTime + ".csv", true);
                    //sw.WriteLine("Process cancelled at:" + endT.ToString(format));
                    MessageBox.Show("drawing defintions cancelled");
                    //sw.Close();
                    //updateGHLog("Cancelled");
                }
            }
            if (makeDrawings)
            {
                bw.WorkerReportsProgress = false;
                bw.WorkerSupportsCancellation = true;
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                if (bw.IsBusy != true)
                {
                    bw.RunWorkerAsync();
                    MessageBox.Show("drawing defintions running");
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
                if ((worker.CancellationPending == true))
                {
                    e.Cancel = true;
                }
                else
                {
                    string[] parts = SWHelpers.getPartsFromFolder(folder);
                    //DrawingSW drawing = new DrawingSW(swApp, templateFile,parts);
                    foreach (string part in parts)
                    {
                        DrawingSW drawing = new DrawingSW(swApp, templateFile, part);
                    }

                }
            }
            if (!worker.CancellationPending == true)
            {
                //finish off log data 
                //DateTime endT = DateTime.Now;
                //TimeSpan runtime = endT - startT;
                //sw = new StreamWriter(folder + "\\errorLog" + startTime + ".csv", true);
                //sw.WriteLine(panelsTotal.ToString() + " panels defined with " + errorCount.ToString() + " errors");
                //updateGHLog(panelsTotal.ToString() + " panels defined with " + errorCount.ToString() + " errors");
                //sw.WriteLine("End process" + endT.ToString(format) + " runtime: " + Math.Round(runtime.TotalSeconds, 1) + " seconds");
                //updateGHLog("End process" + endT.ToString(format) + " runtime: " + Math.Round(runtime.TotalSeconds, 1) + " seconds");
                //sw.Close();
                bw.CancelAsync();
                makeDrawings = false;
                return;
            }
        }
        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("53e1856d-21e5-4caa-b491-039c6e07cf57"); }
        }
    }
}