using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;
using System.ComponentModel;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;

namespace solidConnect2
{
    public class SolidWorksAssembly : GH_Component
    {
        bool makeAssem;
        bool fixParts;
        string assmName;
        string folder;
        BackgroundWorker bw = new BackgroundWorker();
        private string templateFile;
        public SolidWorksAssembly()
          : base("SolidWorksAssembly", "SWA",
              "Description",
              "Extra", "RH")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Make assembly", "", "", GH_ParamAccess.item, false);
            pManager.AddTextParameter("Parts folder", "", "", GH_ParamAccess.item);
            pManager.AddTextParameter("Assembly name", "", "", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Fix parts", "", "", GH_ParamAccess.item, false);
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
            makeAssem = false;
            templateFile = "D:\\Programs\\SOLIDWORKS\\lang\\english\\Tutorial\\assem.asmdot";
            fixParts =false;
            assmName="";
            if (!DA.GetData(0, ref makeAssem)) return;
            if (!DA.GetData(1, ref folder)) return;
            if (!DA.GetData(2, ref assmName)) return;
            if (!DA.GetData(3, ref fixParts)) return;
            if (!makeAssem)
            {
                if (bw.WorkerSupportsCancellation == true && bw.IsBusy)
                {
                    bw.CancelAsync();
                    DateTime endT = DateTime.Now;
                    //sw = new StreamWriter(folder + "\\errorLog" + startTime + ".csv", true);
                    //sw.WriteLine("Process cancelled at:" + endT.ToString(format));
                    MessageBox.Show("assembly defintion cancelled");
                    //sw.Close();
                    //updateGHLog("Cancelled");
                }
            }
            if (makeAssem)
            {
                bw.WorkerReportsProgress = false;
                bw.WorkerSupportsCancellation = true;
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                if (bw.IsBusy != true)
                {
                    bw.RunWorkerAsync();
                    MessageBox.Show("assembly defintion running");
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
                    AssemblySW assem = new AssemblySW(swApp, templateFile);
                    //assem.addPartsFromFolder(@"C:\Users\r.hudson\Documents\WORK\projects\Frontis3D\P2");
                    assem.addMulitParts(folder, fixParts);
                    assem.saveCloseAssembly(folder, assmName);
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
                makeAssem = false;
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
            get { return new Guid("176e1506-7803-404a-91b6-eaee779bfa1d"); }
        }
    }
}