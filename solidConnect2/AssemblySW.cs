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
using System.IO;

namespace solidConnect2
{
    class AssemblySW
    {
        private SldWorks swApp;
        string template;
        private List<string> log = new List<string>();
        private AssemblyDoc swAssem;

        public AssemblySW(SldWorks app, string temp)
        {
            template = temp;
            swApp = app;

            swAssem = SWHelpers.newAssembly(template, swApp);
            //swAssem.sa
        }
        public void addMulitParts(string folder,bool fixparts)
        {
            //http://help.solidworks.com/2018/english/api/sldworksapi/Add_Components_Example_CSharp.htm
            List<string> compNames = new List<string>();
            List<double> compXforms = new List<double>();
            List<string> compCoordSysNames = new List<string>();
            object vcompNames;
            object vcompXforms;
            object vcompCoordSysNames;
            object vcomponents;
            bool status;
            string[] parts = Directory.GetFiles(folder);
            for (int i = 0; i < parts.Length; i++)
            {
                string ext = Path.GetExtension(parts[i]);
                if (ext == ".sldprt")
                {
                    compNames.Add(parts[i]);
                    double[] xforms = defineTransfrom();
                    compXforms.AddRange(xforms);
                    compCoordSysNames.Add("Coordinate System"+i);
                }
            }
            vcompNames = compNames.ToArray();
            vcompXforms = compXforms.ToArray();
            vcompCoordSysNames = compCoordSysNames.ToArray();
            vcomponents = swAssem.AddComponents3((vcompNames), (vcompXforms), (vcompCoordSysNames));

            if (fixparts)
            {
                ModelDoc2 doc = (ModelDoc2)swApp.ActiveDoc;

                foreach (object c in (Array)vcomponents)
                {
                    Component2 comp = (Component2)c;
                    string compname = comp.Name2;
                    string id = comp.GetSelectByIDString();
                    //select part
                    status = doc.Extension.SelectByID2(id, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    swAssem.FixComponent();
                }
            }
        }
        public void saveCloseAssembly(string filepath, string name)
        {
            ModelDoc2 doc = (ModelDoc2)swApp.ActiveDoc;
            doc.SaveAs(filepath + "\\" + name);
            swApp.CloseDoc(name);
        }

        public void addPartsFromFolder(string folder)
        {
            string[] parts = Directory.GetFiles(folder);
            ModelDoc2 temppart =default(ModelDoc2);
            int longwarnings=0;
            int errors=0;
            ModelDoc2 doc = (ModelDoc2)swApp.ActiveDoc;
            string AssemblyTitle = doc.GetTitle();
            MathUtility swMU = default(MathUtility);
            swMU = swApp.GetMathUtility();
            for (int i=0;i<parts.Length;i++)
            {
                string ext = Path.GetExtension(parts[i]);
                if(ext==".sldprt")
                {
                    //preload part
                    temppart = swApp.OpenDoc6(parts[i], 1, 32, "",ref errors,ref longwarnings);
                    doc = swApp.ActivateDoc3(AssemblyTitle, true, 0, errors);
                    Component2 insertedComp = default(Component2);
                    //need the component center XYZ
                    double compX = 0;
                    double compY = 0;
                    double compZ = 0;
                    insertedComp = swAssem.AddComponent5(parts[i], 0, "", false, "", compX, compY, compZ);
                    string compname = insertedComp.Name2;
                    string id = insertedComp.GetSelectByIDString();
                    swApp.CloseDoc(parts[i]);
                    double[] transform = defineTransfrom();
                    var swTransform = swMU.CreateTransform(transform);
                    bool status = insertedComp.SetTransformAndSolve2(swTransform);
                }
            }
        }
        private double[] defineTransfrom()
        {
            double[] compXforms = new double[16];
            // Define the transformation matrix. See the IMathTransform API documentation. 

            // Add a rotational diagonal unit matrix (zero rotation) to the transform
            // x-axis components of rotation
            compXforms[0] = 1.0;
            compXforms[1] = 0.0;
            compXforms[2] = 0.0;
            // y-axis components of rotation
            compXforms[3] = 0.0;
            compXforms[4] = 1.0;
            compXforms[5] = 0.0;
            // z-axis components of rotation
            compXforms[6] = 0.0;
            compXforms[7] = 0.0;
            compXforms[8] = 1.0;

            // Add a translation vector to the transform (zero translation) 
            compXforms[9] = 0.0;
            compXforms[10] = 0.0;
            compXforms[11] = 0.0;

            // Add a scaling factor to the transform
            compXforms[12] = 0.0;

            // The last three elements in the transformation matrix are unused
            return compXforms;
        }
    }
}
