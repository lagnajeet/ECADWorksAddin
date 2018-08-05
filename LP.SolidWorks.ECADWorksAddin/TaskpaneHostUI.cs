using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using System.Diagnostics;
using Emgu.CV.Util;
using System.Collections.Concurrent;
using System.Threading;

namespace LP.SolidWorks.ECADWorksAddin
{

    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]

    public partial class TaskpaneHostUI : UserControl
    {
        double PI = 3.14159;
        double RadPerDeg = (Math.PI / 180);
        private string boardPartPath = "";
        private string boardPartID = "";
        private string boardSketch = "";
        dynamic boardJSON = null;
        JArray boardStyleJSON = null;
        double boardHeight = 0;
        double[] boardFirstPoint = new double[2];
        private string topDecal = "";
        private string bottomDecal = "";
        private static Random random = new Random();
        private messageFrm messageForm = new messageFrm();
        double[] boardOrigin = new double[3];
        private AssemblyDoc globalSwAssy = null;
        private bool FileOpened = false;
        private bool isIDFFile = false;
        private string gerberFile = "";
        private int idstart = 5842;
        private class ComponentAttribute
        {
            public PointF location;
            public bool side;      //true is top and false is bottom
            public float rotation;
            public string modelID;
            public string modelPath;
            public string componentID;
            public string noteID;
            public string redef;
            public string part_number;

        };

        private class IDFBoardOutline
        {
            public PointF location;
            public float rotation;

        };
        private class ComponentLibrary
        {
            public string ModelPath;
            public double[] Transformation;
            //public double[] rotationCenter;
        };
        private ConcurrentDictionary<string, List<ComponentAttribute>> boardComponents = new ConcurrentDictionary<string, List<ComponentAttribute>> { };
        private ConcurrentDictionary<string, string> ComponentModels = new ConcurrentDictionary<string, string> { };
        private ConcurrentDictionary<string, double[]> ComponentRotationCenter = new ConcurrentDictionary<string, double[]> { };
        private ConcurrentDictionary<string, double[]> ComponentTransformations = new ConcurrentDictionary<string, double[]> { };
        private ConcurrentDictionary<string, string> ComponentNames = new ConcurrentDictionary<string, string> { };
        private ConcurrentDictionary<string, double[]> ComponentLibraryTransformations = new ConcurrentDictionary<string, double[]> { };
        private ConcurrentDictionary<string, double[]> ComponentLibraryCenterOfRotation = new ConcurrentDictionary<string, double[]> { };
        private List<string> CurrentNotesTop = new List<string> { };
        private List<string> CurrentNotesBottom = new List<string> { };
        private string CheckIconPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "check.ico");
        private string CrossIconPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "cross.ico");
        //private string PCBStylePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "PCBStyle.json");
        //bool selectionEventHandler = false;
        private string pluginName = "ECADWorks";
        private string libraryFileName = "PackageLibrary.json";
        private string PCBStylePath = "PCBStyle.json";
        private int topDecalID=-1;
        private int bottomDecalID = -1;
        private string openedFile = "";
        private bool modelExists(string componentID, SldWorks mSolidworksApplication)
        {
            if (componentID.Length < 2)
                return false;
            Component2 swComp;
            SelectionMgr swSelMgr = default(SelectionMgr);
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
                return true;
            return false;
        }
        private void getLibraryTransform()
        {
            UpdateAllTransformations(TaskpaneIntegration.mSolidworksApplication);
            ComponentLibraryTransformations.Clear();
            ComponentLibraryCenterOfRotation.Clear();
            foreach (var pair in ComponentModels)
            {
                if (boardComponents.ContainsKey(pair.Key))
                {
                    if (boardComponents[pair.Key].Count > 0)
                    {
                        for (int idx = 0; idx < boardComponents[pair.Key].Count; idx++)
                        {
                            if (modelExists(boardComponents[pair.Key][idx].modelID, TaskpaneIntegration.mSolidworksApplication))              //first component that exist
                            {
                                ComponentAttribute c = boardComponents[pair.Key][idx];
                                PointF locationOrig = c.location;
                                bool side = c.side;
                                float rotationOrig = c.rotation;
                                string componentID = c.componentID;
                                double[] ComponentTransformationOnBoard = ComponentTransformations[componentID];
                                double[] rotationCenter = ComponentRotationCenter[componentID];
                                MathUtility swMathUtil = TaskpaneIntegration.mSolidworksApplication.GetMathUtility();
                                double[] temp = null;
                                double[] zVector = getPerpendicullarUnitVector(c.modelID, TaskpaneIntegration.mSolidworksApplication);

                                double[] undoZRotation;
                                if (side)
                                    undoZRotation = rodriguesRotation(zVector, 360 - c.rotation);
                                else
                                    undoZRotation = rodriguesRotation(zVector, c.rotation);
                                temp = multiplySwTransform(ComponentTransformationOnBoard, undoZRotation);
                                double[] yVector = getParallelUnitVector(temp);
                                double[] undoYRotation;
                                if (side)
                                {
                                    undoYRotation = rodriguesRotation(yVector, 0);
                                }
                                else
                                {
                                    undoYRotation = rodriguesRotation(yVector, 180);
                                    temp = multiplySwTransform(temp, undoYRotation);
                                }

                                double[] nPts1 = new double[3];

                                nPts1[0] = ComponentTransformationOnBoard[9] - rotationCenter[0];
                                nPts1[1] = ComponentTransformationOnBoard[10] - rotationCenter[1];
                                nPts1[2] = ComponentTransformationOnBoard[11] - rotationCenter[2];
                                double[] transformationOrigin = new double[16];
                                for (int i = 0; i < 16; i++)
                                {
                                    if (!(i == 9 || i == 10 || i == 11))
                                    {
                                        transformationOrigin[i] = Math.Round(temp[i], 5);
                                    }
                                    else
                                        transformationOrigin[i] = 0;
                                }
                                //offset calutation. it has nothing to do the angles.
                                //transformationOrigin[9] = Math.Round((ComponentTransformationOnBoard[9] - boardOrigin[0]) - convertUnit(locationOrig.X), 8);
                                //if (!side)
                                //    transformationOrigin[9] = -transformationOrigin[9];
                                //transformationOrigin[10] = Math.Round((ComponentTransformationOnBoard[10] - boardOrigin[1]) - convertUnit(locationOrig.Y), 8);
                                transformationOrigin[9] = Math.Round((rotationCenter[0] - boardOrigin[0]) - convertUnit(locationOrig.X), 8);
                                transformationOrigin[10] = Math.Round((rotationCenter[1] - boardOrigin[1]) - convertUnit(locationOrig.Y), 8);
                                if (side)
                                    transformationOrigin[11] = Math.Round((ComponentTransformationOnBoard[11] - (boardHeight / 2)), 8);
                                else
                                    transformationOrigin[11] = Math.Round((-(ComponentTransformationOnBoard[11] + (boardHeight / 2))), 8);
                                if (ComponentLibraryTransformations.ContainsKey(pair.Key))
                                    ComponentLibraryTransformations[pair.Key] = transformationOrigin;
                                else
                                    ComponentLibraryTransformations.TryAdd(pair.Key, transformationOrigin);
                                if (ComponentLibraryCenterOfRotation.ContainsKey(pair.Key))
                                    ComponentLibraryCenterOfRotation[pair.Key] = nPts1;
                                else
                                    ComponentLibraryCenterOfRotation.TryAdd(pair.Key, nPts1);
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void loadLibrary(string libraryFileName, string packageName)
        {
            if (boardComponents.Count > 0)
            {
                ConcurrentDictionary<string, ComponentLibrary> libraryJSONList = JsonConvert.DeserializeObject<ConcurrentDictionary<string, ComponentLibrary>>(File.ReadAllText(@libraryFileName));
                int count = libraryJSONList.Count;
                foreach (var pair in boardComponents)
                {
                    if (packageName.Length > 2)
                    {
                        if (pair.Key == packageName)
                        {
                            if (libraryJSONList.ContainsKey(pair.Key))
                            {
                                List<ComponentAttribute> l = pair.Value;
                                for (int i = 0; i < l.Count; i++)
                                {
                                    ComponentAttribute c = l[i];
                                    if (c.modelID.Length > 2)              //it already exist in the board
                                    {
                                        boardComponents[pair.Key][i].modelPath = "";
                                        removeComponent(boardComponents[pair.Key][i].modelID, TaskpaneIntegration.mSolidworksApplication);
                                        boardComponents[pair.Key][i].modelID = "";
                                    }
                                    bool side = c.side;
                                    double[] ComponentTransformationOnLibrary = libraryJSONList[pair.Key].Transformation;
                                    //double[] ComponentCenterOfRotationOnLibrary = libraryJSONList[pair.Key].rotationCenter;

                                    MathTransform swTransform = default(MathTransform);
                                    MathUtility swMathUtil = TaskpaneIntegration.mSolidworksApplication.GetMathUtility();
                                    swTransform = swMathUtil.CreateTransform(ComponentTransformationOnLibrary);
                                    double[] transformationFinal = new double[16];
                                    double[] yVector = getParallelUnitVector(ComponentTransformationOnLibrary);
                                    double[] temp = null;
                                    double[] doYRotation;
                                    double[] doZRotation;
                                    double[] zVector;
                                    if (side)
                                        doYRotation = rodriguesRotation(yVector, 0);
                                    else
                                        doYRotation = rodriguesRotation(yVector, 180);
                                    temp = multiplySwTransform(ComponentTransformationOnLibrary, doYRotation);
                                    zVector = getPerpendicullarUnitVector(temp);
                                    if (side)
                                        doZRotation = rodriguesRotation(zVector, c.rotation);
                                    else
                                        doZRotation = rodriguesRotation(zVector, -c.rotation);
                                    temp = multiplySwTransform(temp, doZRotation);
                                    for (int j = 0; j < 16; j++)
                                    {
                                        if (!(j == 9 || j == 10 || j == 11))
                                        {
                                            transformationFinal[j] = temp[j];
                                        }
                                        else
                                        {
                                            transformationFinal[j] = 0;
                                        }
                                    }
                                    //offset calutation. it has nothing to do the angles.
                                    //if (side)
                                    //    transformationFinal[9] = convertUnit(c.location.X) + ComponentTransformationOnLibrary[9] + boardOrigin[0];
                                    //else
                                    //    transformationFinal[9] = convertUnit(c.location.X) - ComponentTransformationOnLibrary[9] + boardOrigin[0];
                                    //transformationFinal[10] = convertUnit(c.location.Y) + (ComponentTransformationOnLibrary[10] + boardOrigin[1]);
                                    if (side)
                                        transformationFinal[11] = (boardOrigin[2] + ComponentTransformationOnLibrary[11] + (boardHeight / 2));
                                    else
                                        transformationFinal[11] = (boardOrigin[2] - (ComponentTransformationOnLibrary[11] + (boardHeight / 2)));
                                    string insertedModelId= insertPart(@libraryJSONList[pair.Key].ModelPath, transformationFinal, c.componentID, TaskpaneIntegration.mSolidworksApplication);
                                    if (insertedModelId.Length > 2)
                                    {
                                        boardComponents[pair.Key][i].modelID = insertedModelId;
                                        boardComponents[pair.Key][i].modelPath = libraryJSONList[pair.Key].ModelPath;
                                        moveLibraryComponent(insertedModelId, new double[2] { ComponentTransformationOnLibrary[9], ComponentTransformationOnLibrary[10] }, c.location, TaskpaneIntegration.mSolidworksApplication);
                                        if (ComponentModels.ContainsKey(pair.Key))
                                            ComponentModels[pair.Key] = libraryJSONList[pair.Key].ModelPath;
                                        else
                                            ComponentModels.TryAdd(pair.Key, libraryJSONList[pair.Key].ModelPath);
                                    }
                                    else
                                    {
                                        boardComponents[pair.Key][i].modelID = "";
                                        boardComponents[pair.Key][i].modelPath = "";
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (libraryJSONList.ContainsKey(pair.Key))
                        {
                            List<ComponentAttribute> l = pair.Value;
                            for (int i = 0; i < l.Count; i++)
                            {
                                ComponentAttribute c = l[i];
                                if (c.modelID.Length > 2)              //it already exist in the board
                                {
                                    boardComponents[pair.Key][i].modelPath = "";
                                    removeComponent(boardComponents[pair.Key][i].modelID, TaskpaneIntegration.mSolidworksApplication);
                                    boardComponents[pair.Key][i].modelID = "";
                                }
                                bool side = c.side;
                                double[] ComponentTransformationOnLibrary = libraryJSONList[pair.Key].Transformation;
                                //double[] ComponentCenterOfRotationOnLibrary = libraryJSONList[pair.Key].rotationCenter;

                                MathTransform swTransform = default(MathTransform);
                                MathUtility swMathUtil = TaskpaneIntegration.mSolidworksApplication.GetMathUtility();
                                swTransform = swMathUtil.CreateTransform(ComponentTransformationOnLibrary);
                                double[] transformationFinal = new double[16];
                                double[] yVector = getParallelUnitVector(ComponentTransformationOnLibrary);
                                double[] temp = null;
                                double[] doYRotation;
                                double[] doZRotation;
                                double[] zVector;
                                if (side)
                                    doYRotation = rodriguesRotation(yVector, 0);
                                else
                                    doYRotation = rodriguesRotation(yVector, 180);
                                temp = multiplySwTransform(ComponentTransformationOnLibrary, doYRotation);
                                zVector = getPerpendicullarUnitVector(temp);
                                if (side)
                                    doZRotation = rodriguesRotation(zVector, c.rotation);
                                else
                                    doZRotation = rodriguesRotation(zVector, -c.rotation);
                                temp = multiplySwTransform(temp, doZRotation);
                                for (int j = 0; j < 16; j++)
                                {
                                    if (!(j == 9 || j == 10 || j == 11))
                                    {
                                        transformationFinal[j] = temp[j];
                                    }
                                    else
                                    {
                                        transformationFinal[j] = 0;
                                    }
                                }
                                //offset calutation. it has nothing to do the angles.
                                if (side)
                                    transformationFinal[11] = (boardOrigin[2] + ComponentTransformationOnLibrary[11] + (boardHeight / 2));
                                else
                                    transformationFinal[11] = (boardOrigin[2] - (ComponentTransformationOnLibrary[11] + (boardHeight / 2)));

                                string insertedModelId = insertPart(@libraryJSONList[pair.Key].ModelPath, transformationFinal, c.componentID, TaskpaneIntegration.mSolidworksApplication);
                                if (insertedModelId.Length > 2)
                                {
                                    boardComponents[pair.Key][i].modelID = insertedModelId;
                                    boardComponents[pair.Key][i].modelPath = libraryJSONList[pair.Key].ModelPath;
                                    moveLibraryComponent(insertedModelId, new double[2] { ComponentTransformationOnLibrary[9], ComponentTransformationOnLibrary[10] }, c.location, TaskpaneIntegration.mSolidworksApplication);
                                    if (ComponentModels.ContainsKey(pair.Key))
                                        ComponentModels[pair.Key] = libraryJSONList[pair.Key].ModelPath;
                                    else
                                        ComponentModels.TryAdd(pair.Key, libraryJSONList[pair.Key].ModelPath);
                                }
                                else
                                {
                                    boardComponents[pair.Key][i].modelID = "";
                                    boardComponents[pair.Key][i].modelPath = "";
                                }
                            }

                        }
                    }
                }
            }
        }
        private void saveLibraryasProgramData()
        {
            string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
            string libraryJSONPath = @Path.Combine(directory, pluginName, libraryFileName);
            string programdatapath = @Path.Combine(directory, pluginName);
            ConcurrentDictionary<string, ComponentLibrary> libraryJSONList = new ConcurrentDictionary<string, ComponentLibrary> { };
            if (File.Exists(libraryJSONPath))
                libraryJSONList = JsonConvert.DeserializeObject<ConcurrentDictionary<string, ComponentLibrary>>(File.ReadAllText(@libraryJSONPath));

            foreach (var pair in ComponentModels)
            {
                if (ComponentLibraryTransformations.ContainsKey(pair.Key))
                {
                    if (File.Exists(ComponentModels[pair.Key]))
                    {
                        ComponentLibrary c = new ComponentLibrary
                        {
                            ModelPath = ComponentModels[pair.Key],
                            Transformation = ComponentLibraryTransformations[pair.Key]
                        };
                        //c.rotationCenter= ComponentLibraryCenterOfRotation[pair.Key];
                        if (libraryJSONList != null)
                        {
                            if (libraryJSONList.ContainsKey(pair.Key))
                                libraryJSONList[pair.Key] = c;
                            else
                                libraryJSONList.TryAdd(pair.Key, c);
                        }
                        else
                        {
                            libraryJSONList = new ConcurrentDictionary<string, ComponentLibrary> { };
                            libraryJSONList.TryAdd(pair.Key, c);
                        }
                    }
                }
            }

            //SaveFileDialog Savefile = new SaveFileDialog();
            //Savefile.Filter = "Library JSON files (*.json) | *.json";
            //Savefile.OverwritePrompt = true;
            //Savefile.Title = "Save Library JSON file As. . .";
            //Savefile.SupportMultiDottedExtensions = false;
            //if (Savefile.ShowDialog() == DialogResult.OK)
            //{
            System.IO.Directory.CreateDirectory(programdatapath);
            using (StreamWriter file = File.CreateText(libraryJSONPath))
            {
                //string json = JsonConvert.SerializeObject(boardJSON);
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, libraryJSONList, typeof(ConcurrentDictionary<string, ComponentLibrary>));
            }
            //}
        }
        private void saveLibrary()
        {
            ConcurrentDictionary<string, ComponentLibrary> libraryJSONList = new ConcurrentDictionary<string, ComponentLibrary> { };
            foreach (var pair in ComponentModels)
            {
                if (ComponentLibraryTransformations.ContainsKey(pair.Key))
                {
                    if (File.Exists(ComponentModels[pair.Key]))
                    {
                        ComponentLibrary c = new ComponentLibrary
                        {
                            ModelPath = ComponentModels[pair.Key],
                            Transformation = ComponentLibraryTransformations[pair.Key]
                        };
                        //c.rotationCenter= ComponentLibraryCenterOfRotation[pair.Key];
                        libraryJSONList.TryAdd(pair.Key, c);
                    }
                }
            }
            SaveFileDialog Savefile = new SaveFileDialog
            {
                Filter = "Library JSON files (*.json) | *.json",
                OverwritePrompt = true,
                Title = "Save Library JSON file As. . .",
                SupportMultiDottedExtensions = false
            };
            if (Savefile.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter file = File.CreateText(@Savefile.FileName))
                {
                    //string json = JsonConvert.SerializeObject(boardJSON);
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(file, libraryJSONList, typeof(ConcurrentDictionary<string, ComponentLibrary>));
                }
            }
        }

        private dynamic readTemplate()
        {
            const string jsontemplate = "{\"IDF_Header\":{\"Version\":\"V4.0\",\"Creation_Date_Time\":\"Date_string\",\"Owner_Name\":\"None\",\"Owner_Phone\":\"\",\"Owner_EMail\":\"\",\"Source_App_Type\":\"ECAD\",\"Source_App_Vendor\":\"Autodesk\",\"Source_App_Name\":\"Eagle CAD\",\"Source_App_Version\":\"V1.0\",\"Source_Tx_Name\":\"Eagle JSON Export\",\"Source_Tx_Version\":\"V1.0\",\"Entity_Count\":{\"Elec_Part_Defs\":\"5\",\"Elec_Part_Insts\":\"7\",\"Mech_Part_Defs\":\"0\",\"Mech_Part_Insts\":\"0\",\"Board_Part_Defs\":\"1\",\"Board_Part_Insts\":\"1\",\"Board_Assy_Defs\":\"0\",\"Board_Assy_Insts\":\"0\",\"Panel_Part_Defs\":\"0\",\"Panel_Part_Insts\":\"0\",\"Panel_Assy_Defs\":\"0\",\"Panel_Assy_Insts\":\"0\"},\"Comp_Part\":[\"Annotation\",\"Cavity\",\"Circle\",\"Circular_Arc\",\"Cutout\",\"Electrical_Part\",\"Extrusion\",\"Leader\",\"Material\",\"Mechanical_Part\",\"Pin\",\"Polycurve\",\"Polycurve_Area\",\"Polygon\",\"Polyline\",\"Text\",\"Thermal_Model\",\"Thermal_CV\",\"Thermal_RV\"],\"Board_Part\":[\"Annotation\",\"Board_Part\",\"Cavity\",\"Circle\",\"Cutout\",\"Extrusion\",\"Figure\",\"Filled_Area\",\"Footprint\",\"Graphic\",\"Hole\",\"Keepin\",\"Keepout\",\"Leader\",\"Pad\",\"Physical_Layer\",\"Polycurve_Area\",\"Polygon\",\"Polyline\",\"Text\",\"Trace\"],\"Board_Assy\":[\"Board_Assembly\",\"Board_Part_Instance\",\"Electrical_Part_Instance\",\"Mechanical_Part_Instance\",\"Sublayout\"],\"Default_Units\":\"MM\",\"Min_Res\":\"0\",\"Notes\":\"\"},\"Assemblies\":{\"Board_Assembly\":{\"Entity_ID\":\"#2501\",\"Assy_Name\":\"Board_Assembly\",\"Part_Number \":\"Board_Assembly\",\"Units\":\"Global\",\"Type\":\"Traditional\",\"Board_Inst\":{\"Board_Part_Instance\":{\"Entity_ID\":\"#2502\",\"Part_Name\":\"untitled\",\"Refdes\":\"Unassigned\",\"XY_Loc\":[0.0,0.0],\"Rotation\":\"0.0\"}},\"Comp_Insts\":[]}},\"Parts\":{\"Board_Part\":{\"Entity_ID\":\"#2525\",\"Part_Name\":\"untitled\",\"Units\":\"Global\",\"Type\":\"Unspecified\",\"Phy_Layers\":[{\"Entity_ID\":\"#2526\",\"Layer_Name\":\"Conductor_Bottom\",\"Type\":\"Conductive\",\"Position\":\"1\",\"Thickness\":\"0.05\"},{\"Entity_ID\":\"#2527\",\"Layer_Name\":\"Dielectric_Middle\",\"Type\":\"Dielectric\",\"Position\":\"2\",\"Thickness\":\"1.500000\"},{\"Entity_ID\":\"#2528\",\"Layer_Name\":\"Conductor_Top\",\"Type\":\"Conductive\",\"Position\":\"3\",\"Thickness\":\"0.05\"}],\"Shape\":{\"Extrusion\":{\"Entity_ID\":\"#2529\",\"Top_Height\":\"1.6\",\"Bot_Height\":\"0\",\"Outline\":{\"Polycurve_Area\":{\"Entity_ID\":\"#2530\",\"Line_Font\":\"Solid\",\"Line_Color\":[0.0,0.0,0.0],\"Fill_Color\":[0.0,0.0,0.0],\"Vertices\":[]}}}},\"Features\":[]},\"Electrical_Part\":[]}}";
            dynamic ret=JsonConvert.DeserializeObject<dynamic>(jsontemplate);
            if (ret != null)
                return ret;
            return null;
        }
        private void populateStyleComboBox()
        {

            string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
            string boardStyleJSONFile = @Path.Combine(directory, pluginName, PCBStylePath);
            string pluginDataDir = @Path.Combine(directory, pluginName);
            Directory.CreateDirectory(pluginDataDir);
            if (!File.Exists(boardStyleJSONFile))
            {
                using (StreamWriter file = File.CreateText(@boardStyleJSONFile))
                {
                    file.Write("[{StyleName:\"Barebone\",Mask:[0,255,255,255],Trace:[255,51,115,184],Pads:[255,51,115,184],Silk:[0,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Green\",Mask:[255,0,142,37],Trace:[255,0,187,49],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Red\",Mask:[255,22,0,147],Trace:[255,38,0,232],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"White\",Mask:[225,225,225,255],Trace:[255,250,250,250],Pads:[255,195,195,195],Silk:[255,0,0,0],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Black\",Mask:[255,34,27,29],Trace:[255,50,40,40],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Blue\",Mask:[255,116,74,0],Trace:[255,160,100,0],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"OHS Park\",Mask:[255,71,3,54],Trace:[255,142,10,109],Pads:[255,0,150,188],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,0,150,188]}]");
                    file.Close();
                }
            }
            else
            {
                if (new FileInfo(boardStyleJSONFile).Length <= 0)
                {
                    using (StreamWriter file = File.CreateText(@boardStyleJSONFile))
                    {
                        file.Write("[{StyleName:\"Barebone\",Mask:[0,255,255,255],Trace:[255,51,115,184],Pads:[255,51,115,184],Silk:[0,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Green\",Mask:[255,0,142,37],Trace:[255,0,187,49],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Red\",Mask:[255,22,0,147],Trace:[255,38,0,232],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"White\",Mask:[225,225,225,255],Trace:[255,250,250,250],Pads:[255,195,195,195],Silk:[255,0,0,0],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Black\",Mask:[255,34,27,29],Trace:[255,50,40,40],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"Blue\",Mask:[255,116,74,0],Trace:[255,160,100,0],Pads:[255,195,195,195],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,128,128,128]},\r\n{StyleName:\"OHS Park\",Mask:[255,71,3,54],Trace:[255,142,10,109],Pads:[255,0,150,188],Silk:[255,255,255,255],Board:[255,0,63,79],Holes:[255,0,150,188]}]");
                        file.Close();
                    }
                }
            }
            boardStyleJSON = JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(@boardStyleJSONFile));
            if (boardStyleJSON != null)
            {
                if (boardStyleJSON.Count > 0)
                {
                    colorPresets.Items.Clear();
                    for (int i = 0; i < boardStyleJSON.Count; i++)
                    {
                        dynamic obj = boardStyleJSON[i];
                        string name = (string)obj.StyleName;
                        colorPresets.Items.Add(name);
                    }
                }
            }
            else
            {
                MessageBox.Show(PCBStylePath + " File is corrupted. Please fix (or delete) the file " + boardStyleJSONFile+" to enable the PCB style presets.");
            }

        }
        private List<Bgra> applyPCBStyle(int index)
        {
            Bgra maskColor;
            Bgra traceColor;
            Bgra padColor;
            Bgra silkColor;
            Bgra boardColor;
            Bgra holeColor;
            if (boardStyleJSON != null)
            {
                if (boardStyleJSON.Count > 0)
                {
                    dynamic obj = boardStyleJSON[index];
                    JArray array = obj.Mask;
                    maskColor = new Bgra((double)array[1], (double)array[2], (double)array[3], (double)array[0]);
                    array = obj.Trace;
                    traceColor = new Bgra((double)array[1], (double)array[2], (double)array[3], (double)array[0]);
                    array = obj.Pads;
                    padColor = new Bgra((double)array[1], (double)array[2], (double)array[3], (double)array[0]);
                    array = obj.Silk;
                    silkColor = new Bgra((double)array[1], (double)array[2], (double)array[3], (double)array[0]);
                    array = obj.Board;
                    boardColor = new Bgra((double)array[1], (double)array[2], (double)array[3], (double)array[0]);
                    array = obj.Holes;
                    holeColor = new Bgra((double)array[1], (double)array[2], (double)array[3], (double)array[0]);
                    return new List<Bgra> { maskColor, traceColor, padColor, silkColor, boardColor, holeColor };
                }
            }
            maskColor = new Bgra(37, 142, 0, 255);
            traceColor = new Bgra(49, 187, 0, 255);
            padColor = new Bgra(195, 195, 195, 255);
            silkColor = new Bgra(255, 255, 255, 255);
            boardColor = new Bgra(79, 63, 0, 255);
            holeColor = new Bgra(128, 128, 128, 255);
            return new List<Bgra> { maskColor, traceColor, padColor, silkColor, boardColor, holeColor };
        }
        private string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void readBoardData(string JSONfile)
        {
            boardJSON = JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(@JSONfile));
        }
        private string insertPart(string sCompName, double[] transformation, string componentID, SldWorks mSolidworksApplication, bool isBoard = false)
        {
            int errors = 0;
            int warnings = 0;
            bool boolstatus = false;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            MathTransform swTransform = default(MathTransform);
            double[] nPts = new double[16];
            string AssemblyTitle = swModel.GetTitle();
            //MessageBox.Show(sCompName);
            ModelDoc2 Part = mSolidworksApplication.OpenDoc6(sCompName, (int)swDocumentTypes_e.swDocPART, 0, "", ref errors, ref warnings);
            if (Part != null)
            {
                swModel = mSolidworksApplication.ActivateDoc3(AssemblyTitle, true, (int)swRebuildOnActivation_e.swUserDecision, errors);
                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                Component2 swComponent = swAssy.AddComponent5(sCompName, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                Part.Visible = false;
                swModel.ClearSelection2(true);
                if (swComponent != null)
                {
                    MathTransform swRotTransform = default(MathTransform);
                    MathUtility swMathUtil;
                    swMathUtil = mSolidworksApplication.GetMathUtility();
                    swRotTransform = swMathUtil.CreateTransform(transformation);
                    boolstatus = swComponent.SetTransformAndSolve2(swRotTransform);
                    //boolstatus = swModel.ForceRebuild3(false);
                    swModel.ClearSelection2(true);
                    swTransform = swComponent.Transform2;
                    for (int i = 0; i < 16; i++)
                        nPts[i] = (double)swTransform.ArrayData[i];
                    swMathUtil = mSolidworksApplication.GetMathUtility();
                    double[] BoxArray = (double[])swComponent.GetBox(false, false);
                    ComponentTransformations[componentID] = nPts;
                    ComponentRotationCenter[componentID] = new double[3] { (BoxArray[0] + BoxArray[3]) / 2, (BoxArray[1] + BoxArray[4]) / 2, (BoxArray[2] + BoxArray[5]) / 2 };
                }
                swModel.ClearSelection2(true);
                return (swComponent.Name2 + "@" + AssemblyTitle);
            }
            else
                swModel.ClearSelection2(true);
            return "";

        }
        private string insertNote(string text, double X, double Y, double Z, bool side, SldWorks mSolidworksApplication)
        {
            ModelDoc2 Part;
            Note myNote;
            Annotation myAnnotation;
            bool boolstatus;
            Part = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            Color color = lblSilkColor.BackColor;
            var argbarray = BitConverter.GetBytes(color.ToArgb())
                            .Reverse()
                            .ToArray();
            string swColor = "0x00" + argbarray[3].ToString("X2") + argbarray[2].ToString("X2") + argbarray[1].ToString("X2");
            myNote = (Note)Part.InsertNote("< FONT color = " + swColor + " >< FONT style = B > " + text);
            if ((myNote != null))
            {
                //myNote.SetTextVerticalJustification((int)swTextAlignmentVertical_e.swTextAlignmentBottom);        //SW > 2017
                myNote.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);
                myNote.LockPosition = true;
                myAnnotation = (Annotation)myNote.GetAnnotation();
                if ((myAnnotation != null))
                {
                    boolstatus = myAnnotation.SetPosition2(X, Y, Z);
                }
                Part.ClearSelection2(true);
                Part.WindowRedraw();
                return myNote.GetName() + "@Annotations";
            }

            return "";
        }
        private void removeNote(string noteID, SldWorks mSolidworksApplication)
        {
            ModelDoc2 Part;
            bool boolstatus;
            Part = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = Part.Extension.SelectByID2(noteID + "@Annotations", "NOTE", 0, 0, 0, false, 0, null, 0);
            //if (boolstatus)
            //{
            int DeleteOption = (int)swDeleteSelectionOptions_e.swDelete_Absorbed;
            Part.Extension.DeleteSelection2(DeleteOption);
            //}
        }
        private string insertPart(string sCompName, double x, double y, bool side, bool createnew, float angle, int rotationaxis, string componentID, SldWorks mSolidworksApplication, bool isBoard = false)
        {
            int errors = 0;
            int warnings = 0;
            bool boolstatus = false;
            if (createnew)
                mSolidworksApplication.NewAssembly();
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;

            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);
            double[] nPts = new double[3];
            object vData;
            MathPoint swOriginPt;
            MathVector sw_Axis;
            string AssemblyTitle = swModel.GetTitle();
            double z = 0;
            //MessageBox.Show(sCompName);
            ModelDoc2 Part = mSolidworksApplication.OpenDoc6(sCompName, (int)swDocumentTypes_e.swDocPART, 0, "", ref errors, ref warnings);
            if (Part != null)
            {
                swModel = mSolidworksApplication.ActivateDoc3(AssemblyTitle, true, (int)swRebuildOnActivation_e.swUserDecision, errors);
                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                if (!isBoard)
                {
                    x += boardOrigin[0];
                    y += boardOrigin[1];
                    if (side)
                        z = boardOrigin[2] + 0.0064;
                    else
                        z = boardOrigin[2] - 0.0064;
                }
                Component2 swComponent = swAssy.AddComponent5(sCompName, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", x, y, z);
                Part.Visible = false;
                swModel.ClearSelection2(true);

                if (isBoard)
                {
                    swTransform = default(MathTransform);
                    swTransform = swComponent.Transform2;
                    boardOrigin[0] = (double)swTransform.ArrayData[9];
                    boardOrigin[1] = (double)swTransform.ArrayData[10];
                    boardOrigin[2] = (double)swTransform.ArrayData[11];
                    translateComponentDragOperator(boardPartID, new double[3] { 0, 0, -swTransform.ArrayData[11] }, false, TaskpaneIntegration.mSolidworksApplication,true);
                    swTransform = default(MathTransform);
                    swTransform = swComponent.Transform2;
                    boardOrigin[0] = (double)swTransform.ArrayData[9];
                    boardOrigin[1] = (double)swTransform.ArrayData[10];
                    boardOrigin[2] = 0;// (double)swTransform.ArrayData[11];
                }
                else
                {
                    if (swComponent != null)
                    {
                        swTransform = default(MathTransform);
                        swTransform = swComponent.Transform2;
                        //MessageBox.Show(swTransform.ArrayData);
                        MathUtility swMathUtil;
                        swMathUtil = mSolidworksApplication.GetMathUtility();
                        double[] BoxArray = (double[])swComponent.GetBox(false, false); ;
                        nPts[0] = (BoxArray[0] + BoxArray[3]) / 2;
                        nPts[1] = (BoxArray[1] + BoxArray[4]) / 2;
                        nPts[2] = (BoxArray[2] + BoxArray[5]) / 2;
                        vData = nPts;

                        swOriginPt = swMathUtil.CreatePoint(vData);

                        //if (rotationaxis == 0)
                        //{
                        //    nPts[0] = (double)swTransform.ArrayData[0];
                        //    nPts[1] = (double)swTransform.ArrayData[1];
                        //    nPts[2] = (double)swTransform.ArrayData[2];
                        //}
                        //else if (rotationaxis == 1)
                        //{
                        //    nPts[0] = (double)swTransform.ArrayData[3];
                        //    nPts[1] = (double)swTransform.ArrayData[4];
                        //    nPts[2] = (double)swTransform.ArrayData[5];
                        //}
                        //else if (rotationaxis == 2)
                        //{
                        nPts[0] = (double)swTransform.ArrayData[6];
                        nPts[1] = (double)swTransform.ArrayData[7];
                        nPts[2] = (double)swTransform.ArrayData[8];
                        //}
                        vData = nPts;
                        sw_Axis = swMathUtil.CreateVector(vData);

                        swRotTransform = swMathUtil.CreateTransformRotateAxis(swOriginPt, sw_Axis, angle * RadPerDeg);
                        swTransform = swTransform.Multiply(swRotTransform);
                        //boolstatus = swComponent.SetTransformAndSolve2(swTransform);
                        //swTransform = swComponent.Transform2;

                        //swTransform = default(MathTransform);
                        //swTransform = swComponent.Transform2;
                        //MessageBox.Show(swTransform.ArrayData);
                        //swMathUtil = mSolidworksApplication.GetMathUtility();

                        //nPts[0] = (double)swTransform.ArrayData[9];
                        //nPts[1] = (double)swTransform.ArrayData[10];
                        //nPts[2] = (double)swTransform.ArrayData[11];
                        //vData = nPts;

                        //swOriginPt = swMathUtil.CreatePoint(vData);
                        //if (!side)
                        //{
                        //    if (rotationaxis == 0)
                        //    {
                        //        nPts[0] = (double)swTransform.ArrayData[3];
                        //        nPts[1] = (double)swTransform.ArrayData[4];
                        //        nPts[2] = (double)swTransform.ArrayData[5];
                        //    }
                        //    else if (rotationaxis == 1)
                        //    {
                        //        nPts[0] = (double)swTransform.ArrayData[6];
                        //        nPts[1] = (double)swTransform.ArrayData[7];
                        //        nPts[2] = (double)swTransform.ArrayData[8];
                        //    }
                        //    else if (rotationaxis == 2)
                        //    {
                        nPts[0] = (double)swTransform.ArrayData[0];
                        nPts[1] = (double)swTransform.ArrayData[1];
                        nPts[2] = (double)swTransform.ArrayData[2];
                        // }
                        vData = nPts;
                        sw_Axis = swMathUtil.CreateVector(vData);
                        if (side)
                            swRotTransform = swMathUtil.CreateTransformRotateAxis(swOriginPt, sw_Axis, 90 * RadPerDeg);
                        else
                            swRotTransform = swMathUtil.CreateTransformRotateAxis(swOriginPt, sw_Axis, 270 * RadPerDeg);
                        swTransform = swTransform.Multiply(swRotTransform);
                        //}
                        boolstatus = swComponent.SetTransformAndSolve2(swTransform);

                        //AutoAlignComponent(swComponent.Name2 + "@" + AssemblyTitle, x, y, side, TaskpaneIntegration.mSolidworksApplication);
                        //SelectionMgr swSelMgr = default(SelectionMgr);
                        //DateTime nNow;
                        //object Box = null;
                        //double[] BoxArray = new double[6];
                        //Box = swComponent.GetBox(false, false);
                        //BoxArray = (double[])Box;
                        //Box = null;
                        //double currentX = (BoxArray[0] + BoxArray[3]) / 2;
                        //double currentY = (BoxArray[1] + BoxArray[4]) / 2;
                        //double currentZ = (BoxArray[2] + BoxArray[5]) / 2;
                        //double ZHeight = Math.Abs(BoxArray[2] - BoxArray[5]);
                        //double actualX = x;
                        //double actualY = y;
                        //double actualZ = 0;
                        //if (side)
                        //    actualZ = boardOrigin[2] + boardHeight + ZHeight ;
                        //else
                        //    actualZ = boardOrigin[2] - ZHeight;
                        //double deltaX = actualX - currentX;
                        //double deltaY = actualY - currentY;
                        //double deltaZ = actualZ - currentZ;
                        //swTransform = swComponent.Transform2;
                        //swMathUtil = mSolidworksApplication.GetMathUtility();
                        double[] nPts1 = new double[16];
                        //for (int i = 0; i < 16; i++)
                        //{
                        //    if (i == 9)
                        //        nPts1[i] = (double)swTransform.ArrayData[i] + deltaX;// (double)swTransform.ArrayData[i]+ 0;
                        //    else if (i == 10)
                        //        nPts1[i] = (double)swTransform.ArrayData[i] + deltaY;// (double)swTransform.ArrayData[i]+ 0;
                        //    else if (i == 11)
                        //        nPts1[i] = (double)swTransform.ArrayData[i] + deltaZ;
                        //    else
                        //        nPts1[i] = (double)swTransform.ArrayData[i];

                        //}
                        //swRotTransform = swMathUtil.CreateTransform(nPts1);
                        //boolstatus = swComponent.SetTransformAndSolve2(swRotTransform);
                        ////boolstatus = swModel.ForceRebuild3(false);
                        //swModel.ClearSelection2(true);
                        x -= boardOrigin[0];
                        y -= boardOrigin[1];
                        AutoAlignComponentSMD(swComponent.Name2 + "@" + AssemblyTitle, x, y, side, mSolidworksApplication);
                        swTransform = swComponent.Transform2;
                        for (int i = 0; i < 16; i++)
                            nPts1[i] = (double)swTransform.ArrayData[i];
                        ComponentTransformations[componentID] = nPts1;
                        BoxArray = (double[])swComponent.GetBox(false, false); ;
                        nPts[0] = (BoxArray[0] + BoxArray[3]) / 2;
                        nPts[1] = (BoxArray[1] + BoxArray[4]) / 2;
                        nPts[2] = (BoxArray[2] + BoxArray[5]) / 2;
                        ComponentRotationCenter[componentID] = nPts;
                    }
                }
                //swModel.ViewZoomtofit2();
                //MathTransform swTransform = default(MathTransform);
                //swTransform = swComponent.Transform2;
                //string msg = "component origin = " + ((double)swTransform.ArrayData[9]).ToString() + "," + ((double)swTransform.ArrayData[10]).ToString() + "," + ((double)swTransform.ArrayData[11]).ToString();
                //MessageBox.Show(msg);
                //swModel.ViewZoomtofit2();
                swModel.ClearSelection2(true);
                return (swComponent.Name2 + "@" + AssemblyTitle);
            }
            else
                swModel.ClearSelection2(true);
            return "";

        }

        private void removeDeletedPackages()
        {
            List<int> deletedIndices = new List<int>();
            Parallel.ForEach(boardComponents, (c, loopState) =>
            {
                if (ComponentModels.ContainsKey(c.Key))
                {
                    if (ComponentModels[c.Key].Length > 2 && File.Exists(ComponentModels[c.Key]))
                    {
                        int deletedCount = 0;
                        for (int i = 0; i < c.Value.Count; i++)
                        {
                            if (c.Value[i].modelID == null || c.Value[i].modelID.Length < 2)
                            {
                                deletedCount++;
                            }
                        }

                        if (c.Value.Count > 0 && deletedCount >= c.Value.Count)
                        {
                            ComponentModels[c.Key] = "";
                            int j = 0;
                            bool found = false;
                            foreach (var pair in boardComponents)
                            {
                                if (pair.Key.Equals(c.Key))
                                {
                                    found = true;
                                    break;
                                }
                                j++;
                            }
                            if (found)
                            {
                                lock ("Mylock")
                                {
                                    deletedIndices.Add(j);
                                }
                            }

                        }
                    }
                }
            });
            // for (int i = 0; i < comboComponentList.Items.Count; i++)
            //{
            //   comboComponentList.SelectedIndex = i;
            //comboComponentList.Refresh();
            //labelAssignedFile.Refresh();
            // }
            for (int i = 0; i < deletedIndices.Count; i++)
            {
                comboComponentList.SelectedIndex = deletedIndices[i];
                labelAssignedFile.Text = "None";
                comboComponentList.Refresh();
                labelAssignedFile.Refresh();
            }
        }
        private int SwAssy_RenameItemNotify(int EntityType, string oldName, string NewName)
        {
            SldWorks swApp = TaskpaneIntegration.mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            string AssemblyTitle = swModel.GetTitle();

            //MessageBox.Show(swComp.Name2);
            string selectedName = oldName + "@" + AssemblyTitle;
            //List<ComponentAttribute> c = new List<ComponentAttribute>;
            //string componentPackage = "";
            Parallel.ForEach(boardComponents, (c, loopState) =>
            {
                for (int i = 0; i < c.Value.Count; i++)
                {
                    if (c.Value[i].modelID == selectedName)
                    {
                        c.Value[i].modelID = NewName + "@" + AssemblyTitle;
                        //loopState.Stop();
                        //break;
                    }
                }
                //Your stuff
            });

            if (selectedName == boardPartID)
                boardPartID = NewName + "@" + AssemblyTitle;
            return 0;
        }
        private int SwAssy_DeleteItemNotify(int EntityType, string itemName)
        {
            SldWorks swApp = TaskpaneIntegration.mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            string AssemblyTitle = swModel.GetTitle();

            //MessageBox.Show(swComp.Name2);
            string selectedName = itemName + "@" + AssemblyTitle;
            //List<ComponentAttribute> c = new List<ComponentAttribute>;
            //string componentPackage = "";
            Parallel.ForEach(boardComponents, (c, loopState) =>
            {
                for (int i = 0; i < c.Value.Count; i++)
                {
                    if (c.Value[i].modelID == selectedName)
                    {
                        c.Value[i].modelID = "";
                        c.Value[i].modelPath = "";
                        //loopState.Stop();
                        //break;
                    }
                }
                //Your stuff
            });

            removeDeletedPackages();
            comboComponentList.Refresh();
            return 0;
        }

        private int SwAssy_UserSelectionPostNotify()
        {
            SelectionMgr swSelMgr = default(SelectionMgr);
            SldWorks swApp = TaskpaneIntegration.mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            string AssemblyTitle = swModel.GetTitle();
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            Component2 swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                //MessageBox.Show(swComp.Name2);
                string selectedName = swComp.Name2 + "@" + AssemblyTitle;
                //List<ComponentAttribute> c = new List<ComponentAttribute>;
                string boardComponentsKey = "";
                Parallel.ForEach(boardComponents, (c, loopState) =>
                {
                    for (int i = 0; i < c.Value.Count; i++)
                    {
                        if (c.Value[i].modelID == selectedName)
                        {
                            //MessageBox.Show(c.Key);
                            boardComponentsKey = c.Key;
                            loopState.Stop();
                            break;
                        }
                    }
                });
                int j = 0;
                bool found = false;
                foreach (var pair in boardComponents)
                {
                    if (pair.Key.Equals(boardComponentsKey))
                    {
                        found = true;
                        break;
                    }
                    j++;
                }
                if (found)
                    comboComponentList.SelectedIndex = j;
            }

            return 0;
        }

        //angle is angle
        //0 =x, 1=y,2=z

        private void saveJSONtoFile(dynamic boardJSON)
        {
            dynamic components = boardJSON.Assemblies.Board_Assembly.Comp_Insts;
            int component_count = (int)components.Count;
            if (component_count > 0)
            {
                for (int i = 0; i < component_count; i++)
                {
                    dynamic currentComponent = components[i];
                    if (currentComponent.Entity_ID != null)
                    {
                        string Entity_ID = (string)currentComponent.Entity_ID;
                        if (ComponentTransformations.ContainsKey(Entity_ID))
                        {
                            double[] transformation = ComponentTransformations[Entity_ID];
                            if (transformation != null)
                            {
                                if (transformation.Length == 16)
                                {
                                    JArray array = JArray.FromObject(transformation);
                                    boardJSON.Assemblies.Board_Assembly.Comp_Insts[i].Transformation = array;
                                }
                            }
                        }
                    }
                }
            }

            dynamic packages = boardJSON.Parts.Electrical_Part;
            int packages_count = (int)packages.Count;
            for (int i = 0; i < packages_count; i++)
            {
                if (ComponentModels.ContainsKey((string)boardJSON.Parts.Electrical_Part[i].Part_Name))
                    boardJSON.Parts.Electrical_Part[i].Properties.Part_Model = ComponentModels[(string)boardJSON.Parts.Electrical_Part[i].Part_Name];
            }
            SaveFileDialog Savefile = new SaveFileDialog
            {
                Filter = "Board JSON files (*.json) | *.json",
                OverwritePrompt = true,
                Title = "Save Board JSON file As. . .",
                SupportMultiDottedExtensions = false
            };
            if (Savefile.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter file = File.CreateText(@Savefile.FileName))
                {
                    //string json = JsonConvert.SerializeObject(boardJSON);
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(file, boardJSON);
                }
            }
        }


        private void insertModels()
        {
            List<string> keys = new List<string>(boardComponents.Keys);
            foreach (string key in keys)
            {
                string componentPackage = key;
                if (ComponentModels.ContainsKey(componentPackage))
                {
                    string modelPath = ComponentModels[componentPackage];
                    List<ComponentAttribute> c = boardComponents[key];
                    for (int i = 0; i < c.Count; i++)
                    {
                        PointF pointFs = c[i].location;
                        bool side = c[i].side;
                        float rotation = c[i].rotation;
                        string entityID = c[i].componentID;
                        bool hadTransformation = false;
                        if (ComponentTransformations.ContainsKey(entityID))
                        {
                            double[] transformation = ComponentTransformations[entityID];
                            if (transformation != null)
                            {
                                if (transformation.Length == 16)
                                {
                                    boardComponents[key][i].modelID = insertPart(@modelPath, transformation, c[i].componentID, TaskpaneIntegration.mSolidworksApplication);
                                    hadTransformation = true;
                                }
                            }
                        }
                        if (!hadTransformation)
                            boardComponents[key][i].modelID = insertPart(@modelPath, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, false, rotation, 2, c[i].componentID, TaskpaneIntegration.mSolidworksApplication);
                    }
                }
            }
            //foreach (KeyValuePair<string, string> entry in ComponentModels)
            //{

            //}
        }
        private void propagateComponentList()
        {
            comboComponentList.Items.Clear();
            foreach (KeyValuePair<string, List<ComponentAttribute>> entry in boardComponents)
            {
                comboComponentList.Items.Add(entry.Key);
            }
        }

        private void AutoAlignAllComponent(SldWorks mSolidworksApplication)
        {
            List<string> keys = new List<string>(boardComponents.Keys);
            foreach (string key in keys)
            {
                string componentPackage = key;
                if (ComponentModels.ContainsKey(componentPackage))
                {
                    string modelPath = ComponentModels[componentPackage];
                    List<ComponentAttribute> c = boardComponents[key];
                    for (int i = 0; i < c.Count; i++)
                    {
                        string modelID = (string)boardComponents[key][i].modelID;
                        if (modelID.Length > 2)
                        {
                            PointF pointFs = c[i].location;
                            bool side = c[i].side;
                            AutoAlignComponent(modelID, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, TaskpaneIntegration.mSolidworksApplication);
                        }
                    }
                }
            }
        }

        private void UpdateAllTransformations(SldWorks mSolidworksApplication)
        {
            List<string> keys = new List<string>(boardComponents.Keys);
            foreach (string key in keys)
            {
                string componentPackage = key;
                if (ComponentModels.ContainsKey(componentPackage))
                {
                    string modelPath = ComponentModels[componentPackage];
                    List<ComponentAttribute> c = boardComponents[key];
                    for (int i = 0; i < c.Count; i++)
                    {
                        string modelID = (string)boardComponents[key][i].modelID;
                        if (modelID.Length > 2)
                        {
                            string componentID = (string)boardComponents[key][i].componentID;
                            List<double[]> ret = getTransformations(modelID, TaskpaneIntegration.mSolidworksApplication);
                            if (ret != null)
                            {
                                if (ComponentTransformations.ContainsKey(componentID))
                                    ComponentTransformations[componentID] = ret[1];
                                else
                                    ComponentTransformations.TryAdd(componentID, ret[1]);
                                if (ComponentNames.ContainsKey(componentID))
                                    ComponentNames[componentID] = boardComponents[key][i].modelID;
                                else
                                    ComponentNames.TryAdd(componentID, boardComponents[key][i].modelID);
                                if (ComponentRotationCenter.ContainsKey(componentID))
                                    ComponentRotationCenter[componentID] = ret[0];
                                else
                                    ComponentRotationCenter.TryAdd(componentID, ret[0]);
                            }
                        }
                    }
                }
            }
            
        }

        private List<double[]> getTransformations(string componentID, SldWorks mSolidworksApplication)
        {
            List<double[]> ret = new List<double[]> { };
            Component2 swComp;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            double[] nPts = new double[16];
            double[] center = new double[3];
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                MathUtility swMathUtil;
                swMathUtil = mSolidworksApplication.GetMathUtility();
                double[] BoxArray = (double[])swComp.GetBox(false, false);
                center[0] = (BoxArray[0] + BoxArray[3]) / 2;
                center[1] = (BoxArray[1] + BoxArray[4]) / 2;
                center[2] = (BoxArray[2] + BoxArray[5]) / 2;
                for (int i = 0; i < 16; i++)
                {
                    nPts[i] = (double)swTransform.ArrayData[i];
                }
                swModel.ClearSelection2(true);
                ret.Add(center);
                ret.Add(nPts);
                return ret;
            }
            swModel.ClearSelection2(true);
            return null;
        }
        private void AutoAlignComponentSMD(string componentID, double X, double Y, bool side, SldWorks mSolidworksApplication)
        {
            ////AssemblyDoc swAssy;
            bool boolstatus;
            Component2 swComp;
            object Box = null;
            double[] BoxArray = new double[6];
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);
            double[] nPts = new double[16];
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                Box = swComp.GetBox(false, false);
                BoxArray = (double[])Box;
                Box = null;
                double currentX = (BoxArray[0] + BoxArray[3]) / 2;
                double currentY = (BoxArray[1] + BoxArray[4]) / 2;
                double currentZ = (BoxArray[2] + BoxArray[5]) / 2;
                double ZHeight = Math.Abs(BoxArray[2] - BoxArray[5]);
                double actualX = X + boardOrigin[0];
                double actualY = Y + boardOrigin[1];
                double actualZ = 0;
                if (side)
                    actualZ = boardOrigin[2] + boardHeight/2 + (ZHeight / 2);
                else
                    actualZ = boardOrigin[2] - boardHeight/2 - (ZHeight / 2);
                double deltaX = actualX - currentX;
                double deltaY = actualY - currentY;
                double deltaZ = actualZ - currentZ;
                swTransform = swComp.Transform2;
                MathUtility swMathUtil;
                swMathUtil = swApp.GetMathUtility();
                for (int i = 0; i < 16; i++)
                {
                    if (i == 9)
                        nPts[i] = (double)swTransform.ArrayData[i] + deltaX;// (double)swTransform.ArrayData[i]+ 0;
                    else if (i == 10)
                        nPts[i] = (double)swTransform.ArrayData[i] + deltaY;// (double)swTransform.ArrayData[i]+ 0;
                    else if (i == 11)
                        nPts[i] = (double)swTransform.ArrayData[i] + deltaZ;
                    else
                        nPts[i] = (double)swTransform.ArrayData[i];

                }
                swRotTransform = swMathUtil.CreateTransform(nPts);
                boolstatus = swComp.SetTransformAndSolve2(swRotTransform);
                //boolstatus = swModel.ForceRebuild3(false);
                swModel.ClearSelection2(true);
            }
        }
        private void AutoAlignComponent(string componentID, double X, double Y, bool side, SldWorks mSolidworksApplication)
        {
            ////AssemblyDoc swAssy;
            bool boolstatus;
            bool bRet;
            Component2 swComp;
            object Box = null;
            double[] BoxArray = new double[6];
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);
            double[] nPts = new double[16];
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            DateTime nNow;
            if (swComp != null)
            {
                Box = swComp.GetBox(false, false);
                BoxArray = (double[])Box;
                Box = null;
                double currentX = (BoxArray[0] + BoxArray[3]) / 2;
                double currentY = (BoxArray[1] + BoxArray[4]) / 2;
                double currentZ = (BoxArray[2] + BoxArray[5]) / 2;
                double ZHeight = Math.Abs(BoxArray[2] - BoxArray[5]);
                double actualX = X + boardOrigin[0];
                double actualY = Y + boardOrigin[1];
                double actualZ = 0;
                if (side)
                    actualZ = boardOrigin[2] + boardHeight + (ZHeight / 2);
                else
                    actualZ = boardOrigin[2] - (ZHeight / 2);
                double deltaX = actualX - currentX;
                double deltaY = actualY - currentY;
                double deltaZ = actualZ - currentZ;
                swTransform = swComp.Transform2;
                MathUtility swMathUtil;
                swMathUtil = swApp.GetMathUtility();
                for (int i = 0; i < 16; i++)
                {
                    if (i == 9)
                        nPts[i] = (double)swTransform.ArrayData[i] + deltaX;// (double)swTransform.ArrayData[i]+ 0;
                    else if (i == 10)
                        nPts[i] = (double)swTransform.ArrayData[i] + deltaY;// (double)swTransform.ArrayData[i]+ 0;
                    else if (i == 11)
                        nPts[i] = (double)swTransform.ArrayData[i] + deltaZ;
                    else
                        nPts[i] = (double)swTransform.ArrayData[i];

                }
                swRotTransform = swMathUtil.CreateTransform(nPts);
                boolstatus = swComp.SetTransformAndSolve2(swRotTransform);
                //boolstatus = swModel.ForceRebuild3(false);
                swModel.ClearSelection2(true);
                swTransform = swComp.Transform2;
                //MessageBox.Show(swTransform.ArrayData);
                swMathUtil = swApp.GetMathUtility();
                for (int i = 0; i < 16; i++)
                {
                    if (i == 11)
                    {
                        if (side)
                            nPts[i] = -0.001;       //if on top side keep moving down untill it hits the board
                        else
                            nPts[i] = 0.001;
                    }
                    else
                        nPts[i] = 0;
                }
                AssemblyDoc swAssy = default(AssemblyDoc);
                DragOperator swDragOp = default(DragOperator);
                swAssy = (AssemblyDoc)swModel;
                MathTransform swXform = default(MathTransform);
                swXform = swMathUtil.CreateTransform(nPts);
                swDragOp = (DragOperator)swAssy.GetDragOperator();
                bRet = swDragOp.AddComponent(swComp, false);
                swModel.ClearSelection2(true);
                boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 1, null, 0);
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                Component2[] Entity_Array = new Component2[1];
                Entity_Array[0] = (Component2)swSelMgr.GetSelectedObjectsComponent4(1, -1);
                //swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);

                bRet = swDragOp.CollisionDetection(Entity_Array, false, true);

                swDragOp.CollisionDetectionEnabled = true;
                swDragOp.DynamicClearanceEnabled = false;

                // Translation Only
                swDragOp.TransformType = 0;

                // Solve by relaxation
                swDragOp.DragMode = 2;

                bRet = swDragOp.BeginDrag();

                for (int i = 0; i <= 10; i++)
                {
                    // Returns false if drag fails
                    bRet = swDragOp.Drag(swXform);
                    if (!bRet)
                        break;
                    // Wait for 0.01 secs
                    nNow = System.DateTime.Now;
                    while (System.DateTime.Now < nNow.AddSeconds(.01))
                    {
                        // Process event loop
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                bRet = swDragOp.EndDrag();
            }
        }

        private void RotateAllComponent(string componentPackage, double angle, int alongAxis, SldWorks mSolidworksApplication)
        {
            if (boardComponents.ContainsKey(componentPackage))
            {
                for (int i = 0; i < boardComponents[componentPackage].Count; i++)
                {
                    string modelID = (string)boardComponents[componentPackage][i].modelID;
                    if (modelID.Length > 2)
                    {
                        //rotate in same direction if rotating along the vertical axis
                        //double[] ret = getPerpendicullarUnitVector(modelID, TaskpaneIntegration.mSolidworksApplication);

                        if (alongAxis==2)
                            if (!boardComponents[componentPackage][i].side)
                                angle = 360 - angle;
                        RotateComponentAlongAxis(modelID, angle, alongAxis, TaskpaneIntegration.mSolidworksApplication);
                    }
                }
            }
        }

        private void TranslateAllComponent(string componentPackage, double[] distance, bool detectCollision, SldWorks mSolidworksApplication)
        {
            if (boardComponents.ContainsKey(componentPackage))
            {
                for (int i = 0; i < boardComponents[componentPackage].Count; i++)
                {
                    string modelID = (string)boardComponents[componentPackage][i].modelID;
                    if (modelID.Length > 2)
                    {
                        translateComponentDragOperator(modelID, distance, detectCollision, TaskpaneIntegration.mSolidworksApplication);
                    }
                }
            }
        }

        private void AutoAlignPackage(string componentPackage, bool isSMD, bool isSelected, SldWorks mSolidworksApplication)
        {
            if (isSelected)
            {
                SelectionMgr swSelMgr = default(SelectionMgr);
                SldWorks swApp = TaskpaneIntegration.mSolidworksApplication;
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                string AssemblyTitle = swModel.GetTitle();
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                Component2 swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
                if (swComp != null)
                {
                    //MessageBox.Show(swComp.Name2);
                    string selectedName = swComp.Name2 + "@" + AssemblyTitle;
                    ComponentAttribute component = null;
                    Parallel.ForEach(boardComponents, (c, loopState) =>
                    {
                        for (int i = 0; i < c.Value.Count; i++)
                        {
                            if (c.Value[i].modelID == selectedName)
                            {
                                //MessageBox.Show(c.Key);
                                component = c.Value[i];
                                loopState.Stop();
                                break;
                            }
                        }
                    });

                    PointF pointFs = component.location;
                    bool side = component.side;
                    if (isSMD)
                        AutoAlignComponentSMD(selectedName, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, TaskpaneIntegration.mSolidworksApplication);
                    else
                        AutoAlignComponent(selectedName, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, TaskpaneIntegration.mSolidworksApplication);
                }
            }
            else
            {
                if (boardComponents.ContainsKey(componentPackage))
                {
                    List<ComponentAttribute> c = boardComponents[componentPackage];
                    for (int i = 0; i < c.Count; i++)
                    {
                        string modelID = (string)boardComponents[componentPackage][i].modelID;
                        if (modelID.Length > 2)
                        {
                            PointF pointFs = c[i].location;
                            bool side = c[i].side;
                            if (isSMD)
                                AutoAlignComponentSMD(modelID, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, TaskpaneIntegration.mSolidworksApplication);
                            else
                                AutoAlignComponent(modelID, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, TaskpaneIntegration.mSolidworksApplication);
                        }
                    }
                }
            }
        }

        private void InsertNoteForPackage(string componentPackage, SldWorks mSolidworksApplication)
        {
            if (boardComponents.ContainsKey(componentPackage))
            {
                List<ComponentAttribute> c = boardComponents[componentPackage];
                for (int i = 0; i < CurrentNotesTop.Count; i++)
                {
                    removeNote(CurrentNotesTop[i], mSolidworksApplication);
                }
                CurrentNotesTop.Clear();
                for (int i = 0; i < CurrentNotesBottom.Count; i++)
                {
                    removeNote(CurrentNotesBottom[i], mSolidworksApplication);
                }
                CurrentNotesBottom.Clear();
                for (int i = 0; i < c.Count; i++)
                {
                    PointF pointFs = c[i].location;
                    bool side = c[i].side;
                    double X = convertUnit(c[i].location.X) + boardOrigin[0];
                    double Y = convertUnit(c[i].location.Y) + boardOrigin[1];
                    double Z = 0;
                    if (side)
                        Z = boardOrigin[2] + boardHeight;
                    else
                        Z = boardOrigin[2];
                    if (checkBoxTopOnly.Checked)
                    {
                        if (side)
                        {
                            string ret = insertNote(componentPackage, X, Y, Z, side, mSolidworksApplication);
                            CurrentNotesTop.Add(ret);
                            boardComponents[componentPackage][i].noteID = ret;
                        }
                    }
                    if (checkBoxBottomOnly.Checked)
                    {
                        if (!side)
                        {
                            string ret = insertNote(componentPackage, X, Y, Z, side, mSolidworksApplication);
                            CurrentNotesBottom.Add(ret);
                            boardComponents[componentPackage][i].noteID = ret;
                        }
                    }

                }
            }
        }
        private void RotateAllSelectedComponent(double angle, int alongAxis, SldWorks mSolidworksApplication)
        {
            bool boolstatus;

            SelectionMgr swSelMgr = default(SelectionMgr);
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            int numberOfSelected = swSelMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= numberOfSelected; i++)
            {
                Component2 swComp = (Component2)swSelMgr.GetSelectedObject6(i, -1);
                if (swComp != null)
                {
                    MathTransform swTransform = default(MathTransform);
                    MathTransform swRotTransform = default(MathTransform);
                    double[] nPts = new double[3];
                    object vData;
                    MathPoint swOriginPt;
                    MathVector sw_Axis;
                    swTransform = swComp.Transform2;
                    //MessageBox.Show(swTransform.ArrayData);
                    MathUtility swMathUtil;
                    swTransform = swComp.Transform2;
                    //do the rotation
                    double[] unitVector = null;
                    if (alongAxis == 0)
                        unitVector = getParallelUnitVector(swTransform.ArrayData, true);
                    else if (alongAxis == 1)
                        unitVector = getParallelUnitVector(swTransform.ArrayData, false);
                    else if (alongAxis == 2)
                        unitVector = getPerpendicullarUnitVector(swTransform.ArrayData);

                    swMathUtil = swApp.GetMathUtility();
                    double[] BoxArray = (double[])swComp.GetBox(false, false); ;
                    nPts[0] = (BoxArray[0] + BoxArray[3]) / 2;
                    nPts[1] = (BoxArray[1] + BoxArray[4]) / 2;
                    nPts[2] = (BoxArray[2] + BoxArray[5]) / 2;
                    vData = nPts;
                    swOriginPt = swMathUtil.CreatePoint(vData);

                    if (closeEnough(Math.Abs(unitVector[0]), 1))
                    {
                        nPts[0] = (double)swTransform.ArrayData[0];
                        nPts[1] = (double)swTransform.ArrayData[1];
                        nPts[2] = (double)swTransform.ArrayData[2];
                    }
                    else if (closeEnough(Math.Abs(unitVector[1]), 1))
                    {
                        nPts[0] = (double)swTransform.ArrayData[3];
                        nPts[1] = (double)swTransform.ArrayData[4];
                        nPts[2] = (double)swTransform.ArrayData[5];
                    }
                    else if (closeEnough(Math.Abs(unitVector[2]), 1))
                    {
                        nPts[0] = (double)swTransform.ArrayData[6];
                        nPts[1] = (double)swTransform.ArrayData[7];
                        nPts[2] = (double)swTransform.ArrayData[8];
                    }
                    else
                    {
                        if (alongAxis == 0)
                        {
                            nPts[0] = (double)swTransform.ArrayData[0];
                            nPts[1] = (double)swTransform.ArrayData[1];
                            nPts[2] = (double)swTransform.ArrayData[2];
                        }
                        else if (alongAxis == 1)
                        {
                            nPts[0] = (double)swTransform.ArrayData[3];
                            nPts[1] = (double)swTransform.ArrayData[4];
                            nPts[2] = (double)swTransform.ArrayData[5];
                        }
                        else if (alongAxis == 2)
                        {
                            nPts[0] = (double)swTransform.ArrayData[6];
                            nPts[1] = (double)swTransform.ArrayData[7];
                            nPts[2] = (double)swTransform.ArrayData[8];
                        }
                    }
                    vData = nPts;
                    sw_Axis = swMathUtil.CreateVector(vData);

                    swRotTransform = swMathUtil.CreateTransformRotateAxis(swOriginPt, sw_Axis, angle * RadPerDeg);
                    swTransform = swTransform.Multiply(swRotTransform);
                    boolstatus = swComp.SetTransformAndSolve2(swTransform);
                }
            }
        }
        private void moveLibraryComponent(string componentID, double[] libraryOffsets, PointF location, SldWorks mSolidworksApplication)
        {
            //AssemblyDoc swAssy;
            bool boolstatus;
            //long longstatus;
            //long longwarnings;
            Component2 swComp;
            //object[] vComponents = null;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);
            double[] rotationCenter = new double[3];
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            //swAssy = (AssemblyDoc)swModel;
            // vComponents = (object[])swAssy.GetComponents(true);

            //swComp = (Component2)vComponents[1];
            if (swComp != null)
            {
                //MessageBox.Show(swTransform.ArrayData);
                MathUtility swMathUtil;
                swMathUtil = swApp.GetMathUtility();
                double[] BoxArray = (double[])swComp.GetBox(false, false); ;
                rotationCenter[0] = (BoxArray[0] + BoxArray[3]) / 2;
                rotationCenter[1] = (BoxArray[1] + BoxArray[4]) / 2;
                rotationCenter[2] = (BoxArray[2] + BoxArray[5]) / 2;

                double[] finalLocation = new double[3] { 0, 0, 0 };
                finalLocation[0] = boardOrigin[0] + (convertUnit(location.X) - rotationCenter[0]) + libraryOffsets[0];
                finalLocation[1] = boardOrigin[1] + (convertUnit(location.Y) - rotationCenter[1]) + libraryOffsets[1];
                swTransform = swComp.Transform2;
                //MessageBox.Show(swTransform.ArrayData);
                swMathUtil = swApp.GetMathUtility();
                double[] nPts = new double[16] { 1, 0, 0, 0, 1, 0, 0, 0, 1, finalLocation[0], finalLocation[1], 0, 1.0, 0.0, 0.0, 0.0 };
                swRotTransform = swMathUtil.CreateTransform(nPts);
                swTransform = swTransform.Multiply(swRotTransform);
                boolstatus = swComp.SetTransformAndSolve2(swTransform);
            }
        }
        private void RotateComponentAlongAxis(string componentID, double angle, int alongAxis, SldWorks mSolidworksApplication)
        {
            //AssemblyDoc swAssy;
            bool boolstatus;
            //long longstatus;
            //long longwarnings;
            Component2 swComp;
            //object[] vComponents = null;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);

            object vData;
            MathPoint swOriginPt;
            MathVector sw_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            //swAssy = (AssemblyDoc)swModel;
            // vComponents = (object[])swAssy.GetComponents(true);

            //swComp = (Component2)vComponents[1];
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                MathUtility swMathUtil;
                //do the rotation
                double[] unitVector = null;
                if (alongAxis == 0)
                    unitVector = getParallelUnitVector(swTransform.ArrayData, true);
                else if (alongAxis == 1)
                    unitVector = getParallelUnitVector(swTransform.ArrayData, false);
                else if (alongAxis == 2)
                    unitVector = getPerpendicullarUnitVector(swTransform.ArrayData);
                //if (closeEnough(Math.Abs(unitVector[0]), 1))
                //    alongAxis = 0;
                //else if(closeEnough(Math.Abs(unitVector[1]), 1))
                //    alongAxis = 1;
                //else if(closeEnough(Math.Abs(unitVector[2]), 1))
                //    alongAxis = 2;
                swMathUtil = swApp.GetMathUtility();
                double[] BoxArray = (double[])swComp.GetBox(false, false); ;
                double[] nPts = new double[3];
                nPts[0] = (BoxArray[0] + BoxArray[3]) / 2;
                nPts[1] = (BoxArray[1] + BoxArray[4]) / 2;
                nPts[2] = (BoxArray[2] + BoxArray[5]) / 2;
                vData = nPts;
                swOriginPt = swMathUtil.CreatePoint(vData);

                if (closeEnough(Math.Abs(unitVector[0]), 1))
                {
                    nPts[0] = (double)swTransform.ArrayData[0];
                    nPts[1] = (double)swTransform.ArrayData[1];
                    nPts[2] = (double)swTransform.ArrayData[2];
                }
                else if (closeEnough(Math.Abs(unitVector[1]), 1))
                {
                    nPts[0] = (double)swTransform.ArrayData[3];
                    nPts[1] = (double)swTransform.ArrayData[4];
                    nPts[2] = (double)swTransform.ArrayData[5];
                }
                else if (closeEnough(Math.Abs(unitVector[2]), 1))
                {
                    nPts[0] = (double)swTransform.ArrayData[6];
                    nPts[1] = (double)swTransform.ArrayData[7];
                    nPts[2] = (double)swTransform.ArrayData[8];
                }
                else
                {
                    if (alongAxis == 0)
                    {
                        nPts[0] = (double)swTransform.ArrayData[0];
                        nPts[1] = (double)swTransform.ArrayData[1];
                        nPts[2] = (double)swTransform.ArrayData[2];
                    }
                    else if (alongAxis == 1)
                    {
                        nPts[0] = (double)swTransform.ArrayData[3];
                        nPts[1] = (double)swTransform.ArrayData[4];
                        nPts[2] = (double)swTransform.ArrayData[5];
                    }
                    else if (alongAxis == 2)
                    {
                        nPts[0] = (double)swTransform.ArrayData[6];
                        nPts[1] = (double)swTransform.ArrayData[7];
                        nPts[2] = (double)swTransform.ArrayData[8];
                    }
                }
                vData = nPts;
                sw_Axis = swMathUtil.CreateVector(vData);

                swRotTransform = swMathUtil.CreateTransformRotateAxis(swOriginPt, sw_Axis, angle * RadPerDeg);
                swTransform = swTransform.Multiply(swRotTransform);
                boolstatus = swComp.SetTransformAndSolve2(swTransform);
            }
            //swModel.ClearSelection2(true);

        }
        private void rotateComponent(string componentID, double angle, int alongAxis, SldWorks mSolidworksApplication)
        {
            //AssemblyDoc swAssy;
            bool boolstatus;
            //long longstatus;
            //long longwarnings;
            Component2 swComp;
            //object[] vComponents = null;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);
            double[] nPts = new double[3];
            object vData;
            MathPoint swOriginPt;
            MathVector sw_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            //swAssy = (AssemblyDoc)swModel;
            // vComponents = (object[])swAssy.GetComponents(true);

            //swComp = (Component2)vComponents[1];
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                //MessageBox.Show(swTransform.ArrayData);
                MathUtility swMathUtil;
                swMathUtil = swApp.GetMathUtility();
                double[] BoxArray = (double[])swComp.GetBox(false, false); ;
                nPts[0] = (BoxArray[0] + BoxArray[3]) / 2;
                nPts[1] = (BoxArray[1] + BoxArray[4]) / 2;
                nPts[2] = (BoxArray[2] + BoxArray[5]) / 2;
                vData = nPts;
                swOriginPt = swMathUtil.CreatePoint(vData);

                if (alongAxis == 0)
                {
                    nPts[0] = (double)swTransform.ArrayData[0];
                    nPts[1] = (double)swTransform.ArrayData[1];
                    nPts[2] = (double)swTransform.ArrayData[2];
                }
                else if (alongAxis == 1)
                {
                    nPts[0] = (double)swTransform.ArrayData[3];
                    nPts[1] = (double)swTransform.ArrayData[4];
                    nPts[2] = (double)swTransform.ArrayData[5];
                }
                else if (alongAxis == 2)
                {
                    nPts[0] = (double)swTransform.ArrayData[6];
                    nPts[1] = (double)swTransform.ArrayData[7];
                    nPts[2] = (double)swTransform.ArrayData[8];
                }
                vData = nPts;
                sw_Axis = swMathUtil.CreateVector(vData);

                swRotTransform = swMathUtil.CreateTransformRotateAxis(swOriginPt, sw_Axis, angle * RadPerDeg);
                swTransform = swTransform.Multiply(swRotTransform);
                boolstatus = swComp.SetTransformAndSolve2(swTransform);
                //boolstatus = swModel.ForceRebuild3(false);
            }
            //swModel.ClearSelection2(true);

        }
        private void translateComponent(double[] distance, SldWorks mSolidworksApplication)
        {
            //AssemblyDoc swAssy;
            bool boolstatus;
            //long longstatus;
            //long longwarnings;
            Component2 swComp;
            //object[] vComponents = null;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swRotTransform = default(MathTransform);
            double[] nPts = new double[16];
            //MathPoint swOriginPt;
            //MathVector swX_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;

            //swAssy = (AssemblyDoc)swModel;
            //vComponents = (object[])swAssy.GetComponents(true);

            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            swTransform = swComp.Transform2;
            //MessageBox.Show(swTransform.ArrayData);
            MathUtility swMathUtil;
            swMathUtil = swApp.GetMathUtility();
            for (int i = 0; i < 16; i++)
            {
                if (!(i == 9 || i == 10 || i == 11))
                    nPts[i] = (double)swTransform.ArrayData[i];
                else
                {
                    nPts[i] = distance[i - 9];
                }
            }
            swRotTransform = swMathUtil.CreateTransform(nPts);
            swTransform = swTransform.Multiply(swRotTransform);
            boolstatus = swComp.SetTransformAndSolve2(swTransform);
            //boolstatus = swModel.ForceRebuild3(false);
            swModel.ClearSelection2(true);

        }
        private void translateComponentDragOperator(string componentID, double[] distance, bool detectCollision, SldWorks mSolidworksApplication,bool forcemove=false)
        {
            //AssemblyDoc swAssy;
            bool boolstatus;
            //long longstatus;
            //long longwarnings;
            Component2 swComp;
            AssemblyDoc swAssy = default(AssemblyDoc);
            DragOperator swDragOp = default(DragOperator);
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swXform = default(MathTransform);
            double[] nPts = new double[16];

            bool bRet = false;
            //MathPoint swOriginPt;
            //MathVector swX_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;

            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                bool was_forced = false;
                swAssy = (AssemblyDoc)swModel;
                if (swComp.IsFixed())
                {
                    if (forcemove)
                    {
                        swAssy.UnfixComponent();
                        was_forced = true;
                    }
                }
                swTransform = swComp.Transform2;
                //MessageBox.Show(swTransform.ArrayData);
                MathUtility swMathUtil;
                swMathUtil = swApp.GetMathUtility();
                for (int i = 0; i < 16; i++)
                {
                    if (i == 9)
                        nPts[i] = distance[0];
                    else if (i == 10)
                        nPts[i] = distance[1];
                    else if (i == 11)
                        nPts[i] = distance[2];
                    else
                        nPts[i] = 0;
                }
                
                swXform = swMathUtil.CreateTransform(nPts);
                swDragOp = (DragOperator)swAssy.GetDragOperator();
                bRet = swDragOp.AddComponent(swComp, false);
                if (detectCollision)
                {
                    swModel.ClearSelection2(true);
                    boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 1, null, 0);
                    swSelMgr = (SelectionMgr)swModel.SelectionManager;
                    Component2[] Entity_Array = new Component2[1];
                    Entity_Array[0] = (Component2)swSelMgr.GetSelectedObjectsComponent4(1, -1);
                    //swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);

                    bRet = swDragOp.CollisionDetection(Entity_Array, false, true);
                    swDragOp.CollisionDetectionEnabled = true;
                }
                else
                    swDragOp.CollisionDetectionEnabled = false;
                swDragOp.DynamicClearanceEnabled = false;

                // Translation Only
                swDragOp.TransformType = 0;
                swDragOp.DragMode = 2;
                bRet = swDragOp.BeginDrag();
                bRet = swDragOp.Drag(swXform);

                bRet = swDragOp.EndDrag();
                swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                if (was_forced)
                {
                    swAssy.FixComponent();
                }
                swModel.ClearSelection2(true);
            }
        }
        private void translateComponentDragOperator(double[] distance, bool detectCollision, SldWorks mSolidworksApplication)
        {
            //AssemblyDoc swAssy;
            bool boolstatus;
            //long longstatus;
            //long longwarnings;
            Component2 swComp;
            AssemblyDoc swAssy = default(AssemblyDoc);
            DragOperator swDragOp = default(DragOperator);
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            MathTransform swXform = default(MathTransform);
            double[] nPts = new double[16];
            bool bRet = false;
            //MathPoint swOriginPt;
            //MathVector swX_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                //MessageBox.Show(swTransform.ArrayData);
                MathUtility swMathUtil;
                swMathUtil = swApp.GetMathUtility();
                for (int i = 0; i < 16; i++)
                {
                    if (i == 9)
                        nPts[i] = distance[0];
                    else if (i == 10)
                        nPts[i] = distance[1];
                    else if (i == 11)
                        nPts[i] = distance[2];
                    else
                        nPts[i] = 0;
                }
                swAssy = (AssemblyDoc)swModel;
                swXform = swMathUtil.CreateTransform(nPts);
                swDragOp = (DragOperator)swAssy.GetDragOperator();
                bRet = swDragOp.AddComponent(swComp, false);
                if (detectCollision)
                {
                    //swModel.ClearSelection2(true);
                    boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 1, null, 0);
                    swSelMgr = (SelectionMgr)swModel.SelectionManager;
                    Component2[] Entity_Array = new Component2[1];
                    Entity_Array[0] = (Component2)swSelMgr.GetSelectedObjectsComponent4(1, -1);
                    //swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);

                    bRet = swDragOp.CollisionDetection(Entity_Array, false, true);
                    swDragOp.CollisionDetectionEnabled = true;
                }
                else
                    swDragOp.CollisionDetectionEnabled = false;
                swDragOp.DynamicClearanceEnabled = false;

                // Translation Only
                swDragOp.TransformType = 0;

                // Solve by relaxation
                swDragOp.DragMode = 2;
                bRet = swDragOp.BeginDrag();
                bRet = swDragOp.Drag(swXform);
                bRet = swDragOp.EndDrag();
                swSelMgr.AddSelectionListObject(swComp, null);
            }
        }
        private double convertUnit(double input,bool reverse=false)
        {
            if(reverse)
                return input * 1000;
            else
                return input / 1000;
        }
        private PointF[] FindCircles(PointF p, PointF q, double radius)
        {
            //if (radius < 0) throw new ArgumentException("Negative radius.");
            //if (radius == 0)
            //{
            //    if (p == q) return new[] { p };
            //    else throw new InvalidOperationException("No circles.");
            //}
            //if (p == q) throw new InvalidOperationException("Infinite number of circles.");

            //double sqDistance = Math.Sqrt((p.X - q.X) * (p.X - q.X) + (p.Y - q.Y) * (p.Y - q.Y));
            //double sqDiameter = 4 * radius * radius;
            //if (sqDistance > sqDiameter) throw new InvalidOperationException("Points are too far apart.");

            //PointF midPoint = new PointF((p.X + q.X) / 2, (p.Y + q.Y) / 2);
            //if (sqDistance == sqDiameter) return new[] { midPoint };

            //double d = Math.Sqrt(radius * radius - sqDistance / 4);
            //double distance = Math.Sqrt(sqDistance);
            //double ox = d * (q.X - p.X) / distance, oy = d * (q.Y - p.Y) / distance;
            //return new[] {
            //new PointF((float)(midPoint.X - oy), (float)(midPoint.Y + ox)),
            //new PointF((float)(midPoint.X + oy), (float)(midPoint.Y - ox))
            //};
            //double c = p.X * p.X + p.Y * p.Y - q.X * q.X - q.Y * q.Y;
            //double x3 = p.X - q.X;
            //double y3 = p.Y - q.Y;
            //double cy = c / (2 * y3) - p.Y;
            //double a = (x3 * x3 + y3 * y3) / y3 * y3;
            //double b = 2* p.X - 2*cy*(x3/y3);
            //double c2 = p.X * p.X + cy * cy - radius * radius;
            //double X1 = (-b + Math.Sqrt(b * b - 4 * a * c2)) / 2 * a;
            //double X2 = (-b - Math.Sqrt(b * b - 4 * a * c2)) / 2 * a;
            //double Y1 = -(x3 / y3) * X1 + c / (2 * y3);
            //double Y2 = -(x3 / y3) * X2 + c / (2 * y3);
            double x1 = p.X;
            double x2 = q.X;
            double y1 = p.Y;
            double y2 = q.Y;
            double r = radius;
            //double ca = (x1 * x1) + (y1 * y1) - (x2 * x2) - (y2 * y2);
            //double x3 = x1 - x2;
            //double y3 = y1 - y2;
            //double cy = (ca / (2 * y3)) - y1;

            //double a = ((x3 * x3) + (y3 * y3)) / (y3 * y3);
            //double b = ((2 * x1) - (2 * (cy) * x3 / y3));
            //double c = ((x1 * x1) + (cy * cy) - (r * r));

            //double[] roots;
            //double denominator = 2 * a;
            //double bNumerator = -b;
            //double underSquare = (b * b) - (4 * a * c);
            //if (underSquare < 0 || denominator == 0)
            //{
            //    roots = new double[0];
            //}
            //else
            //{
            //    double sqrt = Math.Sqrt(underSquare);
            //    roots = new double[] { (bNumerator + sqrt) / denominator, (bNumerator - sqrt) / denominator };
            //}
            //double X1 = roots[0];
            //double X2 = roots[1];
            //double Y1 = -(x3 / y3) * X1 + c / (2 * y3);
            //double Y2 = -(x3 / y3) * X2 + c / (2 * y3);
            double c = (x1 * x1 + y1 * y1 - x2 * x2 - y2 * y2);
            double b = 2 * x2 - 2 * x1;
            double a = 2 * y1 - 2 * y2;

            double qa = 1 + ((a * a)/ (b * b));
            double qb = -(2 * x1 * a) / b - (2 * c * a) / (b * b) - 2 * y1;
            double qc = x1 * x1 + (c * c) / (b * b) + (2 * x1 * c) / b + y1 * y1 - r * r;
            double Y1 = (-qb + Math.Sqrt(qb * qb - 4 * qa * qc)) /( 2 * qa);
            double Y2 = (-qb - Math.Sqrt(qb * qb - 4 * qa * qc)) /( 2 * qa);
            double X1 = (a * Y1 - c) / b;
            double X2 = (a * Y2 - c) / b;
            return new[] {
            new PointF((float)(X1), (float)(Y1)),
            new PointF((float)(X2), (float)(Y2))
            };
        }
        private double normalizeAngle(double angle)
        {
            double arcAngle = (angle * 180) / Math.PI;
            arcAngle = arcAngle % 360;
            if (arcAngle < 0)
                arcAngle = 360 + arcAngle;
            return arcAngle;
        }
        private PointF convertArc(PointF P1,PointF P2, double arcAngle)
        {
            double angle= Math.Abs(arcAngle);
            //if (arcAngle < 0)
            //    angle = 360 + arcAngle;
            angle = (angle * Math.PI) / 180;
            double dist = Math.Sqrt((P1.X - P2.X) * (P1.X - P2.X) + (P1.Y - P2.Y) * (P1.Y - P2.Y));
            double radius = (dist / 2) / (Math.Sin(angle / 2));
            PointF[] centers = FindCircles(P1, P2, radius);
            //now figureout which center shd we choose
            double tP1 = Math.Atan2(P1.Y - centers[0].Y, P1.X - centers[0].X);
            double tP2 = Math.Atan2(P2.Y - centers[0].Y, P2.X - centers[0].X);
            double t1 = normalizeAngle(tP2 - tP1);
            if (closeEnough(t1, Math.Abs(arcAngle)))
            {
                if(arcAngle > 0)
                    return new PointF(centers[0].X, centers[0].Y);
                else
                    return new PointF(centers[1].X, centers[1].Y);
            }
                
            tP1 = Math.Atan2(P1.Y - centers[1].Y, P1.X - centers[1].X);
            tP2 = Math.Atan2(P2.Y - centers[1].Y, P2.X - centers[1].X);
            double t2 = normalizeAngle(tP2 - tP1);
            //if (closeEnough(t2, arcAngle))
            //return new PointF(centers[1].X, centers[1].Y);
            //return new PointF[] { centers[0], centers[1] };
            if (arcAngle > 0)
                return new PointF(centers[1].X, centers[1].Y);
            else
                return new PointF(centers[0].X, centers[0].Y);
        }
        private string createBoardFromIDF(string emnfile, SldWorks mSolidworksApplication)
        {
            
            List<List<IDFBoardOutline>> IDFboardData = new List<List<IDFBoardOutline>> { new List<IDFBoardOutline> { } };
            double boardThickness = 0;
            const Int32 BufferSize = 2048;
            openedFile = emnfile;
            using (var fileStream = File.OpenRead(emnfile))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;
                List<IDFBoardOutline> currentOutline = new List<IDFBoardOutline> { };
                bool startProcessing = false;
                int previndex = -1;
                int curIndex = -1;
                while ((line = streamReader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (!startProcessing && line.IndexOf(".BOARD_OUTLINE")!=-1)
                    {
                        startProcessing = true;
                        line = streamReader.ReadLine();
                        boardThickness =convertUnit(Convert.ToDouble(line));
                        IDFboardData.Clear();
                        continue;
                    }
                    if (startProcessing && line.IndexOf(".END_BOARD_OUTLINE") != -1)
                    {
                        IDFboardData.Add(new List<IDFBoardOutline>(currentOutline));
                        break;
                    }
                    if (startProcessing)
                    {
                        if (line.Length > 2)
                        {
                            //List<string> lineElements = line.Split(' ').ToList();
                            List<string> lineElements = splitString(line);
                            List<double> lineElementsDouble = new List<double> { };
                            lineElementsDouble.Clear();
                            for (int i = 0; i < lineElements.Count; i++)
                            {
                                lineElements[i] = lineElements[i].Trim();
                                if (lineElements[i].Length > 0)
                                {
                                    lineElementsDouble.Add(Convert.ToDouble(lineElements[i]));
                                }
                            }
                            curIndex = (int)lineElementsDouble[0];
                            if (curIndex!= previndex)
                            {
                                if(currentOutline.Count>0)
                                {
                                    IDFboardData.Add(new List<IDFBoardOutline>(currentOutline));
                                }
                                currentOutline.Clear();
                            }
                            IDFBoardOutline outline = new IDFBoardOutline();
                            outline.location.X= (float)Convert.ToDouble(lineElementsDouble[1]);
                            outline.location.Y = (float)Convert.ToDouble(lineElementsDouble[2]);
                            outline.rotation = (float)Convert.ToDouble(lineElementsDouble[3]);
                            currentOutline.Add(outline);
                            previndex = curIndex;
                        }
                    }
                }
            }
            bool boolstatus = false;
            mSolidworksApplication.NewPart();
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            string ModelTitle = swModel.GetTitle();
            //swModel.ShowNamedView2("*Front", 1);
            boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            Sketch activeSketch = swModel.SketchManager.ActiveSketch;
            Feature feature = (Feature)activeSketch;
            SketchSegment skSegment;
            skSegment = null;

            swModel.SketchManager.AddToDB = true;
            boardHeight = boardThickness;
            numericUpDownBoardHeight.Value = (decimal)convertUnit(boardThickness,true);
            List<IDFBoardOutline> vertices = IDFboardData[0];
            for (int i = 1; i < vertices.Count; i++)
            {
                if (i == 1)
                {
                    boardFirstPoint[0] = convertUnit((double)vertices[i - 1].location.X);
                    boardFirstPoint[1] = convertUnit((double)vertices[i - 1].location.Y);
                }
                if (Math.Abs(vertices[i].rotation) <= 0.0001)
                    skSegment = swModel.SketchManager.CreateLine(convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0);
                else
                {
                    PointF centers = convertArc(vertices[i - 1].location, vertices[i].location, vertices[i].rotation);
                    if (vertices[i].rotation == 360)
                    {
                        skSegment = swModel.SketchManager.CreateCircle(convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0);
                    }
                    else
                    {
                        if (vertices[i].rotation > 0)
                            skSegment = swModel.SketchManager.CreateArc(convertUnit(centers.X), convertUnit(centers.Y), 0, convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0, 1);
                        else
                            skSegment = swModel.SketchManager.CreateArc(convertUnit(centers.X), convertUnit(centers.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0, convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, 1);
                    }
                }
            }

            for (int j = 1; j < IDFboardData.Count; j++)
            {

                //textBox1.Text += "\r\n" + currentFeature.Outline.Polycurve_Area.Vertices;
                vertices = IDFboardData[j];
                for (int i = 1; i < vertices.Count; i++)
                {
                    if (Math.Abs(vertices[i].rotation) <= 0.0001)
                        skSegment = swModel.SketchManager.CreateLine(convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0);
                    else
                    {
                        PointF centers = convertArc(vertices[i - 1].location, vertices[i].location, vertices[i].rotation);
                        if (vertices[i].rotation == 360)
                        {
                            skSegment = swModel.SketchManager.CreateCircle(convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0);
                        }
                        else
                        {
                            if (vertices[i].rotation > 0)
                                skSegment = swModel.SketchManager.CreateArc(convertUnit(centers.X), convertUnit(centers.Y), 0, convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0, 1);
                            else
                                skSegment = swModel.SketchManager.CreateArc(convertUnit(centers.X), convertUnit(centers.Y), 0, convertUnit((double)vertices[i].location.X), convertUnit((double)vertices[i].location.Y), 0, convertUnit((double)vertices[i - 1].location.X), convertUnit((double)vertices[i - 1].location.Y), 0, 1);
                        }
                    }
                }
            }
            swModel.SketchManager.AddToDB = false;


            boolstatus = swModel.EditRebuild3();
            swModel.ClearSelection2(true);
            boolstatus = swModel.Extension.SelectByID2(feature.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
            Feature extrude = swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.000001, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            extrude.Name = "BOARD_OUTLINE";
            boardSketch = extrude.Name;
            boolstatus = swModel.Extension.SelectByID2(boardSketch, "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            // Get the sketch feature
            Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
            swExtrudeData.BothDirections = true;
            swExtrudeData.SetDepth(true, boardHeight / 2);
            swExtrudeData.SetDepth(false, boardHeight / 2);
            swFeature.ModifyDefinition(swExtrudeData, swModel, null);
            swExtrudeData.ReleaseSelectionAccess();
            swModel.ClearSelection2(true);
            // Set the color of the sketch
            double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
            swMaterialPropertyValues[0] = (double)((double)lblBoardColor.BackColor.R / 255);
            swMaterialPropertyValues[1] = (double)((double)lblBoardColor.BackColor.G / 255);
            swMaterialPropertyValues[2] = (double)((double)lblBoardColor.BackColor.B / 255);
            swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swThisConfiguration, "");

            swModel.SelectionManager.EnableContourSelection = false;
            string directoryPath = Path.GetDirectoryName(emnfile);
            string boardPath = Path.Combine(directoryPath, "board.SLDPRT");
            long status = swModel.SaveAs3(boardPath, 0, 2);
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
            return boardPath;
        }
        private string createBoardFromJSON(string JSONfile, SldWorks mSolidworksApplication)
        {
            bool boolstatus = false;
            readBoardData(JSONfile);
            mSolidworksApplication.NewPart();
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            string ModelTitle = swModel.GetTitle();
            //swModel.ShowNamedView2("*Front", 1);
            boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            Sketch activeSketch = swModel.SketchManager.ActiveSketch;
            Feature feature = (Feature)activeSketch;
            SketchSegment skSegment;
            skSegment = null;

            swModel.SketchManager.AddToDB = true;
            boardHeight = convertUnit((double)boardJSON.Parts.Board_Part.Shape.Extrusion.Top_Height);
            numericUpDownBoardHeight.Value = (decimal)boardJSON.Parts.Board_Part.Shape.Extrusion.Top_Height;
            JArray vertices = (JArray)boardJSON.Parts.Board_Part.Shape.Extrusion.Outline.Polycurve_Area.Vertices;
            for (int i = 1; i < vertices.Count; i++)
            {
                if (i == 1)
                {
                    boardFirstPoint[0] = convertUnit((double)vertices[i - 1][0].ToObject(typeof(double)));
                    boardFirstPoint[1] = convertUnit((double)vertices[i - 1][1].ToObject(typeof(double)));
                }
                skSegment = swModel.SketchManager.CreateLine(convertUnit((double)vertices[i - 1][0].ToObject(typeof(double))), convertUnit((double)vertices[i - 1][1].ToObject(typeof(double))), 0, convertUnit((double)vertices[i][0].ToObject(typeof(double))), convertUnit((double)vertices[i][1].ToObject(typeof(double))), 0);
            }
            dynamic features = boardJSON.Parts.Board_Part.Features;
            for (int j = 0; j < features.Count; j++)
            {
                dynamic currentFeature = features[j];
                string featureType = (string)currentFeature.Feature_Type;
                if (featureType != null)
                    if (featureType.ToUpper() == "CUTOUT")
                    {
                        //textBox1.Text += "\r\n" + currentFeature.Outline.Polycurve_Area.Vertices;
                        vertices = currentFeature.Outline.Polycurve_Area.Vertices;
                        for (int i = 1; i < vertices.Count; i++)
                        {
                            skSegment = swModel.SketchManager.CreateLine(convertUnit((double)vertices[i - 1][0].ToObject(typeof(double))), convertUnit((double)vertices[i - 1][1].ToObject(typeof(double))), 0, convertUnit((double)vertices[i][0].ToObject(typeof(double))), convertUnit((double)vertices[i][1].ToObject(typeof(double))), 0);
                        }
                    }
            }
            swModel.SketchManager.AddToDB = false;


            boolstatus = swModel.EditRebuild3();
            swModel.ClearSelection2(true);

            boolstatus = swModel.Extension.SelectByID2(feature.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
            //swModel.ShowNamedView2("*Trimetric", 8);
            //boolstatus = swModel.Extension.SelectByID2("Sketch2", "SKETCH", 0, 0, 0, False, 4, Nothing, 0)
            Feature extrude = swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.000001, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            extrude.Name = "BOARD_OUTLINE";
            boardSketch = extrude.Name;
            boolstatus = swModel.Extension.SelectByID2(boardSketch, "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            // Get the sketch feature
            Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
            swExtrudeData.BothDirections = true;
            swExtrudeData.SetDepth(true, boardHeight / 2);
            swExtrudeData.SetDepth(false, boardHeight / 2);
            swFeature.ModifyDefinition(swExtrudeData, swModel, null);
            swExtrudeData.ReleaseSelectionAccess();
            swModel.ClearSelection2(true);
            // Set the color of the sketch
            double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
            swMaterialPropertyValues[0] = (double)((double)lblBoardColor.BackColor.R / 255);
            swMaterialPropertyValues[1] = (double)((double)lblBoardColor.BackColor.G / 255);
            swMaterialPropertyValues[2] = (double)((double)lblBoardColor.BackColor.B / 255);
            swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swThisConfiguration, "");

            swModel.SelectionManager.EnableContourSelection = false;
            string directoryPath = Path.GetDirectoryName(JSONfile);
            string boardPath = Path.Combine(directoryPath, "board.SLDPRT");
            //swModel.ViewZoomtofit2();
            //swModel.ShowNamedView2("*Front", 1);
            long status = swModel.SaveAs3(boardPath, 0, 2);
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
            return boardPath;
        }

        private void makeHolesFromIDF(string emnfile, string strpath, SldWorks mSolidworksApplication)
        {

            int status = 0;
            int warnings = 0;
            bool boolstatus = false;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            swAssy.OpenCompFile();
            swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            string ModelTitle = swModel.GetTitle();
            boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            Sketch activeSketch = swModel.SketchManager.ActiveSketch;
            Feature feature = (Feature)activeSketch;
            SketchSegment skSegment;
            skSegment = null;
            swModel.SketchManager.AddToDB = true;

            const Int32 BufferSize = 2048;
            using (var fileStream = File.OpenRead(emnfile))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;
                List<IDFBoardOutline> currentOutline = new List<IDFBoardOutline> { };
                bool startProcessing = false;
                while ((line = streamReader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (!startProcessing && line.IndexOf(".DRILLED_HOLES") != -1)
                    {
                        startProcessing = true;
                        continue;
                    }
                    if (startProcessing && line.IndexOf(".END_DRILLED_HOLES") != -1)
                    {
                        break;
                    }
                    if (startProcessing)
                    {
                        if (line.Length > 2)
                        {
                            //List<string> lineElements = line.Split(' ').ToList();
                            List<string> lineElements = splitString(line);
                            List<double> lineElementsDouble = new List<double> { };
                            lineElementsDouble.Clear();
                            int k = 0;
                            for (int i = 0; i < lineElements.Count; i++)
                            {
                                lineElements[i] = lineElements[i].Trim();
                                if (lineElements[i].Length > 0)
                                {
                                    lineElementsDouble.Add(Convert.ToDouble(lineElements[i]));
                                    k++;
                                    if (k >= 3)
                                        break;
                                }
                            }
                            double radius = lineElementsDouble[0] / 2;
                            skSegment = swModel.SketchManager.CreateCircleByRadius(convertUnit(lineElementsDouble[1]), convertUnit(lineElementsDouble[2]), 0, convertUnit(radius));
                        }
                    }
                }
            }

            swModel.SketchManager.AddToDB = false;
            boolstatus = swModel.EditRebuild3();
            swModel.ClearSelection2(true);
            boolstatus = swModel.Extension.SelectByID2(feature.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
            Feature extrude = swModel.FeatureManager.FeatureCut3(false, false, false, 0, 0, boardHeight * 2, boardHeight * 2, false, false, false, false, 0, 0,
                                                                    false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            //Feature extrude = swModel.FeatureManager.FeatureCut4(false, false, false, 0, 0, boardHeight * 2, boardHeight * 2, false, false, false, false, 0, 0, 
            //                                                        false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);        SW > 2017
            if (extrude != null)
            {
                extrude.Name = "BOARD_HOLES";
                boardSketch = extrude.Name;
                boolstatus = swModel.Extension.SelectByID2(boardSketch, "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
                // Get the sketch feature
                Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
                swExtrudeData.BothDirections = true;
                swExtrudeData.SetDepth(true, boardHeight * 0.55);
                swExtrudeData.SetDepth(false, boardHeight * 0.55);
                swFeature.ModifyDefinition(swExtrudeData, swModel, null);
                swExtrudeData.ReleaseSelectionAccess();
                swModel.ClearSelection2(true);
                // Set the color of the sketch
                double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
                swMaterialPropertyValues[0] = (double)((double)lblHolesColor.BackColor.R / 255);
                swMaterialPropertyValues[1] = (double)((double)lblHolesColor.BackColor.G / 255);
                swMaterialPropertyValues[2] = (double)((double)lblHolesColor.BackColor.B / 255);
                swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            }
            swModel.SelectionManager.EnableContourSelection = false;
            string boardPath = strpath;
            //swModel.ViewZoomtofit2();
            //swModel.ShowNamedView2("*Front", 1);
            status = swModel.SaveAs3(boardPath, 0, 2);
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
        }

        private string CalculateMD5Hash(string input)

        {
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
            byte[] hash = md5.ComputeHash(inputBytes);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
                sb.Append(hash[i].ToString("X2"));
            return sb.ToString();
        }
        private List<string> splitString(string line)
        {
            List<string> lineElements = line.Split('"')
                     .Select((element, index) => index % 2 == 0  // If even index
                                           ? element.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)  // Split the item
                                           : new string[] { element })  // Keep the entire item
                     .SelectMany(element => element).ToList();
            List<string> cleanLineElements = line.Split(' ').ToList();
            cleanLineElements.Clear();
            for (int i = 0; i < lineElements.Count; i++)
            {
                lineElements[i] = lineElements[i].Trim();
                if (lineElements[i].Length > 0)
                {
                    cleanLineElements.Add(lineElements[i]);
                }
            }
            return cleanLineElements;
        }
        private void readComponentFromIDF(string emnfile, SldWorks mSolidworksApplication)
        {
            boardComponents.Clear();
            ComponentTransformations.Clear();
            ComponentModels.Clear();
            const Int32 BufferSize = 2048;
            using (var fileStream = File.OpenRead(emnfile))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;
                List<IDFBoardOutline> currentOutline = new List<IDFBoardOutline> { };
                bool startProcessing = false;
                string packageName = "";
                string redef = "";
                string partNumber = "";
                while ((line = streamReader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (!startProcessing && line.IndexOf(".PLACEMENT") != -1)
                    {
                        startProcessing = true;
                        continue;
                    }
                    if (startProcessing && line.IndexOf(".END_PLACEMENT") != -1)
                    {
                        break;
                    }
                    if (startProcessing)
                    {
                        if (line.Length > 2)
                        {
                            List<string> cleanLineElements = splitString(line);
                            if(cleanLineElements.Count>3)
                            {
                                //it's placement data for the package
                                ComponentAttribute c = new ComponentAttribute();
                                c.location = new PointF(Convert.ToSingle(cleanLineElements[0]), Convert.ToSingle(cleanLineElements[1]));
                                c.rotation = Convert.ToSingle(cleanLineElements[3]);
                                if (cleanLineElements[4].ToUpper().IndexOf("BOTTOM") == -1)
                                    c.side = true;
                                else
                                    c.side = false;
                                c.redef = redef;
                                c.part_number = partNumber;
                                c.modelID = "";
                                c.modelPath = "";
                                c.componentID = CalculateMD5Hash(cleanLineElements[0]+ cleanLineElements[0]+ packageName+ cleanLineElements[4].ToUpper()+ RandomString(10));         //May be we can generate an ID as IDF doesn't do that for us
                                ComponentTransformations.TryAdd(c.componentID, new double[] { });
                                if (boardComponents.ContainsKey(packageName))
                                    boardComponents[packageName].Add(c);
                                else
                                    boardComponents.TryAdd(packageName, new List<ComponentAttribute> { c });
                                if (!ComponentModels.ContainsKey(packageName))
                                    ComponentModels.TryAdd(packageName, "");
                            }
                            else
                            {
                                //it's package name
                                packageName = cleanLineElements[0];
                                redef = cleanLineElements[2];
                                partNumber = cleanLineElements[1];
                            }
                        }
                    }
                }
            }
        }
        private void readComponents(dynamic boardJSON)
        {
            dynamic packages = boardJSON.Parts.Electrical_Part;
            int packages_count = (int)packages.Count;
            if (packages_count > 0)
            {
                ComponentModels.Clear();
                for (int i = 0; i < packages_count; i++)
                {
                    dynamic currentPackage = packages[i];
                    if (currentPackage.Properties.Part_Model != null)
                    {
                        string Part_Model = "";
                        if (File.Exists((string)currentPackage.Properties.Part_Model))
                            Part_Model = (string)currentPackage.Properties.Part_Model;
                        if (ComponentModels.ContainsKey((string)currentPackage.Part_Name))
                            ComponentModels[(string)currentPackage.Part_Name] = Part_Model;
                        else
                            ComponentModels.TryAdd((string)currentPackage.Part_Name, Part_Model);
                    }
                }
            }
            dynamic components = boardJSON.Assemblies.Board_Assembly.Comp_Insts;
            int component_count = (int)components.Count;
            if (component_count > 0)
            {
                boardComponents.Clear();
                for (int i = 0; i < component_count; i++)
                {
                    dynamic currentComponent = components[i];
                    if (currentComponent.XY_Loc != null)
                    {
                        JArray componentLoc = currentComponent.XY_Loc;

                        string componentPackage = (string)currentComponent.Part_Name;
                        ComponentAttribute c = new ComponentAttribute
                        {
                            location = new PointF((float)componentLoc[0].ToObject(typeof(float)), (float)componentLoc[1].ToObject(typeof(float)))
                        };
                        string side = (string)currentComponent.Side.ToObject(typeof(string));
                        c.side = (side.ToUpper() == "TOP") ? true : false;
                        c.rotation = (float)currentComponent.Rotation.ToObject(typeof(float));
                        c.modelID = "";
                        c.modelPath = "";
                        c.componentID = (string)currentComponent.Entity_ID.ToObject(typeof(string));
                        double[] tx;
                        if (currentComponent.Transformation != null)
                            tx = currentComponent.Transformation.ToObject<double[]>();
                        else
                            tx = new double[] { };
                        if (ComponentTransformations.ContainsKey(c.componentID))
                            ComponentTransformations[c.componentID] = tx;
                        else
                            ComponentTransformations.TryAdd(c.componentID, tx);

                        if (ComponentModels.ContainsKey(componentPackage))
                            if (File.Exists(ComponentModels[componentPackage]))
                                c.modelPath = ComponentModels[componentPackage];

                        if (boardComponents.ContainsKey(componentPackage))
                        {
                            boardComponents[componentPackage].Add(c);
                        }
                        else
                        {
                            boardComponents.TryAdd(componentPackage, new List<ComponentAttribute> { c });
                        }
                    }
                }
            }

        }

        private void ChangeSketchColor(ModelDoc2 swModel, string featureName, double red, double green, double blue, ref bool status, ref string errorMessage)
        {
            if (swModel == null)
                return;
            status = false;
            errorMessage = "";
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            string[] swModelConfigs = (string[])swModel.GetConfigurationNames();

            bool boolStatus;
            // Select the sketch
            boolStatus = swModel.Extension.SelectByID2(featureName, "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            if (!boolStatus)
            {
                status = false;
                errorMessage = "Unable to select sketch.";
            }
            else
            {
                // Get the sketch feature
                Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                // Clean up
                swModel.ClearSelection2(true);
                // Set the color of the sketch
                double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
                swMaterialPropertyValues[0] = red;
                swMaterialPropertyValues[1] = green;
                swMaterialPropertyValues[2] = blue;
                swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swAllConfiguration, swModelConfigs);
                status = true;
            }
        }
        private void makeHoles(string strpath, SldWorks mSolidworksApplication)
        {
            int status = 0;
            int warnings = 0;
            bool boolstatus = false;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            swAssy.OpenCompFile();
            swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            string ModelTitle = swModel.GetTitle();
            boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            Sketch activeSketch = swModel.SketchManager.ActiveSketch;
            Feature feature = (Feature)activeSketch;
            SketchSegment skSegment;
            skSegment = null;
            swModel.SketchManager.AddToDB = true;

            dynamic features = boardJSON.Parts.Board_Part.Features;
            for (int j = 0; j < features.Count; j++)
            {
                dynamic currentFeature = features[j];
                string featureType = (string)currentFeature.Feature_Type;
                if (featureType != null)
                {
                    if (featureType.ToUpper() == "HOLE")
                    {
                        if (currentFeature.Outline.Polycurve_Area != null)
                        {
                            //handle polycurve area
                        }
                        if (currentFeature.Outline.Circle != null)
                        {
                            JArray center = currentFeature.XY_Loc;
                            double radius = (double)currentFeature.Outline.Circle.Radius;
                            skSegment = swModel.SketchManager.CreateCircle(convertUnit((double)center[0].ToObject(typeof(double))), convertUnit((double)center[1].ToObject(typeof(double))), 0, convertUnit((double)center[0].ToObject(typeof(double))), convertUnit(radius + (double)center[1].ToObject(typeof(double))), 0);
                        }
                    }
                }
            }
            swModel.SketchManager.AddToDB = false;
            boolstatus = swModel.EditRebuild3();
            swModel.ClearSelection2(true);
            boolstatus = swModel.Extension.SelectByID2(feature.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
            Feature extrude = swModel.FeatureManager.FeatureCut3(false, false, false, 0, 0, boardHeight * 2, boardHeight * 2, false, false, false, false, 0, 0, 
                                                                    false, false, false, false, false, true, true, true, true, false, 0, 0, false); 
            //Feature extrude = swModel.FeatureManager.FeatureCut4(false, false, false, 0, 0, boardHeight * 2, boardHeight * 2, false, false, false, false, 0, 0, 
            //                                                        false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);        SW > 2017
            if (extrude != null)
            {
                extrude.Name = "BOARD_HOLES";
                boardSketch = extrude.Name;
                boolstatus = swModel.Extension.SelectByID2(boardSketch, "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
                // Get the sketch feature
                Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
                swExtrudeData.BothDirections = true;
                swExtrudeData.SetDepth(true, boardHeight * 0.55);
                swExtrudeData.SetDepth(false, boardHeight * 0.55);
                swFeature.ModifyDefinition(swExtrudeData, swModel, null);
                swExtrudeData.ReleaseSelectionAccess();
                swModel.ClearSelection2(true);
                // Set the color of the sketch
                double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
                swMaterialPropertyValues[0] = (double)((double)lblHolesColor.BackColor.R / 255);
                swMaterialPropertyValues[1] = (double)((double)lblHolesColor.BackColor.G / 255);
                swMaterialPropertyValues[2] = (double)((double)lblHolesColor.BackColor.B / 255);
                swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            }
            swModel.SelectionManager.EnableContourSelection = false;
            string boardPath = strpath;
            //swModel.ViewZoomtofit2();
            //swModel.ShowNamedView2("*Front", 1);
            status = swModel.SaveAs3(boardPath, 0, 2);
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
        }
        private void changeBoardColor(Color color, bool isBoard, SldWorks mSolidworksApplication)
        {
            string errors = "";
            int status = 0;
            int warnings = 0;
            bool boolstatus = false;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            if (swModel != null && boardPartID.Length > 1)
            {
                boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                swAssy.OpenCompFile();
                swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
                swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
                if (swModel != null)
                {
                    if (isBoard)
                        ChangeSketchColor(swModel, "BOARD_OUTLINE", ((double)color.R / 255), ((double)color.G / 255), ((double)color.B / 255), ref boolstatus, ref errors);
                    else
                        ChangeSketchColor(swModel, "BOARD_HOLES", ((double)color.R / 255), ((double)color.G / 255), ((double)color.B / 255), ref boolstatus, ref errors);
                    swModel.Visible = false;
                    status = swModel.SaveAs3(boardPartPath, 0, 2);
                    string ModelTitle = swModel.GetTitle();
                    swModel = null;
                    mSolidworksApplication.CloseDoc(ModelTitle);
                }
            }
        }
        private string createBoard(string emnfile, SldWorks mSolidworksApplication)
        {
            bool boolstatus;
            boolstatus = mSolidworksApplication.LoadFile2(emnfile, "r");
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            string ModelTitle = swModel.GetTitle();
            boolstatus = swModel.ForceRebuild3(false);
            swModel.ClearSelection2(true);
            string directoryPath = Path.GetDirectoryName(emnfile);
            string boardPath = Path.Combine(directoryPath, "board.SLDPRT");
            swModel.ViewZoomtofit2();
            swModel.ShowNamedView2("*Front", 1);
            long status = swModel.SaveAs3(boardPath, 0, 2);
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
            return boardPath;
        }
        private bool closeEnough(double a, double b)
        {
            return (0.00001 >= Math.Abs(a - b));
        }
        private int perpendicullarAxis(double[] angles)
        {
            //old XYZ are 0,1,2
            double old_x = 0;
            double old_y = 1;
            double old_z = 2;
            double new_x = 0;
            double new_y = 1;
            double new_z = 2;

            //X Rotation
            new_x = old_x;
            new_y = old_y * Math.Cos((angles[0] * Math.PI) / 180) - old_z * Math.Sin((angles[0] * Math.PI) / 180);
            new_z = old_y * Math.Sin((angles[0] * Math.PI) / 180) + old_z * Math.Cos((angles[0] * Math.PI) / 180);
            old_x = new_x;
            old_y = new_y;
            old_z = new_z;

            //Y rotation
            new_x = old_z * Math.Sin((angles[1] * Math.PI) / 180) + old_x * Math.Cos((angles[1] * Math.PI) / 180);
            new_y = old_y;
            new_z = old_z * Math.Cos((angles[1] * Math.PI) / 180) - old_x * Math.Sin((angles[1] * Math.PI) / 180);
            old_x = new_x;
            old_y = new_y;
            old_z = new_z;

            //Z rotation
            new_x = old_x * Math.Cos((angles[2] * Math.PI) / 180) - old_y * Math.Sin((angles[2] * Math.PI) / 180);
            new_y = old_x * Math.Sin((angles[2] * Math.PI) / 180) + old_y * Math.Cos((angles[2] * Math.PI) / 180);
            new_z = old_z;

            if (Math.Abs(2 - Math.Abs(new_x)) < 0.001)
                return 0;
            else if (Math.Abs(2 - Math.Abs(new_y)) < 0.001)
                return 1;
            else if (Math.Abs(2 - Math.Abs(new_z)) < 0.001)
                return 2;
            return 2;
        }
        private double[] getEulerAngle(string componentID, SldWorks mSolidworksApplication)
        {
            Component2 swComp;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            double[] nPts = new double[16];
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);

            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                for (int i = 0; i < 9; i++)
                    nPts[i] = swTransform.ArrayData[i];
                double E1, E2, E3;
                double sy = Math.Sqrt(swTransform.ArrayData[7] * swTransform.ArrayData[7] + swTransform.ArrayData[8] * swTransform.ArrayData[8]);
                if (sy > 1e-4)
                {
                    //x = atan2(R.at<double>(2, 1), R.at<double>(2, 2));
                    //y = atan2(-R.at<double>(2, 0), sy);
                    //z = atan2(R.at<double>(1, 0), R.at<double>(0, 0));

                    E1 = Math.Atan2(swTransform.ArrayData[7], swTransform.ArrayData[8]);
                    E2 = Math.Atan2(-swTransform.ArrayData[6], sy);
                    E3 = Math.Atan2(swTransform.ArrayData[3], swTransform.ArrayData[0]);
                }
                else
                {
                    E1 = Math.Atan2(-swTransform.ArrayData[5], swTransform.ArrayData[4]);
                    E2 = Math.Atan2(-swTransform.ArrayData[6], sy);
                    E3 = 0;
                }
                E1 = -E1 * 180 / PI;
                if (Math.Abs(E1) < 0.0001) E1 = 0;
                if (E1 < 0)
                    E1 = 360 + E1;
                if (E1 >= 359.99)
                    E1 = 0;
                E2 = -E2 * 180 / PI;
                if (Math.Abs(E2) < 0.0001) E2 = 0;
                if (E2 < 0)
                    E2 = 360 + E2;
                if (E2 >= 359.99)
                    E2 = 0;
                E3 = -E3 * 180 / PI;
                if (Math.Abs(E3) < 0.0001) E3 = 0;
                if (E3 < 0)
                    E3 = 360 + E3;
                if (E3 >= 359.99)
                    E3 = 0;
                return new double[] { E1, E2, E3 };
            }
            return null;
        }
        private double[] multiplySwTransform(double[] matrix1, double[] matrix2)
        {
            Emgu.CV.Matrix<double> rotationMatrix1 = new Emgu.CV.Matrix<double>(3, 3);
            rotationMatrix1[0, 0] = matrix1[0];
            rotationMatrix1[0, 1] = matrix1[3];
            rotationMatrix1[0, 2] = matrix1[6];

            rotationMatrix1[1, 0] = matrix1[1];
            rotationMatrix1[1, 1] = matrix1[4];
            rotationMatrix1[1, 2] = matrix1[7];

            rotationMatrix1[2, 0] = matrix1[2];
            rotationMatrix1[2, 1] = matrix1[5];
            rotationMatrix1[2, 2] = matrix1[8];

            Emgu.CV.Matrix<double> rotationMatrix2 = new Emgu.CV.Matrix<double>(3, 3);
            rotationMatrix2[0, 0] = matrix2[0];
            rotationMatrix2[0, 1] = matrix2[3];
            rotationMatrix2[0, 2] = matrix2[6];

            rotationMatrix2[1, 0] = matrix2[1];
            rotationMatrix2[1, 1] = matrix2[4];
            rotationMatrix2[1, 2] = matrix2[7];

            rotationMatrix2[2, 0] = matrix2[2];
            rotationMatrix2[2, 1] = matrix2[5];
            rotationMatrix2[2, 2] = matrix2[8];

            Emgu.CV.Matrix<double> matrix = new Emgu.CV.Matrix<double>(3, 3);
            matrix = rotationMatrix1 * rotationMatrix2;
            return new double[16] { matrix[0, 0], matrix[1, 0], matrix[2, 0], matrix[0, 1], matrix[1, 1], matrix[2, 1], matrix[0, 2], matrix[1, 2], matrix[2, 2], 0, 0, 0, 1, 0, 0, 0 };
        }

        private double[] matrixToSwTransform(double[,] matrix)
        {
            return new double[16] { matrix[0, 0], matrix[1, 0], matrix[2, 0], matrix[0, 1], matrix[1, 1], matrix[2, 1], matrix[0, 2], matrix[1, 2], matrix[2, 2], 0, 0, 0, 1, 0, 0, 0 };
        }
        /// <summary>
        /// https://math.stackexchange.com/questions/142821/matrix-for-rotation-around-a-vector
        /// </summary>
        /// <param name="unitVector"></param>
        /// <param name="angle"></param>
        /// <returns></returns>
        private double[] rodriguesRotation(double[] unitVector, double angle)
        {
            Emgu.CV.Matrix<double> W = new Emgu.CV.Matrix<double>(3, 3);
            W[0, 0] = 0;
            W[0, 1] = -unitVector[2];
            W[0, 2] = unitVector[1];

            W[1, 0] = unitVector[2];
            W[1, 1] = 0;
            W[1, 2] = -unitVector[0];

            W[2, 0] = -unitVector[1];
            W[2, 1] = unitVector[0];
            W[2, 2] = 0;

            Emgu.CV.Matrix<double> I = new Emgu.CV.Matrix<double>(3, 3);
            I.SetIdentity();
            Emgu.CV.Matrix<double> R = new Emgu.CV.Matrix<double>(3, 3);
            R = I + W.Mul(Math.Sin(angle * Math.PI / 180)) + (W * W).Mul(2 * Math.Sin(angle * Math.PI / (180 * 2)) * Math.Sin(angle * Math.PI / (180 * 2)));
            //double[,] ret = new double[3, 3];
            //R.CopyTo(ret);
            //ret =(double[,]) R.ManagedArray;
            double[] ret = matrixToSwTransform((double[,])R.ManagedArray);
            return ret;
        }

        private double[] getParallelUnitVector(string componentID, SldWorks mSolidworksApplication, bool XorY = false)
        {
            Component2 swComp;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            //double[] nPts = new double[16];
            //MathPoint swOriginPt;
            //MathVector swX_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                //for (int i = 0; i < 16; i++)
                //    nPts[i] = swTransform.ArrayData[i];
                Emgu.CV.Matrix<double> rotationMatrix1 = new Emgu.CV.Matrix<double>(3, 3);
                rotationMatrix1[0, 0] = swTransform.ArrayData[0];
                rotationMatrix1[0, 1] = swTransform.ArrayData[3];
                rotationMatrix1[0, 2] = swTransform.ArrayData[6];

                rotationMatrix1[1, 0] = swTransform.ArrayData[1];
                rotationMatrix1[1, 1] = swTransform.ArrayData[4];
                rotationMatrix1[1, 2] = swTransform.ArrayData[7];

                rotationMatrix1[2, 0] = swTransform.ArrayData[2];
                rotationMatrix1[2, 1] = swTransform.ArrayData[5];
                rotationMatrix1[2, 2] = swTransform.ArrayData[8];
                //double[,] rotationMatrix=new double[3,3];
                Emgu.CV.Matrix<double> unitVectorAlongZ1 = new Emgu.CV.Matrix<double>(1, 3);
                Emgu.CV.Matrix<double> newUnitVector1 = new Emgu.CV.Matrix<double>(1, 3);
                if (XorY)
                {
                    unitVectorAlongZ1[0, 0] = 1;
                    unitVectorAlongZ1[0, 1] = 0;
                    unitVectorAlongZ1[0, 2] = 0;

                    newUnitVector1[0, 0] = 1;
                    newUnitVector1[0, 1] = 0;
                    newUnitVector1[0, 2] = 0;
                }
                else
                {
                    unitVectorAlongZ1[0, 0] = 0;
                    unitVectorAlongZ1[0, 1] = 1;
                    unitVectorAlongZ1[0, 2] = 0;

                    newUnitVector1[0, 0] = 0;
                    newUnitVector1[0, 1] = 1;
                    newUnitVector1[0, 2] = 0;
                }
                newUnitVector1 = unitVectorAlongZ1 * rotationMatrix1;
                return new double[] { newUnitVector1[0, 0], newUnitVector1[0, 1], newUnitVector1[0, 2] };
            }
            return null;
        }

        private double[] getParallelUnitVector(double[] ArrayData, bool XorY = false)
        {
            //for (int i = 0; i < 16; i++)
            //    nPts[i] = swTransform.ArrayData[i];
            Emgu.CV.Matrix<double> rotationMatrix1 = new Emgu.CV.Matrix<double>(3, 3);
            rotationMatrix1[0, 0] = ArrayData[0];
            rotationMatrix1[0, 1] = ArrayData[3];
            rotationMatrix1[0, 2] = ArrayData[6];

            rotationMatrix1[1, 0] = ArrayData[1];
            rotationMatrix1[1, 1] = ArrayData[4];
            rotationMatrix1[1, 2] = ArrayData[7];

            rotationMatrix1[2, 0] = ArrayData[2];
            rotationMatrix1[2, 1] = ArrayData[5];
            rotationMatrix1[2, 2] = ArrayData[8];
            //double[,] rotationMatrix=new double[3,3];
            Emgu.CV.Matrix<double> unitVectorAlongZ1 = new Emgu.CV.Matrix<double>(1, 3);
            Emgu.CV.Matrix<double> newUnitVector1 = new Emgu.CV.Matrix<double>(1, 3);
            if (XorY)
            {
                unitVectorAlongZ1[0, 0] = 1;
                unitVectorAlongZ1[0, 1] = 0;
                unitVectorAlongZ1[0, 2] = 0;

                newUnitVector1[0, 0] = 1;
                newUnitVector1[0, 1] = 0;
                newUnitVector1[0, 2] = 0;
            }
            else
            {
                unitVectorAlongZ1[0, 0] = 0;
                unitVectorAlongZ1[0, 1] = 1;
                unitVectorAlongZ1[0, 2] = 0;

                newUnitVector1[0, 0] = 0;
                newUnitVector1[0, 1] = 1;
                newUnitVector1[0, 2] = 0;
            }

            newUnitVector1 = unitVectorAlongZ1 * rotationMatrix1;
            return new double[] { newUnitVector1[0, 0], newUnitVector1[0, 1], newUnitVector1[0, 2] };
        }

        private double[] getPerpendicullarUnitVector(double[] ArrayData)
        {
            //for (int i = 0; i < 16; i++)
            //    nPts[i] = swTransform.ArrayData[i];
            Emgu.CV.Matrix<double> rotationMatrix1 = new Emgu.CV.Matrix<double>(3, 3);
            rotationMatrix1[0, 0] = ArrayData[0];
            rotationMatrix1[0, 1] = ArrayData[3];
            rotationMatrix1[0, 2] = ArrayData[6];

            rotationMatrix1[1, 0] = ArrayData[1];
            rotationMatrix1[1, 1] = ArrayData[4];
            rotationMatrix1[1, 2] = ArrayData[7];

            rotationMatrix1[2, 0] = ArrayData[2];
            rotationMatrix1[2, 1] = ArrayData[5];
            rotationMatrix1[2, 2] = ArrayData[8];
            //double[,] rotationMatrix=new double[3,3];
            Emgu.CV.Matrix<double> unitVectorAlongZ1 = new Emgu.CV.Matrix<double>(1, 3);
            unitVectorAlongZ1[0, 0] = 0;
            unitVectorAlongZ1[0, 1] = 0;
            unitVectorAlongZ1[0, 2] = 1;
            Emgu.CV.Matrix<double> newUnitVector1 = new Emgu.CV.Matrix<double>(1, 3);
            newUnitVector1[0, 0] = 0;
            newUnitVector1[0, 1] = 0;
            newUnitVector1[0, 2] = 1;
            newUnitVector1 = unitVectorAlongZ1 * rotationMatrix1;
            return new double[] { newUnitVector1[0, 0], newUnitVector1[0, 1], newUnitVector1[0, 2] };
        }

        private double[] getPerpendicullarUnitVector(string componentID, SldWorks mSolidworksApplication)
        {
            Component2 swComp;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            //double[] nPts = new double[16];
            //MathPoint swOriginPt;
            //MathVector swX_Axis;
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                //for (int i = 0; i < 16; i++)
                //    nPts[i] = swTransform.ArrayData[i];
                Emgu.CV.Matrix<double> rotationMatrix1 = new Emgu.CV.Matrix<double>(3, 3);
                rotationMatrix1[0, 0] = swTransform.ArrayData[0];
                rotationMatrix1[0, 1] = swTransform.ArrayData[3];
                rotationMatrix1[0, 2] = swTransform.ArrayData[6];

                rotationMatrix1[1, 0] = swTransform.ArrayData[1];
                rotationMatrix1[1, 1] = swTransform.ArrayData[4];
                rotationMatrix1[1, 2] = swTransform.ArrayData[7];

                rotationMatrix1[2, 0] = swTransform.ArrayData[2];
                rotationMatrix1[2, 1] = swTransform.ArrayData[5];
                rotationMatrix1[2, 2] = swTransform.ArrayData[8];
                //double[,] rotationMatrix=new double[3,3];
                Emgu.CV.Matrix<double> unitVectorAlongZ1 = new Emgu.CV.Matrix<double>(1, 3);
                unitVectorAlongZ1[0, 0] = 0;
                unitVectorAlongZ1[0, 1] = 0;
                unitVectorAlongZ1[0, 2] = 1;
                Emgu.CV.Matrix<double> newUnitVector1 = new Emgu.CV.Matrix<double>(1, 3);
                newUnitVector1[0, 0] = 0;
                newUnitVector1[0, 1] = 0;
                newUnitVector1[0, 2] = 1;
                swTransform = swComp.Transform2;
                newUnitVector1 = unitVectorAlongZ1 * rotationMatrix1;
                return new double[] { newUnitVector1[0, 0], newUnitVector1[0, 1], newUnitVector1[0, 2] };
            }
            return null;
        }

        private double[] getPerpendicullarUnitVector(SldWorks mSolidworksApplication)
        {
            Component2 swComp;
            SelectionMgr swSelMgr = default(SelectionMgr);
            MathTransform swTransform = default(MathTransform);
            double[] nPts = new double[16];
            SldWorks swApp = mSolidworksApplication;
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComp != null)
            {
                swTransform = swComp.Transform2;
                for (int i = 0; i < 16; i++)
                    nPts[i] = swTransform.ArrayData[i];
                Emgu.CV.Matrix<double> rotationMatrix1 = new Emgu.CV.Matrix<double>(3, 3);
                rotationMatrix1[0, 0] = swTransform.ArrayData[0];
                rotationMatrix1[0, 1] = swTransform.ArrayData[3];
                rotationMatrix1[0, 2] = swTransform.ArrayData[6];

                rotationMatrix1[1, 0] = swTransform.ArrayData[1];
                rotationMatrix1[1, 1] = swTransform.ArrayData[4];
                rotationMatrix1[1, 2] = swTransform.ArrayData[7];

                rotationMatrix1[2, 0] = swTransform.ArrayData[2];
                rotationMatrix1[2, 1] = swTransform.ArrayData[5];
                rotationMatrix1[2, 2] = swTransform.ArrayData[8];
                //double[,] rotationMatrix=new double[3,3];
                Emgu.CV.Matrix<double> unitVectorAlongZ1 = new Emgu.CV.Matrix<double>(1, 3);
                unitVectorAlongZ1[0, 0] = 0;
                unitVectorAlongZ1[0, 1] = 0;
                unitVectorAlongZ1[0, 2] = 1;
                Emgu.CV.Matrix<double> newUnitVector1 = new Emgu.CV.Matrix<double>(1, 3);
                newUnitVector1[0, 0] = 0;
                newUnitVector1[0, 1] = 0;
                newUnitVector1[0, 2] = 1;
                swTransform = swComp.Transform2;
                newUnitVector1 = unitVectorAlongZ1 * rotationMatrix1;
                return new double[] { newUnitVector1[0, 0], newUnitVector1[0, 1], newUnitVector1[0, 2] };
            }
            return null;
        }
        private double getBoardThickness(SldWorks mSolidworksApplication)
        {
            int status = 0;
            int warnings = 0;
            bool boolstatus = false;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            swAssy.OpenCompFile();
            swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            string[] swModelConfigs = (string[])swModel.GetConfigurationNames();

            bool boolStatus;
            // Select the sketch
            boolStatus = swModel.Extension.SelectByID2(boardSketch, "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            // Get the sketch feature
            Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            // Clean up
            swModel.ClearSelection2(true);


            ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
            //swExtrudeData.SetDepth(true, 0.02);
            //swFeature.ModifyDefinition(swExtrudeData, swModel, null); 
            swModel.Visible = false;
            status = swModel.SaveAs3(boardPartPath, 0, 2);
            string ModelTitle = swModel.GetTitle();
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);

            return swExtrudeData.GetDepth(true);
        }
        private void setBoardThickness(double thickness,SldWorks mSolidworksApplication)
        {
            int status = 0;
            int warnings = 0;
            bool boolstatus = false;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            swAssy.OpenCompFile();
            swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            //string[] swModelConfigs = (string[])swModel.GetConfigurationNames();

            bool boolStatus;
            // Select the sketch
            boolStatus = swModel.Extension.SelectByID2("BOARD_OUTLINE", "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            // Get the sketch feature
            Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
            swExtrudeData.BothDirections = true;
            swExtrudeData.SetDepth(true, thickness/2);
            swExtrudeData.SetDepth(false, thickness / 2);
            swFeature.ModifyDefinition(swExtrudeData, swModel, null);
            swExtrudeData.ReleaseSelectionAccess();

            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            boolStatus = swModel.Extension.SelectByID2("BOARD_HOLES", "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            // Get the sketch feature
            swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            swExtrudeData = (ExtrudeFeatureData2)swFeature.GetDefinition();
            swExtrudeData.BothDirections = true;
            swExtrudeData.SetDepth(true, thickness*0.55);
            swExtrudeData.SetDepth(false, thickness * 0.55);
            swFeature.ModifyDefinition(swExtrudeData, swModel, null);
            swExtrudeData.ReleaseSelectionAccess();

            // Clean up
            swModel.ClearSelection2(true);
            swModel.Visible = false;
            status = swModel.SaveAs3(boardPartPath, 0, 2);
            string ModelTitle = swModel.GetTitle();
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
            double halfHeightDifference = (thickness-boardHeight)/2;
            boardHeight = thickness;

            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            Component2 swComponent = (Component2)swSelMgr.GetSelectedObject6(1, -1);
            if (swComponent != null)
            {
                MathTransform swTransform = default(MathTransform);
                swTransform = swComponent.Transform2;
                
                //double[] BoxArray = (double[])swComponent.GetBox(false, false);
                //double[] boardOriginShift = new double[3] { swTransform .ArrayData[9] - boardOrigin[0], swTransform.ArrayData[10] - boardOrigin[1] , swTransform.ArrayData[11] - boardOrigin[2]  };
                translateComponentDragOperator(boardPartID, new double[3] { 0, 0, -swTransform.ArrayData[11] }, false, TaskpaneIntegration.mSolidworksApplication,true);
                swTransform = swComponent.Transform2;
                boardOrigin = new double[3] { swTransform.ArrayData[9], swTransform.ArrayData[10], swTransform.ArrayData[11] };
                //boardOrigin[2] = (BoxArray[2] + BoxArray[5]) / 2;

                if (boardComponents.Count > 0)
                {
                    foreach (KeyValuePair<string, List<ComponentAttribute>> entry in boardComponents)
                    {
                        string componentPackage = entry.Key;
                        for (int i = 0; i < entry.Value.Count; i++)
                        {
                            string modelID = (string)entry.Value[i].modelID;
                            if (modelID.Length > 2)
                            {
                                if (boardComponents[componentPackage][i].side)
                                    translateComponentDragOperator(modelID, new double[3] { 0, 0, +halfHeightDifference }, false, TaskpaneIntegration.mSolidworksApplication);
                                else
                                    translateComponentDragOperator(modelID, new double[3] { 0, 0, -halfHeightDifference }, false, TaskpaneIntegration.mSolidworksApplication);
                            }
                        }
                    }
                }
            }
        }

        private void writeDecalImage(dynamic boardJSON, Bgra maskColor, Bgra traceColor, Bgra padColor, Bgra silkColor)
        {
            Progress P = new Progress( boardJSON,  maskColor,  traceColor,  padColor,  silkColor);
            P.Text = "JSON data rendering Progress";
            P.StartThread(1);
            P.ShowDialog();
            P = null;
            string directoryPath = Path.GetDirectoryName(boardPartPath);
            deleteFile(bottomDecal);

            bottomDecal = RandomString(10) + "bottom_decal.png";
            UtilityClass.BottomDecalImage.PyrUp().Save(@Path.Combine(directoryPath, bottomDecal));
            bottomDecal = @Path.Combine(directoryPath, bottomDecal);

            UtilityClass.TopDecalImage = UtilityClass.TopDecalImage.Flip(FlipType.Vertical);
            deleteFile(topDecal);
            topDecal = RandomString(10) + "top_decal.png";
            UtilityClass.TopDecalImage.PyrUp().Save(@Path.Combine(directoryPath, topDecal));
            topDecal = @Path.Combine(directoryPath, topDecal);


            UtilityClass.TopDecalImage.Dispose();
            UtilityClass.TopDecalImage = null;
            UtilityClass.BottomDecalImage.Dispose();
            UtilityClass.BottomDecalImage = null;
            GC.Collect();
        }
        private void deleteFile(string filePath)
        {
            if (filePath != null)
            {
                if (@filePath.Length > 1)
                {
                    if (File.Exists(@filePath))
                    {
                        File.Delete(@filePath);
                    }
                }
            }
        }
        private string getComponents(SldWorks mSolidworksApplication)
        {
            ModelDoc2 swModel;
            object[] vComps;
            int errors = 0;
            int i = 0;
            string output = "";
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            string AssemblyTitle = swModel.GetTitle();
            swModel = mSolidworksApplication.ActivateDoc3(AssemblyTitle, true, (int)swRebuildOnActivation_e.swUserDecision, errors);
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                AssemblyDoc swAssy = (AssemblyDoc)swModel;
                vComps = (object[])swAssy.GetComponents(true);
                for (i = 0; i < vComps.Length; i++)
                {
                    output += ((Component2)vComps[i]).Name2 + ",";
                }
            }
            return output;
        }
        private void saveDecalsFromBoard(SldWorks mSolidworksApplication,dynamic boardJSON)
        {
            int errors = 0;
            int status = 0;
            int warnings = 0;
            bool boolstatus;
            SelectionMgr swSelMgr = default(SelectionMgr);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            Face2 swFace = default(Face2);
            Decal swDecal = default(Decal);
            RenderMaterial swMaterial = default(RenderMaterial);
            string ModelTitle;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            swAssy.OpenCompFile();
            swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModelDocExt = swModel.Extension;

            swModel = mSolidworksApplication.ActivateDoc2(swModel.GetTitle(), true, errors);
            swModelDocExt = swModel.Extension;

            boolstatus = swModelDocExt.SelectByID2("", "FACE", boardOrigin[0], boardOrigin[1], boardOrigin[2] + (boardHeight / 2), false, 0, null, 0);
            string directoryPath = Path.GetDirectoryName(boardPartPath);
            //boolstatus = swModelDocExt.SelectByRay(boardFirstPoint[0], boardFirstPoint[1], (boardHeight / 2), 0, 0, -1, (boardHeight / 2), (int)swSelectType_e.swSelFACES, false, 0, (int)swSelectOption_e.swSelectOptionDefault); //SW > 2017
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swFace = (Face2)swSelMgr.GetSelectedObject6(1, -1);
            swModel.ClearSelection2(true);

            //Create the decal
            //swModelDocExt.GetDecal
            Object[] d= (Object[])swModelDocExt.GetDecals();
            if (d != null)
            {
                if (d.Length > 0)
                {
                    JArray Phy_Layers = boardJSON.Parts.Board_Part.Phy_Layers;
                    for (int i = 0; i < d.Length; i++)
                    {
                        swDecal = (Decal)d[i];
                        swMaterial = (RenderMaterial)swDecal;
                        int id = swDecal.DecalID;
                        if (id == topDecalID)
                        {
                            bool found = false;
                            int j = 0;
                            for(j=0;j< Phy_Layers.Count;j++)
                            {
                                dynamic temp = Phy_Layers[j];
                                string layerName = (string)temp.Layer_Name;
                                if(layerName.ToUpper()== "CONDUCTOR_TOP")
                                {
                                    found = true;
                                    break;
                                }
                            }
                            if(found)
                            {
                                boardJSON.Parts.Board_Part.Phy_Layers[j].DecalFile = topDecal;
                                boardJSON.Parts.Board_Part.Phy_Layers[j].DecalTransformation = new JArray { swMaterial.FixedAspectRatio,swMaterial.XPosition, swMaterial.YPosition , swMaterial.Width , swMaterial.Height , swMaterial.FitWidth, swMaterial.FitHeight };
                            }
                        }
                        else if (id == bottomDecalID)
                        {
                            bool found = false;
                            int j = 0;
                            for (j = 0; j < Phy_Layers.Count; j++)
                            {
                                dynamic temp = Phy_Layers[j];
                                string layerName = (string)temp.Layer_Name;
                                if (layerName.ToUpper() == "CONDUCTOR_BOTTOM")
                                {
                                    found = true;
                                    break;
                                }
                            }
                            if (found)
                            {
                                boardJSON.Parts.Board_Part.Phy_Layers[j].DecalFile = bottomDecal;
                                boardJSON.Parts.Board_Part.Phy_Layers[j].DecalTransformation = new JArray { swMaterial.FixedAspectRatio,swMaterial.XPosition, swMaterial.YPosition, swMaterial.Width, swMaterial.Height, swMaterial.FitWidth, swMaterial.FitHeight };
                            }
                        }
                    }
                }
            }
            ModelTitle = swModel.GetTitle();
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);

        }

        private void addSavedDecalsToBoard(SldWorks mSolidworksApplication)
        {
            int errors = 0;
            int status = 0;
            int warnings = 0;
            bool boolstatus;
            JArray Phy_Layers = boardJSON.Parts.Board_Part.Phy_Layers;
            JArray topDecalTransformation = null;
            JArray bottomDecalTransformation = null;
            int j = 0;
            bool found = false;
            for (j = 0; j < Phy_Layers.Count; j++)
            {
                dynamic temp = Phy_Layers[j];
                string layerName = (string)temp.Layer_Name;
                if (layerName.ToUpper() == "CONDUCTOR_TOP")
                {
                    found = true;
                    break;
                }
            }
            if (found)
            {
                topDecal=boardJSON.Parts.Board_Part.Phy_Layers[j].DecalFile ;
                if (topDecal != null)
                {
                    if (File.Exists(topDecal))
                        topDecalTransformation = boardJSON.Parts.Board_Part.Phy_Layers[j].DecalTransformation;
                    else
                        topDecal = "";
                }
            }

            j = 0;
            found = false;
            for (j = 0; j < Phy_Layers.Count; j++)
            {
                dynamic temp = Phy_Layers[j];
                string layerName = (string)temp.Layer_Name;
                if (layerName.ToUpper() == "CONDUCTOR_BOTTOM")
                {
                    found = true;
                    break;
                }
            }
            if (found)
            {
                bottomDecal = boardJSON.Parts.Board_Part.Phy_Layers[j].DecalFile;
                if (bottomDecal != null)
                {
                    if (File.Exists(bottomDecal))
                        bottomDecalTransformation = boardJSON.Parts.Board_Part.Phy_Layers[j].DecalTransformation;
                    else
                        bottomDecal = "";
                }
                
            }

            if (topDecalTransformation != null || bottomDecalTransformation != null)
            {
                if (File.Exists(topDecal) || File.Exists(bottomDecal))
                {
                    SelectionMgr swSelMgr = default(SelectionMgr);
                    ModelDocExtension swModelDocExt = default(ModelDocExtension);
                    Face2 swFace = default(Face2);
                    Decal swDecal = default(Decal);
                    RenderMaterial swMaterial = default(RenderMaterial);
                    int nDecalID = 0;
                    string ModelTitle;
                    ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
                    boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    AssemblyDoc swAssy = (AssemblyDoc)swModel;
                    swAssy.OpenCompFile();
                    swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
                    swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
                    swModelDocExt = swModel.Extension;

                    swModel = mSolidworksApplication.ActivateDoc2(swModel.GetTitle(), true, errors);
                    swModelDocExt = swModel.Extension;
                    if (File.Exists(topDecal))
                    {
                        boolstatus = swModelDocExt.SelectByID2("", "FACE", boardOrigin[0], boardOrigin[1], boardOrigin[2] + (boardHeight / 2), false, 0, null, 0);
                        if (boolstatus)
                        {
                            string directoryPath = Path.GetDirectoryName(boardPartPath);
                            //boolstatus = swModelDocExt.SelectByRay(boardFirstPoint[0], boardFirstPoint[1], (boardHeight / 2), 0, 0, -1, (boardHeight / 2), (int)swSelectType_e.swSelFACES, false, 0, (int)swSelectOption_e.swSelectOptionDefault); //SW > 2017
                            swSelMgr = (SelectionMgr)swModel.SelectionManager;
                            swFace = (Face2)swSelMgr.GetSelectedObject6(1, -1);
                            swModel.ClearSelection2(true);

                            //Create the decal
                            swModelDocExt.DeleteAllDecals();
                            swDecal = (Decal)swModelDocExt.CreateDecal();

                            swDecal.MaskType = 3;
                            swMaterial = (RenderMaterial)swDecal;
                            boolstatus = swMaterial.AddEntity(swFace);
                            swMaterial.TextureFilename = topDecal;
                            swMaterial.MappingType = 0;
                            swMaterial.FixedAspectRatio = false;
                            swMaterial.FixedAspectRatio = (bool)topDecalTransformation[0];
                            swMaterial.XPosition = (double)topDecalTransformation[1];
                            swMaterial.YPosition = (double)topDecalTransformation[2];
                            swMaterial.Height = (double)topDecalTransformation[3];
                            swMaterial.Width = (double)topDecalTransformation[4];
                            swMaterial.FitWidth = (bool)topDecalTransformation[5];
                            swMaterial.FitHeight = (bool)topDecalTransformation[6];

                            boolstatus = swModelDocExt.AddDecal(swDecal, out nDecalID);
                            topDecalID = nDecalID;
                        }
                    }
                    if (File.Exists(bottomDecal))
                    {
                        boolstatus = swModelDocExt.SelectByID2("", "FACE", boardOrigin[0], boardOrigin[1], boardOrigin[2] - (boardHeight / 2), false, 0, null, 0);
                        //boolstatus = swModelDocExt.SelectByRay(boardFirstPoint[0], boardFirstPoint[1], (boardHeight / 2), 0, 0, 1, (boardHeight / 2), (int)swSelectType_e.swSelFACES, false, 0, (int)swSelectOption_e.swSelectOptionDefault);     //sw>2017
                        if (boolstatus)
                        {
                            swSelMgr = (SelectionMgr)swModel.SelectionManager;
                            swFace = (Face2)swSelMgr.GetSelectedObject6(1, -1);
                            swModel.ClearSelection2(true);

                            //Create the decal
                            swDecal = (Decal)swModelDocExt.CreateDecal();
                            swDecal.MaskType = 3;
                            swMaterial = (RenderMaterial)swDecal;
                            boolstatus = swMaterial.AddEntity(swFace);
                            swMaterial.TextureFilename = bottomDecal;
                            swMaterial.MappingType = 0;
                            swMaterial.FixedAspectRatio = false;
                            swMaterial.FixedAspectRatio = (bool)bottomDecalTransformation[0];
                            swMaterial.XPosition = (double)bottomDecalTransformation[1];
                            swMaterial.YPosition = (double)bottomDecalTransformation[2];
                            swMaterial.Height = (double)bottomDecalTransformation[3];
                            swMaterial.Width = (double)bottomDecalTransformation[4];
                            swMaterial.FitWidth = (bool)bottomDecalTransformation[5];
                            swMaterial.FitHeight = (bool)bottomDecalTransformation[6];
                            boolstatus = swModelDocExt.AddDecal(swDecal, out nDecalID);
                            bottomDecalID = nDecalID;
                        }
                    }
                    swModel.Visible = false;
                    status = swModel.SaveAs3(boardPartPath, 0, 2);
                    ModelTitle = swModel.GetTitle();
                    swModel = null;
                    mSolidworksApplication.CloseDoc(ModelTitle);
                }
            }

        }

        private void addDecalsToBoard(SldWorks mSolidworksApplication)
        {
            int status = 0;
            int warnings = 0;
            bool boolstatus;
            SelectionMgr swSelMgr = default(SelectionMgr);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            Face2 swFace = default(Face2);
            Decal swDecal = default(Decal);
            RenderMaterial swMaterial = default(RenderMaterial);
            int nDecalID = 0;
            string ModelTitle;
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2(boardPartID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            swAssy.OpenCompFile();
            swModel = mSolidworksApplication.OpenDoc6(boardPartPath, 1, 0, "", status, warnings);
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModelDocExt = swModel.Extension;
            boolstatus = swModelDocExt.SelectByID2("BOARD_OUTLINE", "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            Feature swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
            object Box = null;
            swFeature.GetBox(ref Box);
            double[] BoxArray = (double[])Box;
            swModel.ClearSelection();

            boolstatus = swModelDocExt.SelectByID2("", "FACE", (BoxArray[0]+((BoxArray[3]- BoxArray[0])/2)), (BoxArray[1] + ((BoxArray[4] - BoxArray[1]) / 2)), BoxArray[5], false, 0, null, 0);
            if (boolstatus)
            {
                string directoryPath = Path.GetDirectoryName(boardPartPath);
                //boolstatus = swModelDocExt.SelectByRay(boardFirstPoint[0], boardFirstPoint[1], (boardHeight / 2), 0, 0, -1, (boardHeight / 2), (int)swSelectType_e.swSelFACES, false, 0, (int)swSelectOption_e.swSelectOptionDefault); //SW > 2017
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                swFace = (Face2)swSelMgr.GetSelectedObject6(1, -1);
                swModel.ClearSelection2(true);

                //Create the decal
                swModelDocExt.DeleteAllDecals();
                swDecal = (Decal)swModelDocExt.CreateDecal();

                swDecal.MaskType = 3;
                swMaterial = (RenderMaterial)swDecal;
                boolstatus = swMaterial.AddEntity(swFace);
                swMaterial.TextureFilename = topDecal;
                swMaterial.MappingType = 0;
                swMaterial.FixedAspectRatio = false;
                //swMaterial.Height = convertUnit(boardPhysicalHeight);
                swMaterial.FitHeight = true;
                swMaterial.Width = convertUnit(JSONLibrary.boardPhysicaWidth);
                //swMaterial.FitHeight = true;
                //swMaterial.FixedAspectRatio = true;
                swMaterial.FitWidth = true;
                //swMaterial.
                boolstatus = swModelDocExt.AddDecal(swDecal, out nDecalID);
                topDecalID = nDecalID;
            }
              boolstatus = swModelDocExt.SelectByID2("", "FACE", (BoxArray[0] + ((BoxArray[3] - BoxArray[0]) / 2)), (BoxArray[1] + ((BoxArray[4] - BoxArray[1]) / 2)), BoxArray[2], false, 0, null, 0);
            //boolstatus = swModelDocExt.SelectByRay(boardFirstPoint[0], boardFirstPoint[1], (boardHeight / 2), 0, 0, 1, (boardHeight / 2), (int)swSelectType_e.swSelFACES, false, 0, (int)swSelectOption_e.swSelectOptionDefault);     //sw>2017
            if (boolstatus)
            {
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                swFace = (Face2)swSelMgr.GetSelectedObject6(1, -1);
                swModel.ClearSelection2(true);

                //Create the decal
                swDecal = (Decal)swModelDocExt.CreateDecal();
                swDecal.MaskType = 3;
                swMaterial = (RenderMaterial)swDecal;
                boolstatus = swMaterial.AddEntity(swFace);
                swMaterial.TextureFilename = bottomDecal;
                swMaterial.MappingType = 0;
                //swMaterial.FixedAspectRatio = true;
                //swMaterial.FitHeight = true;
                swMaterial.FixedAspectRatio = false;
                //swMaterial.Height = convertUnit(boardPhysicalHeight);
                swMaterial.FitHeight = true;
                swMaterial.Width = convertUnit(JSONLibrary.boardPhysicaWidth);
                swMaterial.FitWidth = true;
                boolstatus = swModelDocExt.AddDecal(swDecal, out nDecalID);
                bottomDecalID = nDecalID;
            }
            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModelDocExt = swModel.Extension;
            string[] swModelConfigs = (string[])swModel.GetConfigurationNames();
            boolstatus = swModelDocExt.SelectByID2("BOARD_OUTLINE", "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            if (boolstatus)
            {
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                swModel.ClearSelection2(true);
                // Set the color of the sketch
                double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
                swMaterialPropertyValues[0] = (double)((double)lblBoardColor.BackColor.R/255);
                swMaterialPropertyValues[1] = (double)((double)lblBoardColor.BackColor.G/255);
                swMaterialPropertyValues[2] = (double)((double)lblBoardColor.BackColor.B/255);
                swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swAllConfiguration, swModelConfigs);
            }

            swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModelDocExt = swModel.Extension;
            swModelConfigs = (string[])swModel.GetConfigurationNames();
            boolstatus = swModelDocExt.SelectByID2("BOARD_HOLES", "BODYFEATURE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            if (boolstatus)
            {
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                swFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                swModel.ClearSelection2(true);
                // Set the color of the sketch
                double[] swMaterialPropertyValues = (double[])swModel.MaterialPropertyValues;
                swMaterialPropertyValues[0] = (double)((double)lblHolesColor.BackColor.R / 255);
                swMaterialPropertyValues[1] = (double)((double)lblHolesColor.BackColor.G / 255);
                swMaterialPropertyValues[2] = (double)((double)lblHolesColor.BackColor.B / 255);
                swFeature.SetMaterialPropertyValues2(swMaterialPropertyValues, (int)swInConfigurationOpts_e.swAllConfiguration, swModelConfigs);
            }
            swModel.Visible = false;
            status = swModel.SaveAs3(boardPartPath, 0, 2);
            ModelTitle = swModel.GetTitle();
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
            //changeBoardColor(lblBoardColor.BackColor, true, TaskpaneIntegration.mSolidworksApplication);
            //changeBoardColor(lblHolesColor.BackColor, false, TaskpaneIntegration.mSolidworksApplication);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog Openfile = new OpenFileDialog
            {
                Filter = "All supported files (*.json,*.idf,*.bdf,*.emn)|*.json;*.idf;*.bdf;*.emn|" +
                         "Board JSON  files (*.json) |*.json|" +
                         "IDF3 Files (*.idf,*.bdf,*.emn)|*.idf;*.bdf;*.emn" ,
                Title = "Open Board ECAD file"
            };
            if (Openfile.ShowDialog() == DialogResult.OK)
            {
                string ext = Path.GetExtension(@Openfile.FileName).ToUpper();
                ComponentTransformations.Clear();
                boardComponents.Clear();
                ComponentModels.Clear();
                CurrentNotesTop.Clear();
                CurrentNotesBottom.Clear();
                comboComponentList.Items.Clear();
                checkBoxTopOnly.Checked = false;
                checkBoxBottomOnly.Checked = false;
                isIDFFile = false;
                bool failed = false;
                if (ext == ".JSON")
                {
                    messageForm.Text = "Please Wait";
                    messageForm.labelMessage.Text = "Processing Board Outline. . . ";
                    messageForm.Show(this);
                    boardPartPath = createBoardFromJSON(@Openfile.FileName, TaskpaneIntegration.mSolidworksApplication);
                    boardPartID = insertPart(boardPartPath, 0, 0, true, true, 0, 2, "", TaskpaneIntegration.mSolidworksApplication, true);
                    messageForm.labelMessage.Text = "Processing Traces and Silkscreen. . . ";
                    addSavedDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
                    messageForm.labelMessage.Text = "Processing Holes . . . ";
                    makeHoles(boardPartPath, TaskpaneIntegration.mSolidworksApplication);
                    //addDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
                    messageForm.labelMessage.Text = "Processing Components . . . ";
                    readComponents(boardJSON);
                    propagateComponentList();
                    comboComponentList.SelectedIndex = 0;
                    messageForm.labelMessage.Text = "Inserting Components . . . ";
                    insertModels();
                    messageForm.Hide();
                }
                else if (ext == ".IDF" || ext == ".BDF" || ext == ".EMN")
                {
                    messageForm.Text = "Please Wait";
                    messageForm.labelMessage.Text = "Processing Board Outline. . . ";
                    messageForm.Show(this);
                    boardPartPath = createBoardFromIDF(@Openfile.FileName, TaskpaneIntegration.mSolidworksApplication);
                    boardPartID = insertPart(boardPartPath, 0, 0, true, true, 0, 2, "", TaskpaneIntegration.mSolidworksApplication, true);
                    messageForm.labelMessage.Text = "Processing Holes . . . ";
                    makeHolesFromIDF(@Openfile.FileName, boardPartPath, TaskpaneIntegration.mSolidworksApplication);
                    messageForm.labelMessage.Text = "Processing Components . . . ";
                    readComponentFromIDF(@Openfile.FileName, TaskpaneIntegration.mSolidworksApplication);
                    propagateComponentList();
                    comboComponentList.SelectedIndex = 0;
                    insertModels();
                    isIDFFile = true;
                    messageForm.Hide();
                }
                else
                {
                    MessageBox.Show("File format not supported.");
                    failed = true;
                }
                if (!failed)
                {
                    int errors = 0;
                    ModelDoc2 swModel = (ModelDoc2)TaskpaneIntegration.mSolidworksApplication.ActiveDoc;
                    string AssemblyTitle = swModel.GetTitle();
                    swModel = TaskpaneIntegration.mSolidworksApplication.ActivateDoc3(AssemblyTitle, true, (int)swRebuildOnActivation_e.swUserDecision, errors);
                    globalSwAssy = (AssemblyDoc)swModel;
                    globalSwAssy.UserSelectionPostNotify += SwAssy_UserSelectionPostNotify;
                    globalSwAssy.DeleteItemNotify += SwAssy_DeleteItemNotify;
                    globalSwAssy.RenameItemNotify += SwAssy_RenameItemNotify;
                    globalSwAssy.DestroyNotify2 += SwAssy_DestroyNotify2;
                    FileOpened = true;
                }
            }
            //deleteFile(topDecal);
            //deleteFile(bottomDecal);

        }
        private int SwAssy_DestroyNotify2(int DestroyType)
        {
            getLibraryTransform();
            saveLibraryasProgramData();
            ComponentTransformations.Clear();
            boardComponents.Clear();
            ComponentModels.Clear();
            CurrentNotesTop.Clear();
            CurrentNotesBottom.Clear();
            comboComponentList.Items.Clear();
            checkBoxTopOnly.Checked = false;
            checkBoxBottomOnly.Checked = false;
            FileOpened = false;
            labelAssignedFile.Text = "None";
            labelInstances.Text = "None";

            globalSwAssy.UserSelectionPostNotify -= SwAssy_UserSelectionPostNotify;
            globalSwAssy.DeleteItemNotify -= SwAssy_DeleteItemNotify;
            globalSwAssy.RenameItemNotify -= SwAssy_RenameItemNotify;
            globalSwAssy.DestroyNotify2 -= SwAssy_DestroyNotify2;
            
            return 0;
        }

        private void buttonApplyStyle_Click(object sender, EventArgs e)
        {

            if (FileOpened)
            {

                if (gerberFile.Length > 2 && File.Exists(gerberFile))
                {
                    Progress P = new Progress(@gerberFile, lblMaskColor.BackColor, lblSilkColor.BackColor, lblPadColor.BackColor, lblTraceColor.BackColor);
                    P.Text = "Gerber Rendering Progress";
                    P.StartThread(0);
                    P.ShowDialog();
                    P = null;
                    string directoryPath = Path.GetDirectoryName(boardPartPath);
                    Mat imgPng = new Mat();
                    CvInvoke.Imdecode(UtilityClass.TopDecal, ImreadModes.Unchanged, imgPng);
                    Image<Bgra, byte> TopDecalImage = imgPng.ToImage<Bgra, byte>();
                    //TopDecalImage = TopDecalImage.Flip(FlipType.Vertical);
                    deleteFile(topDecal);
                    topDecal = RandomString(10) + "top_decal.png";
                    TopDecalImage.PyrUp().Save(@Path.Combine(directoryPath, topDecal));
                    topDecal = @Path.Combine(directoryPath, topDecal);
                    UtilityClass.TopDecal = null;
                    imgPng.Dispose();
                    TopDecalImage.Dispose();

                    imgPng = new Mat();
                    CvInvoke.Imdecode(UtilityClass.BottomDecal, ImreadModes.Unchanged, imgPng);
                    Image<Bgra, byte> BottomDecalImage = imgPng.ToImage<Bgra, byte>();
                    BottomDecalImage = BottomDecalImage.Flip(FlipType.Vertical).Flip(FlipType.Horizontal);
                    deleteFile(bottomDecal);
                    bottomDecal = RandomString(10) + "bottom_decal.png";
                    BottomDecalImage.PyrUp().Save(@Path.Combine(directoryPath, bottomDecal));
                    bottomDecal = @Path.Combine(directoryPath, bottomDecal);
                    UtilityClass.BottomDecal = null;
                    imgPng.Dispose();
                    BottomDecalImage.Dispose();

                    addDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    if (!isIDFFile)
                    {
                        List<Bgra> bgras = new List<Bgra> { };
                        Color color = lblMaskColor.BackColor;
                        var argbarray = BitConverter.GetBytes(color.ToArgb())
                                        .Reverse()
                                        .ToArray();
                        Bgra temp = new Bgra(argbarray[3], argbarray[2], argbarray[1], argbarray[0]);
                        bgras.Add(temp);
                        color = lblTraceColor.BackColor;
                        argbarray = BitConverter.GetBytes(color.ToArgb())
                                        .Reverse()
                                        .ToArray();
                        temp = new Bgra(argbarray[3], argbarray[2], argbarray[1], argbarray[0]);
                        bgras.Add(temp);
                        color = lblPadColor.BackColor;
                        argbarray = BitConverter.GetBytes(color.ToArgb())
                                        .Reverse()
                                        .ToArray();
                        temp = new Bgra(argbarray[3], argbarray[2], argbarray[1], argbarray[0]);
                        bgras.Add(temp);
                        color = lblSilkColor.BackColor;
                        argbarray = BitConverter.GetBytes(color.ToArgb())
                                        .Reverse()
                                        .ToArray();
                        temp = new Bgra(argbarray[3], argbarray[2], argbarray[1], argbarray[0]);
                        bgras.Add(temp);
                        writeDecalImage(boardJSON, bgras[0], bgras[1], bgras[2], bgras[3]);

                        addDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
                    }
                    else
                    {
                        MessageBox.Show("IDF files contains no style information. You can only change the board and holes color.");
                    }
                }


                //else
                //{
                //    if (gerberFile.Length > 2)
                //    {
                //        if (File.Exists(gerberFile))
                //        {
                //            Progress P = new Progress(@gerberFile, lblMaskColor.BackColor, lblSilkColor.BackColor, lblPadColor.BackColor, lblTraceColor.BackColor);
                //            P.Text = "Gerber Rendering Progress";
                //            P.StartThread(0);
                //            P.ShowDialog();
                //            P = null;
                //            string directoryPath = Path.GetDirectoryName(boardPartPath);
                //            Mat imgPng = new Mat();
                //            CvInvoke.Imdecode(UtilityClass.TopDecal, ImreadModes.Unchanged, imgPng);
                //            Image<Bgra, byte> TopDecalImage = imgPng.ToImage<Bgra, byte>();
                //            //TopDecalImage = TopDecalImage.Flip(FlipType.Vertical);
                //            deleteFile(topDecal);
                //            topDecal = RandomString(10) + "top_decal.png";
                //            TopDecalImage.PyrUp().Save(@Path.Combine(directoryPath, topDecal));
                //            topDecal = @Path.Combine(directoryPath, topDecal);
                //            UtilityClass.TopDecal = null;
                //            imgPng.Dispose();
                //            TopDecalImage.Dispose();

                //            imgPng = new Mat();
                //            CvInvoke.Imdecode(UtilityClass.BottomDecal, ImreadModes.Unchanged, imgPng);
                //            Image<Bgra, byte> BottomDecalImage = imgPng.ToImage<Bgra, byte>();
                //            BottomDecalImage = BottomDecalImage.Flip(FlipType.Vertical).Flip(FlipType.Horizontal);
                //            deleteFile(bottomDecal);
                //            bottomDecal = RandomString(10) + "bottom_decal.png";
                //            BottomDecalImage.PyrUp().Save(@Path.Combine(directoryPath, bottomDecal));
                //            bottomDecal = @Path.Combine(directoryPath, bottomDecal);
                //            UtilityClass.BottomDecal = null;
                //            imgPng.Dispose();
                //            BottomDecalImage.Dispose();

                //            addDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
                //        }
                //        else
                //            MessageBox.Show("IDF files contains no style information. You can only change the board and holes color.");
                //    }
                //    else
                //        MessageBox.Show("IDF files contains no style information. You can only change the board and holes color.");
                //}
            }
            //deleteFile(topDecal);
            // deleteFile(bottomDecal);
        }
        private void ComponentListSelectEvent()
        {
            if (FileOpened)
            {
                string componentPackage = (string)comboComponentList.SelectedItem;
                comboComponentList.Refresh();
                if (boardComponents.ContainsKey(componentPackage))
                {
                    List<ComponentAttribute> c = boardComponents[componentPackage];
                    labelInstances.Text = c.Count.ToString();
                    if (ComponentModels.ContainsKey(componentPackage))
                    {
                        if (ComponentModels[componentPackage].Length > 2 && File.Exists(ComponentModels[componentPackage]))
                        {
                            labelAssignedFile.Text = Path.GetFileName(ComponentModels[componentPackage]);
                            buttonClearModel.Enabled = true;
                        }
                        else
                        {
                            labelAssignedFile.Text = "None";
                            buttonClearModel.Enabled = false;
                        }
                    }
                    else
                    {
                        labelAssignedFile.Text = "None";
                        buttonClearModel.Enabled = false;
                    }
                    buttonBrowseModel.Enabled = true;
                    InsertNoteForPackage(componentPackage, TaskpaneIntegration.mSolidworksApplication);
                }

                if (checkLoadSelected.Checked)
                {
                    bool found = false;
                    if (boardComponents.Count > 0)
                    {
                        string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
                        string libraryJSONPath = @Path.Combine(directory, pluginName, libraryFileName);
                        ConcurrentDictionary<string, ComponentLibrary> libraryJSONList = JsonConvert.DeserializeObject<ConcurrentDictionary<string, ComponentLibrary>>(File.ReadAllText(@libraryJSONPath));
                        int count = libraryJSONList.Count;
                        Parallel.ForEach(boardComponents, (c, loopState) =>
                        {
                            if (c.Key == componentPackage)
                            {
                                if (libraryJSONList.ContainsKey(c.Key))
                                {
                                    found = true;
                                    loopState.Stop();
                                }
                            }
                        });
                        if(found)
                            buttonLoadLibrary.Enabled = true;
                        else
                            buttonLoadLibrary.Enabled = false;

                    }
                }
                else
                {
                    buttonLoadLibrary.Enabled = true;
                }
            }
        }
        private void comboComponentList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComponentListSelectEvent();
        }
        private void removeComponent(string componentID, SldWorks mSolidworksApplication)
        {
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swModel.EditDelete();
        }

        private void selectComponents(string componentID,bool append, SldWorks mSolidworksApplication)
        {
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModel.Extension.SelectByID2(componentID, "COMPONENT", 0, 0, 0, append, 0, null, 0);
        }

        private void deleteSelectComponents(SldWorks mSolidworksApplication)
        {
            ModelDoc2 swModel = (ModelDoc2)mSolidworksApplication.ActiveDoc;
            swModel.EditDelete();
        }
        private void buttonClearModel_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                string componentPackage = (string)comboComponentList.SelectedItem;
                if (boardComponents.ContainsKey(componentPackage))
                {
                    ComponentModels[componentPackage] = "";
                    bool first = true;
                    for (int i = 0; i < boardComponents[componentPackage].Count; i++)
                    {
                        if (boardComponents[componentPackage][i].modelID.Length > 2)
                        {
                            boardComponents[componentPackage][i].modelPath = "";
                            //removeComponent(boardComponents[componentPackage][i].modelID, TaskpaneIntegration.mSolidworksApplication);
                            if (first)
                            {
                                selectComponents(boardComponents[componentPackage][i].modelID, false, TaskpaneIntegration.mSolidworksApplication);
                                first = false;
                            }
                            else
                                selectComponents(boardComponents[componentPackage][i].modelID, true, TaskpaneIntegration.mSolidworksApplication);
                            boardComponents[componentPackage][i].modelID = "";
                        }
                    }
                    if(!first)
                        deleteSelectComponents(TaskpaneIntegration.mSolidworksApplication);
                    labelInstances.Text = "None";
                    labelAssignedFile.Text = "None";
                    comboComponentList.Refresh();
                }
                ComponentListSelectEvent();
            }
        }
        private void buttonBrowseModel_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                string componentPackage = (string)comboComponentList.SelectedItem;
                if (boardComponents.ContainsKey(componentPackage))
                {
                    OpenFileDialog Openfile = new OpenFileDialog
                    {
                        Filter = "Solidworks Part Files(*.SLDPRT) | *.SLDPRT"
                    };
                    if (Openfile.ShowDialog() == DialogResult.OK)
                    {
                        if (boardComponents.ContainsKey(componentPackage))
                        {
                            ComponentModels[componentPackage] = @Openfile.FileName;
                            for (int i = 0; i < boardComponents[componentPackage].Count; i++)
                            {
                                if (boardComponents[componentPackage][i].modelID.Length > 2)
                                {
                                    boardComponents[componentPackage][i].modelPath = "";
                                    removeComponent(boardComponents[componentPackage][i].modelID, TaskpaneIntegration.mSolidworksApplication);
                                    boardComponents[componentPackage][i].modelID = "";
                                }
                            }
                            labelInstances.Text = "None";
                            labelAssignedFile.Text = "None";
                            comboComponentList.Refresh();
                        }
                        if (!ComponentModels.ContainsKey(componentPackage))
                        {
                            ComponentModels.TryAdd(componentPackage, @Openfile.FileName);
                        }
                        labelAssignedFile.Text = Path.GetFileName(@Openfile.FileName);
                        buttonClearModel.Enabled = true;
                        List<ComponentAttribute> c = boardComponents[componentPackage];
                        for (int i = 0; i < c.Count; i++)
                        {
                            PointF pointFs = c[i].location;
                            bool side = c[i].side;
                            float rotation = c[i].rotation;
                            if (boardComponents[componentPackage][i].modelID != null)
                            {
                                if (boardComponents[componentPackage][i].modelID.Length > 2)
                                {
                                    removeComponent(boardComponents[componentPackage][i].modelID, TaskpaneIntegration.mSolidworksApplication);
                                    boardComponents[componentPackage][i].modelID = insertPart(@Openfile.FileName, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, false, rotation, 2, c[i].componentID, TaskpaneIntegration.mSolidworksApplication);
                                }
                                else
                                {
                                    boardComponents[componentPackage][i].modelID = insertPart(@Openfile.FileName, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, false, rotation, 2, c[i].componentID, TaskpaneIntegration.mSolidworksApplication);
                                }
                            }
                            else
                            {
                                boardComponents[componentPackage][i].modelID = insertPart(@Openfile.FileName, convertUnit(pointFs.X), convertUnit(pointFs.Y), side, false, rotation, 2, c[i].componentID, TaskpaneIntegration.mSolidworksApplication);
                            }
                        }
                        ComponentListSelectEvent();
                        comboComponentList.Refresh();
                    }
                }
            }
        }

        private void colorPresets_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<Bgra> bgras = applyPCBStyle(colorPresets.SelectedIndex);
            lblMaskColor.BackColor = Color.FromArgb((int)bgras[0].Alpha, (int)bgras[0].Red, (int)bgras[0].Green, (int)bgras[0].Blue);
            lblTraceColor.BackColor = Color.FromArgb((int)bgras[1].Alpha, (int)bgras[1].Red, (int)bgras[1].Green, (int)bgras[1].Blue);
            lblPadColor.BackColor = Color.FromArgb((int)bgras[2].Alpha, (int)bgras[2].Red, (int)bgras[2].Green, (int)bgras[2].Blue);
            lblSilkColor.BackColor = Color.FromArgb((int)bgras[3].Alpha, (int)bgras[3].Red, (int)bgras[3].Green, (int)bgras[3].Blue);
            lblBoardColor.BackColor = Color.FromArgb((int)bgras[4].Alpha, (int)bgras[4].Red, (int)bgras[4].Green, (int)bgras[4].Blue);
            lblHolesColor.BackColor = Color.FromArgb((int)bgras[5].Alpha, (int)bgras[5].Red, (int)bgras[5].Green, (int)bgras[5].Blue);
        }

        private void TaskpaneHostUI_Load(object sender, EventArgs e)
        {
            populateStyleComboBox();
            colorPresets.SelectedIndex = 1;
            comboAlongAxis.SelectedIndex = 2;
            numericUpDownIncrement.DecimalPlaces = 3;

            //lblBoardColor.BackColor = Color.FromArgb(255, 79, 63, 0);
            //lblHolesColor.BackColor = Color.FromArgb(255, 128, 128, 128);
            //cb.Items.Add(new CComboboxItem("Item Number 1", Color.Green, Color.Yellow));
            //cb.Items.Add(new CComboboxItem("Item Number 2", Color.Blue, Color.Red));
            //cb.Items.Add(new CComboboxItem("Item Number 3", Color.Red, Color.Plum));
        }


        //MessageBox.Show("file closed");
        private void lblMaskColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog dlg = new ColorDialog
                {
                    Color = lblMaskColor.BackColor,
                    FullOpen = true
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblMaskColor.BackColor = dlg.Color;
                }
            }
            catch (Exception)
            { }
        }

        private void lblTraceColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog dlg = new ColorDialog
                {
                    Color = lblTraceColor.BackColor,
                    FullOpen = true
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblTraceColor.BackColor = dlg.Color;
                }
            }
            catch (Exception)
            { }
        }

        private void lblPadColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog dlg = new ColorDialog
                {
                    Color = lblPadColor.BackColor,
                    FullOpen = true
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblPadColor.BackColor = dlg.Color;
                }
            }
            catch (Exception)
            { }
        }

        private void lblSilkColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog dlg = new ColorDialog
                {
                    Color = lblSilkColor.BackColor,
                    FullOpen = true
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblSilkColor.BackColor = dlg.Color;
                }
            }
            catch (Exception)
            { }
        }
        private void saveIDFToJSON()
        {
            List<List<IDFBoardOutline>> IDFboardData = new List<List<IDFBoardOutline>> { new List<IDFBoardOutline> { } };
            double boardThickness = 0;
            const Int32 BufferSize = 2048;
            using (var fileStream = File.OpenRead(openedFile))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;
                List<IDFBoardOutline> currentOutline = new List<IDFBoardOutline> { };
                bool startProcessing = false;
                int previndex = -1;
                int curIndex = -1;
                while ((line = streamReader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (!startProcessing && line.IndexOf(".BOARD_OUTLINE") != -1)
                    {
                        startProcessing = true;
                        line = streamReader.ReadLine();
                        boardThickness = convertUnit(Convert.ToDouble(line));
                        IDFboardData.Clear();
                        continue;
                    }
                    if (startProcessing && line.IndexOf(".END_BOARD_OUTLINE") != -1)
                    {
                        IDFboardData.Add(new List<IDFBoardOutline>(currentOutline));
                        break;
                    }
                    if (startProcessing)
                    {
                        if (line.Length > 2)
                        {
                            //List<string> lineElements = line.Split(' ').ToList();
                            List<string> lineElements = splitString(line);
                            List<double> lineElementsDouble = new List<double> { };
                            lineElementsDouble.Clear();
                            for (int i = 0; i < lineElements.Count; i++)
                            {
                                lineElements[i] = lineElements[i].Trim();
                                if (lineElements[i].Length > 0)
                                {
                                    lineElementsDouble.Add(Convert.ToDouble(lineElements[i]));
                                }
                            }
                            curIndex = (int)lineElementsDouble[0];
                            if (curIndex != previndex)
                            {
                                if (currentOutline.Count > 0)
                                {
                                    IDFboardData.Add(new List<IDFBoardOutline>(currentOutline));
                                }
                                currentOutline.Clear();
                            }
                            IDFBoardOutline outline = new IDFBoardOutline();
                            outline.location.X = (float)Convert.ToDouble(lineElementsDouble[1]);
                            outline.location.Y = (float)Convert.ToDouble(lineElementsDouble[2]);
                            outline.rotation = (float)Convert.ToDouble(lineElementsDouble[3]);
                            currentOutline.Add(outline);
                            previndex = curIndex;
                        }
                    }
                }
            }
            dynamic template = readTemplate();
            JArray vertices1 = new JArray();
            List<IDFBoardOutline> vertices = IDFboardData[0];
            for (int i = 0; i < vertices.Count; i++)
                vertices1.Add(new JArray(new double[] { vertices[i].location.X, vertices[i].location.Y, vertices[i].rotation }));
            template.Parts.Board_Part.Shape.Extrusion.Outline.Polycurve_Area.Vertices = vertices1;
            if(IDFboardData.Count>1)
            {
                vertices1 = new JArray();
                vertices1 = (JArray)template.Parts.Board_Part.Features;
                for(int i=1;i< IDFboardData.Count; i++)
                {
                    vertices1.Add(addFeatureCutout(IDFboardData[i]));
                }
                
            }
            dynamic partEnd = JsonConvert.DeserializeObject<dynamic>("{\"Part_end\":\"True\"}");
            vertices1.Add(partEnd);
            template.Parts.Board_Part.Features = vertices1;
            using (var fileStream = File.OpenRead(openedFile))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;
                vertices1 = new JArray();
                vertices1 = (JArray)template.Parts.Board_Part.Features;
                List<IDFBoardOutline> currentOutline = new List<IDFBoardOutline> { };
                bool startProcessing = false;
                while ((line = streamReader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (!startProcessing && line.IndexOf(".DRILLED_HOLES") != -1)
                    {
                        startProcessing = true;
                        continue;
                    }
                    if (startProcessing && line.IndexOf(".END_DRILLED_HOLES") != -1)
                    {
                        template.Parts.Board_Part.Features = vertices1;
                        break;
                    }
                    if (startProcessing)
                    {
                        if (line.Length > 2)
                        {
                            //List<string> lineElements = line.Split(' ').ToList();
                            List<string> lineElements = splitString(line);
                            List<double> lineElementsDouble = new List<double> { };
                            lineElementsDouble.Clear();
                            int k = 0;
                            for (int i = 0; i < lineElements.Count; i++)
                            {
                                lineElements[i] = lineElements[i].Trim();
                                if (lineElements[i].Length > 0)
                                {
                                    lineElementsDouble.Add(Convert.ToDouble(lineElements[i]));
                                    k++;
                                    if (k >= 3)
                                        break;
                                }
                            }
                            vertices1.Add(addFeatureHoles(lineElementsDouble[1],lineElementsDouble[2], (lineElementsDouble[0]/2)));
                        }
                    }
                }
            }
            vertices1 = new JArray();
            vertices1 = (JArray)template.Assemblies.Board_Assembly.Comp_Insts;
            JArray partsArray = new JArray();
            partsArray = (JArray)template.Parts.Electrical_Part;
            List<string> keys = new List<string>(boardComponents.Keys);
            foreach (string key in keys)
            {
                if (ComponentModels.ContainsKey(key))
                {
                    string modelPath = ComponentModels[key];
                    List<ComponentAttribute> c = boardComponents[key];
                    partsArray.Add(addPart(boardComponents[key][0], key));                  //parts array
                    for (int i = 0; i < c.Count; i++)
                        vertices1.Add(addPartInstance(boardComponents[key][i], key));           //parts array instance
                }
            }
            template.Assemblies.Board_Assembly.Comp_Insts = vertices1;
            template.Parts.Electrical_Part = partsArray;
            template.Parts.Board_Part.Shape.Extrusion.Top_Height = (double)numericUpDownBoardHeight.Value;
            saveDecalsFromBoard(TaskpaneIntegration.mSolidworksApplication, template);
            saveJSONtoFile(template);
            //SaveFileDialog Savefile = new SaveFileDialog
            //{
            //    Filter = "Board JSON files (*.json) | *.json",
            //    OverwritePrompt = true,
            //    Title = "Save Board JSON file As. . .",
            //    SupportMultiDottedExtensions = false
            //};
            //if (Savefile.ShowDialog() == DialogResult.OK)
            //{
            //    using (StreamWriter file = File.CreateText(@Savefile.FileName))
            //    {
            //        //string json = JsonConvert.SerializeObject(boardJSON);
            //        JsonSerializer serializer = new JsonSerializer();
            //        serializer.Serialize(file, template);
            //    }
            //}
        }
        private string generateId()
        {
            idstart++;
            return idstart.ToString();
        }
        private dynamic addFeatureCutout(List<IDFBoardOutline> vertices)
        {
            
            string cutoutTemplate = "{\"Feature_Type\":\"Cutout\",\"Entity_ID\":\"#"+ generateId() + "\",\"Outline\":{\"Polycurve_Area\":{\"Entity_ID\":\"#"+ generateId()+"\",\"Line_Font\":\"Solid\",\"Line_Color\":[0.0,0.0,0.0],\"Fill_Color\":[0.0,0.0,0.0],\"Vertices\":[]}}}";
            dynamic template = JsonConvert.DeserializeObject<dynamic>(cutoutTemplate);
            JArray vertices1 = new JArray();
            for (int i = 0; i < vertices.Count; i++)
                vertices1.Add(new JArray(new double[] { vertices[i].location.X, vertices[i].location.Y, vertices[i].rotation }));
            template.Outline.Polycurve_Area.Vertices = vertices1;
            return template;
        }

        private dynamic addFeatureHoles(double x,double y,double radius)
        {

            string holeTemplate = "{\"Feature_Type\":\"Hole\",\"Entity_ID\":\"#" + generateId() + "\",\"Type\":\"Thru_Pin\",\"Plated\":\"True\",\"Shape_Type\":\"Round\",\"Outline\":{\"Circle\":{\"Entity_ID\":\"#" + generateId() + "\",\"Line_Font\":\"Solid\",\"XY_Loc\":[0,0],\"Radius\":\"0\"}},\"XY_Loc\":[],\"Rotation\":\"0.0\"}";
            dynamic template = JsonConvert.DeserializeObject<dynamic>(holeTemplate);
            JArray center = new JArray();
            center.Add(x);
            center.Add(y);
            template.XY_Loc = center;
            template.Outline.Circle.Radius = radius;
            return template;
        }
        private dynamic addPartInstance(ComponentAttribute c,string packageName)
        {
            string partTemplate = "";
            if (c.side)
                partTemplate = "{\"Type\":\"Electrical_Part_Instance\",\"Entity_ID\":\"" + c.componentID + "\",\"Part_Name\":\""+packageName+"\",\"Part_Number\":\""+ c.part_number + "\",\"In_BOM\":\"True\",\"Refdes\":\""+c.redef + "\",\"Lock\":\"None\",\"XY_Loc\":[],\"Side\":\"Top\",\"Rotation\":\"\",\"Mnt_Offset\":[0,0]}";
            else
                partTemplate = "{\"Type\":\"Electrical_Part_Instance\",\"Entity_ID\":\"" + c.componentID + "\",\"Part_Name\":\"" + packageName + "\",\"Part_Number\":\"" + c.part_number + "\",\"In_BOM\":\"True\",\"Refdes\":\"" + c.redef + "\",\"Lock\":\"None\",\"XY_Loc\":[],\"Side\":\"Bottom\",\"Rotation\":\"\",\"Mnt_Offset\":[0,0]}";
            dynamic template = JsonConvert.DeserializeObject<dynamic>(partTemplate);
            JArray XY_Loc = new JArray();
            XY_Loc.Add(c.location.X);
            XY_Loc.Add(c.location.Y);
            template.XY_Loc = XY_Loc;
            template.Rotation = c.rotation;
            return template;
        }
        private dynamic addPart(ComponentAttribute c, string packageName)
        {
            string partTemplate = "{\"Entity_ID\":\"" + generateId() + "\",\"Part_Name\":\"" + packageName + "\",\"Units\":\"Global\",\"Type\":\"Unspecified\",\"Shape\":{\"Extrusion\":{\"Entity_ID\":\"#" + generateId() + "\",\"Top_Height\":1.0,\"Bot_Height\":0,\"Outline\":{\"Polycurve_Area\":{\"Entity_ID\":\"#2512\",\"Line_Font \":\"Solid\",\"Line_Color\":[0.0,0.0,0.0],\"Fill_Color\":[0.0,0.0,0.0],\"Vertices\":[[-5.588000,4.542981,0.0],[2.540000,4.542981,0],[2.540000,-0.508000,0],[-5.588000,-0.508000,0],[-5.588000,4.542981,0.0]]}}}},\"Properties\":{\"Part_Number\":\""+c.part_number+"\"}}";
            dynamic template = JsonConvert.DeserializeObject<dynamic>(partTemplate);
            return template;
        }
        private void buttonSaveJSON_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                UpdateAllTransformations(TaskpaneIntegration.mSolidworksApplication);
                if (isIDFFile)
                {
                    saveIDFToJSON();
                    //MessageBox.Show("Save function not implemented for IDF files. ");
                }
                else
                {
                    boardJSON.Parts.Board_Part.Shape.Extrusion.Top_Height = (double)numericUpDownBoardHeight.Value;
                    saveDecalsFromBoard(TaskpaneIntegration.mSolidworksApplication, boardJSON);
                    saveJSONtoFile(boardJSON);
                }

            }
        }


        private void lblBoardColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog dlg = new ColorDialog
                {
                    Color = lblBoardColor.BackColor,
                    FullOpen = true
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblBoardColor.BackColor = dlg.Color;
                    changeBoardColor(lblBoardColor.BackColor, true, TaskpaneIntegration.mSolidworksApplication);
                }
            }
            catch (Exception)
            { }
        }

        private void lblHolesColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog dlg = new ColorDialog
                {
                    Color = lblHolesColor.BackColor,
                    FullOpen = true
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblHolesColor.BackColor = dlg.Color;
                    changeBoardColor(lblHolesColor.BackColor, false, TaskpaneIntegration.mSolidworksApplication);
                }
            }
            catch (Exception)
            { }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            if (FileOpened)
            {//rotateComponent((double)numericUpDownAngle.Value, comboAlongAxis.SelectedIndex, TaskpaneIntegration.mSolidworksApplication);
                if (checkRotateSelected.Checked)
                {
                    RotateAllSelectedComponent((double)numericUpDownAngle.Value, comboAlongAxis.SelectedIndex, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    RotateAllComponent(componentPackage, (double)numericUpDownAngle.Value, comboAlongAxis.SelectedIndex, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void buttonAutoAlign_Click_1(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                string componentPackage = (string)comboComponentList.SelectedItem;
                AutoAlignPackage(componentPackage, chkIsSMD.Checked, checkRotateSelected.Checked, TaskpaneIntegration.mSolidworksApplication);
            }
        }

        private void checkBoxTopOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                if (checkBoxTopOnly.Checked)
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    if (boardComponents.ContainsKey(componentPackage))
                    {
                        InsertNoteForPackage(componentPackage, TaskpaneIntegration.mSolidworksApplication);
                    }
                }
                else
                {
                    for (int i = 0; i < CurrentNotesTop.Count; i++)
                    {
                        removeNote(CurrentNotesTop[i], TaskpaneIntegration.mSolidworksApplication);
                    }
                }
            }
        }

        private void checkBoxBottomOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                if (checkBoxBottomOnly.Checked)
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    if (boardComponents.ContainsKey(componentPackage))
                    {
                        InsertNoteForPackage(componentPackage, TaskpaneIntegration.mSolidworksApplication);
                    }
                }
                else
                {
                    for (int i = 0; i < CurrentNotesBottom.Count; i++)
                    {
                        removeNote(CurrentNotesBottom[i], TaskpaneIntegration.mSolidworksApplication);
                    }
                    CurrentNotesBottom.Clear();
                }
            }
        }

        private void comboComponentList_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            string iconPath = "";
            if (e.Index >= 0)
            {
                string componentPackage = boardComponents.ElementAt(e.Index).Key;
                if (ComponentModels.ContainsKey(componentPackage))
                    if (ComponentModels[componentPackage].Length > 2)
                        iconPath = CheckIconPath;
                    else
                        iconPath = CrossIconPath;
                else
                    iconPath = CrossIconPath;

                using (Brush brush = new SolidBrush(e.ForeColor))
                {
                    // e.Graphics.FillRectangle(brush, e.Bounds.X, e.Bounds.Y, 14, 14);
                    Icon icon = new Icon(iconPath);
                    e.Graphics.DrawIcon(icon, e.Bounds.X, e.Bounds.Y);
                }

                using (Brush brush = new SolidBrush(e.ForeColor))
                {
                    e.Graphics.DrawString(componentPackage, e.Font, brush, e.Bounds.Left + 16, e.Bounds.Top);
                }
                e.DrawFocusRectangle();

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (FileOpened)
            {
                if (checkRotateSelected.Checked)
                {
                    translateComponentDragOperator(new double[] { convertUnit((double)numericUpDownIncrement.Value), 0, 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    TranslateAllComponent(componentPackage, new double[] { convertUnit((double)numericUpDownIncrement.Value), 0, 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void buttonXMinus_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                if (checkRotateSelected.Checked)
                {
                    translateComponentDragOperator(new double[] { -convertUnit((double)numericUpDownIncrement.Value), 0, 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    TranslateAllComponent(componentPackage, new double[] { -convertUnit((double)numericUpDownIncrement.Value), 0, 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void buttonYPlus_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                if (checkRotateSelected.Checked)
                {
                    translateComponentDragOperator(new double[] { 0, convertUnit((double)numericUpDownIncrement.Value), 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    TranslateAllComponent(componentPackage, new double[] { 0, convertUnit((double)numericUpDownIncrement.Value), 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void buttonYMinus_Click(object sender, EventArgs e)
        {

            if (FileOpened)
            {
                if (checkRotateSelected.Checked)
                {
                    translateComponentDragOperator(new double[] { 0, -convertUnit((double)numericUpDownIncrement.Value), 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    TranslateAllComponent(componentPackage, new double[] { 0, -convertUnit((double)numericUpDownIncrement.Value), 0 }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void buttonZPlus_Click(object sender, EventArgs e)
        {

            if (FileOpened)
            {
                if (checkRotateSelected.Checked)
                {
                    translateComponentDragOperator(new double[] { 0, 0, convertUnit((double)numericUpDownIncrement.Value) }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    TranslateAllComponent(componentPackage, new double[] { 0, 0, convertUnit((double)numericUpDownIncrement.Value) }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void buttonZMinus_Click(object sender, EventArgs e)
        {

            if (FileOpened)
            {
                if (checkRotateSelected.Checked)
                {
                    translateComponentDragOperator(new double[] { 0, 0, -convertUnit((double)numericUpDownIncrement.Value) }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
                else
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    TranslateAllComponent(componentPackage, new double[] { 0, 0, -convertUnit((double)numericUpDownIncrement.Value) }, checkBoxCollisionDetection.Checked, TaskpaneIntegration.mSolidworksApplication);
                }
            }
        }

        private void colorPresets_DrawItem(object sender, DrawItemEventArgs e)
        {

            if (FileOpened)
            {
                e.DrawBackground();
            int i = 0;
                if (e.Index >= 0)
                {
                    List<Bgra> bgras = applyPCBStyle(e.Index);
                    string styleName = (string)((dynamic)boardStyleJSON[e.Index]).StyleName;
                    for (i = 0; i < bgras.Count; i++)
                    {
                        using (Brush brush = new SolidBrush(Color.FromArgb((int)bgras[i].Alpha, (int)bgras[i].Red, (int)bgras[i].Green, (int)bgras[i].Blue)))
                        {
                            e.Graphics.FillRectangle(brush, e.Bounds.X + (i * 15), e.Bounds.Y, 14, 14);
                        }
                    }
                    using (Brush brush = new SolidBrush(e.ForeColor))
                    {
                        e.Graphics.DrawString(styleName, e.Font, brush, e.Bounds.Left + 2 + (i * 15), e.Bounds.Top);
                    }
                    e.DrawFocusRectangle();
                }
            }
        }

        private void buttonSaveLibrary_Click(object sender, EventArgs e)
        {
            
        }

        private void buttonLoadLibrary_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                string componentPackage = (string)comboComponentList.SelectedItem;
                //OpenFileDialog Openfile = new OpenFileDialog();
                //Openfile.Filter = "Library JSON  files (*.json) | *.json";
                //Openfile.Title = "Open Library JSON file";
                //if (Openfile.ShowDialog() == DialogResult.OK)
                //{
                string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
                string libraryJSONPath = @Path.Combine(directory, pluginName, libraryFileName);
                if (checkLoadSelected.Checked)
                    loadLibrary(@libraryJSONPath, componentPackage);
                else
                    loadLibrary(@libraryJSONPath, "");
                //}
                ComponentListSelectEvent();
            }
        }

        private void comboAlongAxis_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkLoadSelected_CheckedChanged(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                if (checkLoadSelected.Checked)
                {
                    string componentPackage = (string)comboComponentList.SelectedItem;
                    bool found = false;
                    if (boardComponents.Count > 0)
                    {
                        string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
                        string libraryJSONPath = @Path.Combine(directory, pluginName, libraryFileName);
                        if (File.Exists(libraryJSONPath))
                        {
                            ConcurrentDictionary<string, ComponentLibrary> libraryJSONList = JsonConvert.DeserializeObject<ConcurrentDictionary<string, ComponentLibrary>>(File.ReadAllText(@libraryJSONPath));
                            int count = libraryJSONList.Count;
                            foreach (var pair in boardComponents)
                            {
                                if (pair.Key == componentPackage)
                                {
                                    if (libraryJSONList.ContainsKey(pair.Key))
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (found)
                        buttonLoadLibrary.Enabled = true;
                    else
                        buttonLoadLibrary.Enabled = false;
                }
                else
                {
                    buttonLoadLibrary.Enabled = true;
                }
            }
        }

        private void chkIsSMD_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDownBoardHeight_ValueChanged(object sender, EventArgs e)
        {

        }

        private void buttonChangeHeight_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                DialogResult dialogResult = MessageBox.Show("Make sure the board thickness is compatible with the\nfabrication house and the overall electrical design.\n\nAre you sure you want to change the thickness? You can always change it back if needed.", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    setBoardThickness(convertUnit((double)numericUpDownBoardHeight.Value), TaskpaneIntegration.mSolidworksApplication);
                }
                
            }
        }

        private void buttonGerberFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog Openfile = new OpenFileDialog
            {
                Filter = "ZIP Files (*.zip)|*.zip",
                Title = "Open zipped Gerber Files"
            };
            if (Openfile.ShowDialog() == DialogResult.OK)
            {
                gerberFile = Openfile.FileName;
                Progress P = new Progress(@gerberFile, lblMaskColor.BackColor, lblSilkColor.BackColor, lblPadColor.BackColor, lblTraceColor.BackColor);
                P.Text = "Gerber Rendering Progress";
                P.StartThread(0);
                P.ShowDialog();
                P = null;
                string directoryPath = Path.GetDirectoryName(boardPartPath);
                Mat imgPng = new Mat();
                CvInvoke.Imdecode(UtilityClass.TopDecal, ImreadModes.Unchanged, imgPng);
                Image<Bgra, byte> TopDecalImage = imgPng.ToImage<Bgra, byte>();
                //TopDecalImage = TopDecalImage.Flip(FlipType.Vertical);
                deleteFile(topDecal);
                topDecal = RandomString(10) + "top_decal.png";
                TopDecalImage.PyrUp().Save(@Path.Combine(directoryPath, topDecal));
                topDecal = @Path.Combine(directoryPath, topDecal);
                UtilityClass.TopDecal = null;
                imgPng.Dispose();
                TopDecalImage.Dispose();

                imgPng = new Mat();
                CvInvoke.Imdecode(UtilityClass.BottomDecal, ImreadModes.Unchanged, imgPng);
                Image<Bgra, byte> BottomDecalImage = imgPng.ToImage<Bgra, byte>();
                BottomDecalImage = BottomDecalImage.Flip(FlipType.Vertical).Flip(FlipType.Horizontal);
                deleteFile(bottomDecal);
                bottomDecal = RandomString(10) + "bottom_decal.png";
                BottomDecalImage.PyrUp().Save(@Path.Combine(directoryPath, bottomDecal));
                bottomDecal = @Path.Combine(directoryPath, bottomDecal);
                UtilityClass.BottomDecal = null;
                imgPng.Dispose();
                BottomDecalImage.Dispose();

                addDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
            }
        }

        private void buttonDecalImages_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not yet implemented");
        }
    }
}
