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

namespace LP.SolidWorks.BlankAddin
{

    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]

    public partial class TaskpaneHostUI : UserControl
    {
        double PI = 3.14159;
        double RadPerDeg = (3.14159 / 180);
        private string boardPartPath = "";
        private string boardPartID = "";
        private string boardSketch = "";
        dynamic boardJSON = null;
        JArray boardStyleJSON = null;
        double boardHeight = 0;
        double[] boardFirstPoint = new double[2];
        private int max_width = 8000;
        private int max_height = 8000;
        private string topDecal = "";
        private string bottomDecal = "";
        private static Random random = new Random();
        private double boardPhysicalHeight = 0;
        private double boardPhysicaWidth = 0;
        private messageFrm messageForm = new messageFrm();
        double[] boardOrigin = new double[3];
        private AssemblyDoc globalSwAssy = null;
        private bool FileOpened = false;
        private class ComponentAttribute
        {
            public PointF location;
            public bool side;      //true is top and false is bottom
            public float rotation;
            public string modelID;
            public string modelPath;
            public string componentID;
            public string noteID;

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
        private void populateStyleComboBox()
        {

            string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
            string boardStyleJSONFile = @Path.Combine(directory, pluginName, PCBStylePath);
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
                        {
                            ComponentTransformations[c.componentID] = tx;
                        }
                        else
                        {
                            ComponentTransformations.TryAdd(c.componentID, tx);
                        }

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
            boardJSON.Parts.Board_Part.Shape.Extrusion.Top_Height = (double)numericUpDownBoardHeight.Value;
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
        private string createBoardFromIDF(string emnfile, SldWorks mSolidworksApplication)
        {
            bool boolstatus = false;
            readBoardData(emnfile);
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
            string directoryPath = Path.GetDirectoryName(emnfile);
            string boardPath = Path.Combine(directoryPath, "board.SLDPRT");
            //swModel.ViewZoomtofit2();
            //swModel.ShowNamedView2("*Front", 1);
            long status = swModel.SaveAs3(boardPath, 0, 2);
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);
            return boardPath;
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
        public static Image<Bgra, Byte> Overlay(Image<Bgra, Byte> target, Image<Bgra, Byte> overlay)
        {
            Bitmap bmp = target.Bitmap;
            Graphics gra = Graphics.FromImage(bmp);
            gra.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
            gra.DrawImage(overlay.Bitmap, new Point(0, 0));

            return target;
        }
        private void writeDecalImage(dynamic boardJSON, bool side, Bgra maskColor, Bgra traceColor, Bgra padColor, Bgra silkColor)
        {
            dynamic features = boardJSON.Parts.Board_Part.Features;
            int scaling = 1;
            Image<Gray, byte> trace = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> silkscreen = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> polygon_positive = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> polygon_border = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> polygon_negative = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> stops = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> pads = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> holes = new Image<Gray, byte>(max_width, max_height);
            holes.SetValue(new Gray(255));
            //int gray_value = 5;
            //int gray_value_neg = 50;
            int tx = max_width / 2;
            int ty = max_height / 2;
            int prev_polygon_width = -1;

            Image<Gray, byte> My_Image_gray = null;
            Rectangle r;
            {
                List<Point> pointarray = new List<Point>();
                JArray vertices = (JArray)boardJSON.Parts.Board_Part.Shape.Extrusion.Outline.Polycurve_Area.Vertices;
                for (int i = 1; i < vertices.Count; i++)
                {
                    pointarray.Add(new Point(tx + Convert.ToInt32((double)vertices[i][0].ToObject(typeof(double))), ty + Convert.ToInt32((double)vertices[i][1].ToObject(typeof(double)))));
                    // textBox1.Text += "\r\n" + Convert.ToInt32((double)vertices[i][0].ToObject(typeof(double)) * scaling) + "," + Convert.ToInt32((double)vertices[i][1].ToObject(typeof(double)) * scaling);
                }
                My_Image_gray = new Image<Gray, byte>(max_width, max_height);
                My_Image_gray.DrawPolyline(pointarray.ToArray(), true, new Gray(255));
                r = CvInvoke.BoundingRectangle(My_Image_gray.Mat);
                boardPhysicalHeight = r.Height - 1;
                boardPhysicaWidth = r.Width - 1.5;
                int maxDim = r.Width > r.Height ? r.Width : r.Height;
                double s = max_width / (2 * maxDim);
                scaling = Convert.ToInt32(Math.Floor(s));
                pointarray.Clear();
                for (int i = 1; i < vertices.Count; i++)
                {
                    pointarray.Add(new Point(tx + Convert.ToInt32((double)vertices[i][0].ToObject(typeof(double)) * scaling), ty + Convert.ToInt32((double)vertices[i][1].ToObject(typeof(double)) * scaling)));
                    // textBox1.Text += "\r\n" + Convert.ToInt32((double)vertices[i][0].ToObject(typeof(double)) * scaling) + "," + Convert.ToInt32((double)vertices[i][1].ToObject(typeof(double)) * scaling);
                }
                VectorOfPoint vp = new VectorOfPoint(pointarray.ToArray());
                VectorOfVectorOfPoint vvp = new VectorOfVectorOfPoint(vp);
                My_Image_gray.SetZero();
                CvInvoke.FillPoly(My_Image_gray, vvp, new MCvScalar(255), LineType.AntiAlias);
                //My_Image_gray.DrawPolyline(pointarray.ToArray(), true, new Gray(255));
                r = CvInvoke.BoundingRectangle(My_Image_gray.Mat);
            }
            //CvInvoke.cvSetImageROI(My_Image, r);
            //for (int i = 0; i < features.Count; i++)
            int feature_count = (int)features.Count;
            Parallel.For(0, feature_count, i =>
            {
                dynamic currentFeature = features[i];
                string featureType = (string)currentFeature.Feature_Type;

                if (featureType != null)
                {
                    if (featureType.ToUpper() == "HOLE")
                    {
                        if (currentFeature.Outline.Circle != null)
                        {
                            JArray center = currentFeature.XY_Loc;
                            int cx = tx + Convert.ToInt32((double)center[0].ToObject(typeof(double)) * scaling);
                            int cy = ty + Convert.ToInt32((double)center[1].ToObject(typeof(double)) * scaling);
                            double rx = (double)currentFeature.Outline.Circle.Radius.ToObject(typeof(double));
                            int circleRadius = Convert.ToInt32(rx * scaling);
                            lock ("MyLock")
                            {
                                CvInvoke.Circle(holes, new Point(cx, cy), circleRadius, new MCvScalar(0), -1, LineType.AntiAlias);
                            }
                        }
                    }
                    if (featureType.ToUpper() == "TRACE")
                    {
                        if (currentFeature.Curve != null)
                        {
                            string Pad_Type = "TRACE";
                            if (currentFeature.Curve.Polyline.XY_Pts != null)
                            {
                                //handle polycurve area
                                if (currentFeature.Pad_Type != null)
                                    Pad_Type = (string)currentFeature.Pad_Type;
                                JArray vertices = currentFeature.Curve.Polyline.XY_Pts;
                                int linewidth = Convert.ToInt32((double)currentFeature.Curve.Polyline.Width * scaling);
                                List<Point> pointarray = new List<Point>();
                                for (int j = 0; j < vertices.Count; j++)
                                {
                                    int px = tx + Convert.ToInt32((double)vertices[j][0].ToObject(typeof(double)) * scaling);
                                    int py = ty + Convert.ToInt32((double)vertices[j][1].ToObject(typeof(double)) * scaling);
                                    pointarray.Add(new Point(px, py));
                                }
                                string featureLayer = (string)currentFeature.Layer;
                                lock ("MyLock")
                                {
                                    if (side)
                                    {
                                        if (Pad_Type.ToUpper() == "STOP")
                                        {
                                            stops.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                        }
                                        else
                                        {
                                            if (featureLayer.ToUpper() == "CONDUCTOR_TOP")
                                            {
                                                // My_Image.DrawPolyline(pointarray.ToArray(), false, traceColor, linewidth, LineType.AntiAlias);
                                                trace.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                            }
                                            else if (featureLayer.ToUpper() == "SILKSCREEN_TOP")
                                            {
                                                //My_Image.DrawPolyline(pointarray.ToArray(), false, silkColor, linewidth, LineType.AntiAlias);
                                                silkscreen.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (Pad_Type.ToUpper() == "STOP")
                                        {
                                            stops.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                        }
                                        else
                                        {
                                            if (featureLayer.ToUpper() == "CONDUCTOR_BOTTOM")
                                            {
                                                // My_Image.DrawPolyline(pointarray.ToArray(), false, traceColor, linewidth, LineType.AntiAlias);
                                                trace.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                            }
                                            else if (featureLayer.ToUpper() == "SILKSCREEN_BOTTOM")
                                            {
                                                // My_Image.DrawPolyline(pointarray.ToArray(), false, silkColor, linewidth, LineType.AntiAlias);
                                                silkscreen.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                            }
                                        }
                                    }
                                }
                                //textBox1.Text += "\r\n" + currentFeature.Curve.Polyline.XY_Pts + " , " + currentFeature.Curve.Polyline.Width;
                            }
                        }
                        else if (currentFeature.Geometry != null)
                        {
                            string featureLayer = (string)currentFeature.Layer;
                            JArray padLoc = currentFeature.XY_Loc;
                            int cx = tx + Convert.ToInt32((double)padLoc[0].ToObject(typeof(double)) * scaling);
                            int cy = ty + Convert.ToInt32((double)padLoc[1].ToObject(typeof(double)) * scaling);
                            int linewidth = Convert.ToInt32((double)currentFeature.Geometry.Polycurve_Area.Width * scaling);
                            int isPositive = (int)currentFeature.Geometry.Polycurve_Area.Positive;

                            if (currentFeature.Geometry.Polycurve_Area != null)
                            {
                                //handle polycurve area
                                JArray vertices = currentFeature.Geometry.Polycurve_Area.Vertices;
                                Image<Gray, byte> Hatch_plane = new Image<Gray, byte>(max_width, max_height);
                                List<Point> pointarray = new List<Point>();
                                int hatched = (int)currentFeature.Geometry.Polycurve_Area.Hatched.ToObject(typeof(int)); ;
                                if (hatched == 1)
                                {

                                    //if (prev_polygon_width != linewidth)
                                    //{
                                    prev_polygon_width = linewidth;
                                    int spacing = Convert.ToInt32((double)currentFeature.Geometry.Polycurve_Area.Spacing * scaling);
                                    //Hatch_plane = makeHatchedPolygon(max_width, max_height, prev_polygon_width, spacing);
                                    List<Point> linepoints = new List<Point>();
                                    int x = 0, y = 0;
                                    for (int c = 0; c < max_width; c = c + spacing)
                                    {
                                        if (x % 2 == 0)
                                        {
                                            linepoints.Add(new Point(x * spacing, 0));
                                            linepoints.Add(new Point(x * spacing, max_height));
                                        }
                                        else
                                        {
                                            linepoints.Add(new Point(x * spacing, max_height));
                                            linepoints.Add(new Point(x * spacing, 0));
                                        }
                                        x++;
                                    }
                                    Hatch_plane.DrawPolyline(linepoints.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                    linepoints.Clear();
                                    for (int c = 0; c < max_height; c = c + spacing)
                                    {
                                        if (y % 2 == 0)
                                        {
                                            linepoints.Add(new Point(0, y * spacing));
                                            linepoints.Add(new Point(max_width, y * spacing));
                                        }
                                        else
                                        {
                                            linepoints.Add(new Point(max_width, y * spacing));
                                            linepoints.Add(new Point(0, y * spacing));
                                        }
                                        y++;
                                    }
                                    Hatch_plane.DrawPolyline(linepoints.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                    //}
                                }
                                for (int j = 0; j < vertices.Count; j++)
                                {
                                    int px = Convert.ToInt32((double)vertices[j][0].ToObject(typeof(double)) * scaling);
                                    int py = Convert.ToInt32((double)vertices[j][1].ToObject(typeof(double)) * scaling);
                                    pointarray.Add(new Point(cx + px, cy + py));
                                }
                                {
                                    VectorOfPoint vp = new VectorOfPoint(pointarray.ToArray());
                                    VectorOfVectorOfPoint vvp = new VectorOfVectorOfPoint(vp);
                                    vp.Clear();
                                    vp = new VectorOfPoint(pointarray.ToArray());
                                    vvp.Clear();
                                    vvp = new VectorOfVectorOfPoint(vp);
                                    lock ("MyLock")
                                    {
                                        if (side)
                                        {
                                            if (featureLayer.ToUpper() == "CONDUCTOR_TOP")
                                            {
                                                if (isPositive == 1)
                                                {
                                                    if (hatched == 1)
                                                    {
                                                        //fillPolygonHatched(ref polygon_positive, Hatch_plane, pointarray);
                                                        Image<Gray, byte> polygon_gray = new Image<Gray, byte>(polygon_positive.Width, polygon_positive.Height);
                                                        VectorOfPoint vp1 = new VectorOfPoint(pointarray.ToArray());
                                                        VectorOfVectorOfPoint vvp1 = new VectorOfVectorOfPoint(vp); polygon_gray.SetValue(new Gray(0));
                                                        CvInvoke.FillPoly(polygon_gray, vvp1, new MCvScalar(255), LineType.EightConnected);
                                                        Hatch_plane.Copy(polygon_positive, polygon_gray);
                                                        //polygon_gray.Dispose();
                                                    }
                                                    else
                                                    //    CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(traceColor.Blue, traceColor.Green, traceColor.Red, traceColor.Alpha), LineType.AntiAlias);
                                                    //fillPolygonHatched(ref My_Image, pointarray, linewidth, ((6 * scaling)/10), new MCvScalar(traceColor.Blue, traceColor.Green, traceColor.Red, traceColor.Alpha));
                                                    {
                                                        //if (gray_value >= 255)
                                                        //    gray_value = 5;
                                                        CvInvoke.FillPoly(polygon_positive, vvp, new MCvScalar(255), LineType.AntiAlias);
                                                        //gray_value = gray_value + 50;
                                                    }

                                                }
                                                else
                                                {
                                                    //if (gray_value_neg >= 255)
                                                    //    gray_value_neg = 50;
                                                    //CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(maskColor.Blue, maskColor.Green, maskColor.Red, maskColor.Alpha), LineType.AntiAlias);
                                                    CvInvoke.FillPoly(polygon_negative, vvp, new MCvScalar(255), LineType.AntiAlias);
                                                    //gray_value_neg = gray_value_neg + 50;
                                                }
                                                //My_Image.DrawPolyline(pointarray.ToArray(), false, traceColor, linewidth, LineType.AntiAlias);
                                                polygon_border.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                                polygon_border.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                            }
                                        }
                                        else
                                        {
                                            if (featureLayer.ToUpper() == "CONDUCTOR_BOTTOM")
                                            {
                                                if (isPositive == 1)
                                                {
                                                    if (hatched == 1)
                                                    {
                                                        //fillPolygonHatched(ref polygon_positive, Hatch_plane, pointarray);
                                                        Image<Gray, byte> polygon_gray = new Image<Gray, byte>(polygon_positive.Width, polygon_positive.Height);
                                                        VectorOfPoint vp1 = new VectorOfPoint(pointarray.ToArray());
                                                        VectorOfVectorOfPoint vvp1 = new VectorOfVectorOfPoint(vp); polygon_gray.SetValue(new Gray(0));
                                                        CvInvoke.FillPoly(polygon_gray, vvp1, new MCvScalar(255), LineType.EightConnected);
                                                        Hatch_plane.Copy(polygon_positive, polygon_gray);
                                                        //polygon_gray.Dispose();
                                                    }
                                                    else
                                                        //     CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(traceColor.Blue, traceColor.Green, traceColor.Red, traceColor.Alpha), LineType.AntiAlias);
                                                        //fillPolygonHatched(ref My_Image, pointarray, linewidth, ((6 * scaling)/10), new MCvScalar(traceColor.Blue, traceColor.Green, traceColor.Red, traceColor.Alpha));
                                                        CvInvoke.FillPoly(polygon_positive, vvp, new MCvScalar(255), LineType.AntiAlias);

                                                }
                                                else
                                                {
                                                    //CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(maskColor.Blue, maskColor.Green, maskColor.Red, maskColor.Alpha), LineType.AntiAlias);
                                                    CvInvoke.FillPoly(polygon_negative, vvp, new MCvScalar(255), LineType.AntiAlias);
                                                }
                                                //My_Image.DrawPolyline(pointarray.ToArray(), false, traceColor, linewidth, LineType.AntiAlias);
                                                polygon_border.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                                polygon_border.DrawPolyline(pointarray.ToArray(), false, new Gray(255), linewidth, LineType.AntiAlias);
                                            }
                                        }
                                    }
                                }
                                Hatch_plane.Dispose();
                                Hatch_plane = null;
                            }
                            //if (currentFeature.Geometry.Circle != null)
                            //{
                            //    //Handle Circle
                            //    double rx = currentFeature.Geometry.Circle.Radius;
                            //    int circleRadius = Convert.ToInt32(rx * scaling);
                            //    // lock ("MyLock")
                            //    // {
                            //    if (side)
                            //    {
                            //        if (featureLayer.ToUpper() == "CONDUCTOR_TOP")
                            //            CvInvoke.Circle(My_Image, new Point(cx, cy), circleRadius, new MCvScalar(traceColor.Blue, traceColor.Green, traceColor.Red, traceColor.Alpha), -1, LineType.AntiAlias);
                            //        //CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(padColor.Blue, padColor.Green, padColor.Red, padColor.Alpha), LineType.AntiAlias);
                            //    }
                            //    else
                            //    {
                            //        if (featureLayer.ToUpper() == "CONDUCTOR_BOTTOM")
                            //            CvInvoke.Circle(My_Image, new Point(cx, cy), circleRadius, new MCvScalar(traceColor.Blue, traceColor.Green, traceColor.Red, traceColor.Alpha), -1, LineType.AntiAlias);
                            //    }
                            //    // }
                            //}
                        }
                    }
                    if (featureType.ToUpper() == "PAD")
                    {
                        if (currentFeature.XY_Loc != null)
                        {
                            string featureLayer = (string)currentFeature.Layer.ToObject(typeof(string));
                            JArray padLoc = currentFeature.XY_Loc;
                            int cx = tx + Convert.ToInt32((double)padLoc[0].ToObject(typeof(double)) * scaling);
                            int cy = ty + Convert.ToInt32((double)padLoc[1].ToObject(typeof(double)) * scaling);
                            bool isStop = false;
                            if (currentFeature.Pad_Type != null)
                            {
                                string Pad_Type = (string)currentFeature.Pad_Type;
                                if (Pad_Type.ToUpper() == "STOP")
                                    isStop = true;
                                else
                                    isStop = false;
                            }
                            if (currentFeature.Geometry.Polycurve_Area != null)
                            {
                                //handle polycurve area
                                JArray vertices = currentFeature.Geometry.Polycurve_Area.Vertices;
                                List<Point> pointarray = new List<Point>();
                                for (int j = 0; j < vertices.Count; j++)
                                {
                                    int px = Convert.ToInt32((double)vertices[j][0].ToObject(typeof(double)) * scaling);
                                    int py = Convert.ToInt32((double)vertices[j][1].ToObject(typeof(double)) * scaling);
                                    pointarray.Add(new Point(cx + px, cy + py));
                                }
                                {
                                    VectorOfPoint vp = new VectorOfPoint(pointarray.ToArray());
                                    VectorOfVectorOfPoint vvp = new VectorOfVectorOfPoint(vp);
                                    vp.Clear();
                                    vp = new VectorOfPoint(pointarray.ToArray());
                                    vvp.Clear();
                                    vvp = new VectorOfVectorOfPoint(vp);
                                    lock ("MyLock")
                                    {
                                        if (side)
                                        {
                                            if (featureLayer.ToUpper() == "CONDUCTOR_TOP")
                                            {
                                                //CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(padColor.Blue, padColor.Green, padColor.Red, padColor.Alpha), LineType.AntiAlias);
                                                if (isStop)
                                                    CvInvoke.FillPoly(stops, vvp, new MCvScalar(255), LineType.AntiAlias);
                                                else
                                                    CvInvoke.FillPoly(pads, vvp, new MCvScalar(255), LineType.AntiAlias);
                                            }
                                        }
                                        else
                                        {
                                            if (featureLayer.ToUpper() == "CONDUCTOR_BOTTOM")
                                            {
                                                //CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(padColor.Blue, padColor.Green, padColor.Red, padColor.Alpha), LineType.AntiAlias);
                                                if (isStop)
                                                    CvInvoke.FillPoly(stops, vvp, new MCvScalar(255), LineType.AntiAlias);
                                                else
                                                    CvInvoke.FillPoly(pads, vvp, new MCvScalar(255), LineType.AntiAlias);
                                            }
                                        }
                                    }
                                }
                            }
                            else if (currentFeature.Geometry.Circle != null)
                            {
                                //Handle Circle
                                double rx = (double)currentFeature.Geometry.Circle.Radius.ToObject(typeof(double));
                                int circleRadius = Convert.ToInt32(rx * scaling);
                                lock ("MyLock")
                                {
                                    if (side)
                                    {
                                        if (featureLayer.ToUpper() == "CONDUCTOR_TOP")
                                        {
                                            //CvInvoke.Circle(My_Image, new Point(cx, cy), circleRadius, new MCvScalar(padColor.Blue, padColor.Green, padColor.Red, padColor.Alpha), -1, LineType.AntiAlias);
                                            if (isStop)
                                                CvInvoke.Circle(stops, new Point(cx, cy), circleRadius, new MCvScalar(255), -1, LineType.AntiAlias);
                                            else
                                                CvInvoke.Circle(pads, new Point(cx, cy), circleRadius, new MCvScalar(255), -1, LineType.AntiAlias);
                                        }
                                        //CvInvoke.FillPoly(My_Image, vvp, new MCvScalar(padColor.Blue, padColor.Green, padColor.Red, padColor.Alpha), LineType.AntiAlias);
                                    }
                                    else
                                    {
                                        if (featureLayer.ToUpper() == "CONDUCTOR_BOTTOM")
                                        {
                                            //CvInvoke.Circle(My_Image, new Point(cx, cy), circleRadius, new MCvScalar(padColor.Blue, padColor.Green, padColor.Red, padColor.Alpha), -1, LineType.AntiAlias);
                                            if (isStop)
                                                CvInvoke.Circle(stops, new Point(cx, cy), circleRadius, new MCvScalar(255), -1, LineType.AntiAlias);
                                            else
                                                CvInvoke.Circle(pads, new Point(cx, cy), circleRadius, new MCvScalar(255), -1, LineType.AntiAlias);
                                        }
                                    }
                                }
                            }


                            //textBox1.Text += "\r\n" + currentFeature.Curve.Polyline.XY_Pts + " , " + currentFeature.Curve.Polyline.Width;
                        }
                    }
                }
            });
            polygon_positive = polygon_positive.Sub(polygon_negative);
            polygon_positive = polygon_positive.Add(polygon_border);
            polygon_negative.Dispose();
            polygon_negative = null;
            polygon_border.Dispose();
            polygon_border = null;
            trace = trace.Add(polygon_positive);
            trace = trace.Add(pads);
            polygon_positive.Dispose();
            polygon_positive = null;
            Image<Gray, byte> temp_Image2 = new Image<Gray, byte>(max_width, max_height);
            Image<Gray, byte> mask = new Image<Gray, byte>(max_width, max_height);
            mask = My_Image_gray.And(stops.Erode(1).Not());
            mask = mask.And(holes);
            trace = trace.And(holes);
            temp_Image2 = trace.And(stops);
            silkscreen = silkscreen.And(stops.Not());
            CvInvoke.cvSetImageROI(trace, r);
            CvInvoke.cvSetImageROI(silkscreen, r);
            CvInvoke.cvSetImageROI(pads, r);
            CvInvoke.cvSetImageROI(stops, r);
            CvInvoke.cvSetImageROI(temp_Image2, r);
            CvInvoke.cvSetImageROI(mask, r);
            Image<Bgra, byte> output_Image = new Image<Bgra, byte>(r.Width, r.Height);
            Image<Bgra, byte> temp_Image1 = new Image<Bgra, byte>(r.Width, r.Height);
            Image<Bgra, byte> overlayed_Image = new Image<Bgra, byte>(r.Width, r.Height);
            Image<Bgra, byte> trace_Image = new Image<Bgra, byte>(r.Width, r.Height);

            temp_Image1.SetValue(traceColor);
            temp_Image1.Copy(trace_Image, trace);
            if (silkColor.Alpha > 0)
            {
                temp_Image1.SetValue(silkColor);
                temp_Image1.Copy(trace_Image, silkscreen);
            }
            temp_Image1.SetValue(padColor);
            temp_Image1.Copy(trace_Image, temp_Image2);
            //temp_Image1.Copy(trace_Image, pads);
            //CvInvoke.cvSetImageROI(My_Image, r);
            CvInvoke.cvSetImageROI(My_Image_gray, r);
            temp_Image1.SetValue(maskColor);
            temp_Image1.Copy(overlayed_Image, mask);
            overlayed_Image = Overlay(overlayed_Image, trace_Image);
            overlayed_Image.Copy(output_Image, My_Image_gray);
            temp_Image1.Dispose();
            temp_Image1 = null;
            trace_Image.Dispose();
            trace_Image = null;
            pads.Dispose();
            pads = null;
            trace.Dispose();
            trace = null;
            stops.Dispose();
            stops = null;
            silkscreen.Dispose();
            silkscreen = null;
            mask.Dispose();
            mask = null;
            temp_Image2.Dispose();
            temp_Image2 = null;
            My_Image_gray.Dispose();
            My_Image_gray = null;
            overlayed_Image.Dispose();
            overlayed_Image = null;
            string directoryPath = Path.GetDirectoryName(boardPartPath);
            if (!side)
            {
                //output_Image = output_Image.Flip(FlipType.Horizontal);
                deleteFile(bottomDecal);
                bottomDecal = RandomString(10) + "bottom_decal.png";
                output_Image.PyrUp().Save(@Path.Combine(directoryPath, bottomDecal));
            }
            else
            {
                output_Image = output_Image.Flip(FlipType.Vertical);
                deleteFile(topDecal);
                topDecal = RandomString(10) + "top_decal.png";
                output_Image.PyrUp().Save(@Path.Combine(directoryPath, topDecal));
            }
            output_Image.Dispose();
            output_Image = null;
            GC.Collect();
        }
        private void deleteFile(string filePath)
        {
            if (@filePath.Length > 1)
                if (File.Exists(@filePath))
                {
                    File.Delete(@filePath);
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
        private void addDecalsToBoard(SldWorks mSolidworksApplication)
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

            //boolstatus = swModelDocExt.SelectByID2("BOARD", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            //PartDoc swCompModel = (PartDoc)swModel;
            //Feature swFeat = swCompModel.FeatureByName("Base-Extrude");
            //string test = swFeat.Name;
            //ExtrudeFeatureData2 swExtrudeData = (ExtrudeFeatureData2)swFeat.GetDefinition();
            //swExtrudeData.AccessSelections(swCompModel, null);


            boolstatus = swModelDocExt.SelectByID2("", "FACE", boardOrigin[0], boardOrigin[1], boardOrigin[2]+(boardHeight/2), false, 0, null, 0);
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
            swMaterial.Width = convertUnit(boardPhysicaWidth);
            //swMaterial.FitHeight = true;
            //swMaterial.FixedAspectRatio = true;
            swMaterial.FitWidth = true;
            //swMaterial.
            boolstatus = swModelDocExt.AddDecal(swDecal, out nDecalID);

            boolstatus = swModelDocExt.SelectByID2("", "FACE", boardOrigin[0], boardOrigin[1], boardOrigin[2] - (boardHeight / 2), false, 0, null, 0);
            //boolstatus = swModelDocExt.SelectByRay(boardFirstPoint[0], boardFirstPoint[1], (boardHeight / 2), 0, 0, 1, (boardHeight / 2), (int)swSelectType_e.swSelFACES, false, 0, (int)swSelectOption_e.swSelectOptionDefault);     //sw>2017
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
            swMaterial.Width = convertUnit(boardPhysicaWidth);
            swMaterial.FitWidth = true;
            boolstatus = swModelDocExt.AddDecal(swDecal, out nDecalID);
            swModel.Visible = false;
            status = swModel.SaveAs3(boardPartPath, 0, 2);
            ModelTitle = swModel.GetTitle();
            swModel = null;
            mSolidworksApplication.CloseDoc(ModelTitle);

        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog Openfile = new OpenFileDialog
            {
                Filter = "Board JSON  files (*.json) | *.json",
                Title = "Open Board JSON file"
            };
            if (Openfile.ShowDialog() == DialogResult.OK)
            {
                ComponentTransformations.Clear();
                boardComponents.Clear();
                ComponentModels.Clear();
                CurrentNotesTop.Clear();
                CurrentNotesBottom.Clear();
                comboComponentList.Items.Clear();
                checkBoxTopOnly.Checked = false;
                checkBoxBottomOnly.Checked = false;
                messageForm.Text = "Please Wait";
                messageForm.labelMessage.Text = "Processing Board Outline. . . ";
                messageForm.Show(this);
                boardPartPath = createBoardFromIDF(@Openfile.FileName, TaskpaneIntegration.mSolidworksApplication);
                boardPartID = insertPart(boardPartPath, 0, 0, true, true, 0, 2, "", TaskpaneIntegration.mSolidworksApplication, true);
                messageForm.labelMessage.Text = "Processing Traces and Silkscreen. . . ";
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
                //writeDecalImage(boardJSON, true, bgras[0], bgras[1], bgras[2], bgras[3]);
                //writeDecalImage(boardJSON, false, bgras[0], bgras[1], bgras[2], bgras[3]);
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
                int errors = 0;
                ModelDoc2 swModel = (ModelDoc2)TaskpaneIntegration.mSolidworksApplication.ActiveDoc;
                string AssemblyTitle = swModel.GetTitle();
                swModel = TaskpaneIntegration.mSolidworksApplication.ActivateDoc3(AssemblyTitle, true, (int)swRebuildOnActivation_e.swUserDecision, errors);
                globalSwAssy = (AssemblyDoc)swModel;
                //if (selectionEventHandler == false)
                //{
                //    selectionEventHandler = true;
                globalSwAssy.UserSelectionPostNotify += SwAssy_UserSelectionPostNotify;
                globalSwAssy.DeleteItemNotify += SwAssy_DeleteItemNotify;
                globalSwAssy.RenameItemNotify += SwAssy_RenameItemNotify;
                globalSwAssy.DestroyNotify2 += SwAssy_DestroyNotify2;
                FileOpened = true;
                //}
                //else
                //{
                //    swAssy.UserSelectionPostNotify -= SwAssy_UserSelectionPostNotify;
                //    swAssy.UserSelectionPostNotify += SwAssy_UserSelectionPostNotify;

                //    swAssy.DeleteItemNotify -= SwAssy_DeleteItemNotify;
                //    swAssy.DeleteItemNotify += SwAssy_DeleteItemNotify;
                //    //mSolidworksApplication.CommandCloseNotify -= MSolidworksApplication_CommandCloseNotify;
                //    //mSolidworksApplication.CommandCloseNotify += MSolidworksApplication_CommandCloseNotify;
                //}

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
                writeDecalImage(boardJSON, true, bgras[0], bgras[1], bgras[2], bgras[3]);
                //    else if (i == 1)
                writeDecalImage(boardJSON, false, bgras[0], bgras[1], bgras[2], bgras[3]);
                // });
                addDecalsToBoard(TaskpaneIntegration.mSolidworksApplication);
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

        private void buttonSaveJSON_Click(object sender, EventArgs e)
        {
            if (FileOpened)
            {
                UpdateAllTransformations(TaskpaneIntegration.mSolidworksApplication);
                saveJSONtoFile(boardJSON);
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
                    changeBoardColor(dlg.Color, true, TaskpaneIntegration.mSolidworksApplication);
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
                    changeBoardColor(dlg.Color, false, TaskpaneIntegration.mSolidworksApplication);
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
            if (FileOpened)
            {
                //getLibraryTransform();
                //saveLibraryasProgramData();
                setBoardThickness(0.02, TaskpaneIntegration.mSolidworksApplication);
            }
            //saveLibrary();
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
    }
}
