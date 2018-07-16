using Emgu.CV.Structure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LP.SolidWorks.ECADWorksAddin
{
    public partial class Progress : Form, GerberLibrary.ProgressLog
    {
        private List<string> Files;
        private Color SolderMaskColor;
        private Color SilkScreenColor;
        private Color CopperColor;
        private Color TracesColor;
        private dynamic boardJSON;
        private Bgra maskColor;
        private Bgra traceColor;
        private Bgra padColor;
        private Bgra silkColor;

        public Progress(dynamic _boardJSON, Bgra _maskColor, Bgra _traceColor, Bgra _padColor, Bgra _silkColor)
        {
            boardJSON = _boardJSON;
            maskColor = _maskColor;
            traceColor = _traceColor;
            padColor = _padColor;
            silkColor = _silkColor;

            InitializeComponent();
            progressBar1.Value = 0;
        }
        public Progress(string zippedGerber, Color _SolderMaskColor, Color _SilkScreenColor, Color _CopperColor, Color _tracescolor)
        {
            SolderMaskColor = _SolderMaskColor;
            SilkScreenColor = _SilkScreenColor;
            CopperColor = _CopperColor;
            TracesColor = _tracescolor;

            InitializeComponent();
            progressBar1.Value = 0;
            Files = new List<string>();
            Files.Add(zippedGerber);
        }
        Thread T = null;
        //renderType = 0 for gerber , 1 for JSON
        internal void StartThread(int renderType)
        {
            if (renderType == 0)
            {
                T = new Thread(new ThreadStart(renderGerber));
                T.Start();
            }
            if (renderType == 1)
            {
                T = new Thread(new ThreadStart(renderJson));
                T.Start();
            }
        }
        public void renderJson()
        {
            SetProgress("JSON data rendering started", 0);
            JSONLibrary.generateDecalImage(boardJSON,true, maskColor, traceColor, padColor, silkColor,this,50,0);
            //SetProgress("Finished Rendering Top layer", 0);
            JSONLibrary.generateDecalImage(boardJSON, false, maskColor, traceColor, padColor, silkColor,this,50,50);
            SetProgress("Finished Rendering", 100);
        }
        public void renderGerber()
        {
            GerberLibrary.GerberImageCreator GIC = new GerberLibrary.GerberImageCreator();
            GerberLibrary.BoardRenderColorSet Colors = new GerberLibrary.BoardRenderColorSet();
            Colors.BoardRenderColor = SolderMaskColor;
            Colors.BoardRenderSilkColor = SilkScreenColor;
            Colors.BoardRenderPadColor = CopperColor;
            Colors.BoardRenderTraceColor = TracesColor;
            SetProgress("Image generation started", -1);
            GIC.SetColors(Colors);

            GerberLibrary.Gerber.SaveIntermediateImages = false;


            bool fixgroup = true;
            string ext1 = Path.GetExtension(Files[0]);
            if (Files.Count == 1 && ext1 != ".zip") fixgroup = false;
            GIC.AddBoardsToSet(Files, fixgroup, this);

            if (GIC.Errors.Count > 0)
            {
                foreach (var a in GIC.Errors)
                {
                    Errors.Add(String.Format("Error: {0}", a));
                }
            }
            try
            {
                if (GIC.Count() > 1)
                {

                    if (Files.Count() == 1)
                    {
                        string justthefilename = Path.Combine(Path.GetDirectoryName(Files[0]), Path.GetFileNameWithoutExtension(Files[0]));
                        //GIC.WriteImageFiles(justthefilename, 200, false, this);
                        GIC.WriteImageToMemory(justthefilename,ref UtilityClass.TopDecal,ref UtilityClass.BottomDecal, 1000, this);
                    }
                }
            }
            catch(Exception E)
            {
                Errors.Add("Some errors:");
                while(E!=null)
                {
                    Errors.Add(E.Message);
                    E = E.InnerException;
                }
            }
            SetProgress("Done!",100);
            if (Errors.Count > 0)
            {
                SetProgress("Encountered some problems during image generation:", -1);
            }
            foreach (var a in Errors)
            {
                SetProgress(a, -1);
            }

        }
        List<String> Errors = new List<string>();
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (T == null) return;
            if (T.ThreadState == ThreadState.Stopped && Errors.Count == 0)
            {
                this.Close();

            }
        }

        public void AddString(string text, float progress = -1F)
        {
            SetProgress(text, progress);

        }

        delegate void UpdateBar(string text, float progress);

        private void SetProgress(string text, float progress)
        {

            if (this.progressBar1.InvokeRequired)
            {
                UpdateBar d = new UpdateBar(SetProgress);
                this.Invoke(d, new object[] { text, progress });
            }
            else
            {

                if (progress >= 100)
                {
                    progressBar1.Value = 100;
                    this.ControlBox = true;
                }
                else
                {
                    int temp;
                    if (progress == -1)
                    {
                        temp = progressBar1.Value;
                        temp += 2;
                        if (temp >= 100)
                        {
                            temp = 100;
                            this.ControlBox = true;
                        }
                        else
                            this.ControlBox = false;
                        progressBar1.Value = temp;
                    }
                    else
                    {
                        temp = (int)progress;
                        if (temp >= 100)
                        {
                            temp = 100;
                            this.ControlBox = true;
                        }
                        else
                            this.ControlBox = false;
                        progressBar1.Value = temp;
                    }
                }
                if(text.Trim().Length>0)
                    textBox1.Text = text + "\r\n" + textBox1.Text;
                textBox1.Refresh();
            }
        }


    }
}
