using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LP.SolidWorks.BlankAddin
{
    /// <summary>
    /// Our Solidworks taskpane add-in
    /// </summary>
    public class TaskpaneIntegration : ISwAddin
    {
        #region Private Members

        private int mSwCookie;
        private TaskpaneView mTaskpaneView;
        private TaskpaneHostUI mTaskpaneHost;
        public static SldWorks mSolidworksApplication;

        #endregion
        #region Public Members
        public const string SWTASKPANE_PROGID = "LP.SolidWorks.ECADWorksAddin.Taskpane";
        #endregion

        #region SOl;idworks Ad-in Callbacks

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            //throw new NotImplementedException();
            mSolidworksApplication = (SldWorks)ThisSW;
            mSwCookie = Cookie;
            var ok = mSolidworksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);
            LoadUI();
            // Return ok
            return true;
        }


        public bool DisconnectFromSW()
        {
            // Clean up our UI
            UnloadUI();

            // Return ok
            return true;
        }


        #endregion

        /// <summary>
        /// Create our Taskpane and inject our host UI
        /// </summary>
        private void LoadUI()
        {
            // Find location to our taskpane icon
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "logo-small.png");

            // Create our Taskpane
            mTaskpaneView = mSolidworksApplication.CreateTaskpaneView2(imagePath, "ECADWorks");

            // Load our UI into the taskpane
            mTaskpaneHost = (TaskpaneHostUI)mTaskpaneView.AddControl(TaskpaneIntegration.SWTASKPANE_PROGID, string.Empty);
        }



        /// <summary>
        /// Cleanup the taskpane when we disconnect/unload
        /// </summary>
        private void UnloadUI()
        {
            mTaskpaneHost = null;

            // Remove taskpane view
            mTaskpaneView.DeleteView();

            // Release COM reference and cleanup memory
            Marshal.ReleaseComObject(mTaskpaneView);

            mTaskpaneView = null;
        }
        #region COM Registration

        /// <summary>
        /// The COM registration call to add our registry entries to the SolidWorks add-in registry
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Create our registry folder for the add-in
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                // Load add-in when SolidWorks opens
                rk.SetValue(null, 1);

                // Set SolidWorks add-in title and description
                rk.SetValue("Title", "ECADWorks Addin");
                rk.SetValue("Description", "ECAD to Solidworks Converter");
            }
        }

        /// <summary>
        /// The COM unregister call to remove our custom entries we added in the COM register function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Remove our registry entry
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);

        }

        #endregion
    }
}
