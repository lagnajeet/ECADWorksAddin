using Emgu.CV;
using Emgu.CV.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LP.SolidWorks.ECADWorksAddin
{
    static class UtilityClass
    {
        public static byte[] TopDecal=null;
        public static byte[] BottomDecal = null;
        public static Image<Bgra, byte> TopDecalImage;
        public static Image<Bgra, byte> BottomDecalImage;
    }
}
