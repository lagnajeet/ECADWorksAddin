using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using GerberLibrary;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LP.SolidWorks.ECADWorksAddin
{
    static class JSONLibrary
    {
        private static int max_width = 8000;
        private static int max_height = 8000;
        public static double boardPhysicalHeight = 0;
        public static double boardPhysicaWidth = 0;
        
        public static void generateDecalImage(dynamic boardJSON, bool side, Bgra maskColor, Bgra traceColor, Bgra padColor, Bgra silkColor, ProgressLog Logger = null, int progressBarScale=100, int progressBarOffset=0)
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
            int featureProcessed = 0;
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
                lock("Mylock")
                {
                    featureProcessed++;
                    Logger.AddString("",progressBarOffset + ((featureProcessed * progressBarScale) / feature_count));
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
            Image<Bgra, byte> DecalImage = new Image<Bgra, byte>(r.Width, r.Height);
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
            overlayed_Image.Copy(DecalImage, My_Image_gray);
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
            //string directoryPath = Path.GetDirectoryName(boardPartPath);
            if (!side)
            {
                UtilityClass.BottomDecalImage = new Image<Bgra, byte>(DecalImage.Width, DecalImage.Height);
                DecalImage.CopyTo(UtilityClass.BottomDecalImage);
            }
            else
            {
                UtilityClass.TopDecalImage = new Image<Bgra, byte>(DecalImage.Width, DecalImage.Height);
                DecalImage.CopyTo(UtilityClass.TopDecalImage);
            }
            //{
            //    //output_Image = output_Image.Flip(FlipType.Horizontal);
            //    deleteFile(bottomDecal);
            //    bottomDecal = RandomString(10) + "bottom_decal.png";
            //    output_Image.PyrUp().Save(@Path.Combine(directoryPath, bottomDecal));
            //    bottomDecal = @Path.Combine(directoryPath, bottomDecal);
            //}
            //else
            //{
            //    output_Image = output_Image.Flip(FlipType.Vertical);
            //    deleteFile(topDecal);
            //    topDecal = RandomString(10) + "top_decal.png";
            //    output_Image.PyrUp().Save(@Path.Combine(directoryPath, topDecal));
            //    topDecal = @Path.Combine(directoryPath, topDecal);
            //}
            DecalImage.Dispose();
            DecalImage = null;
            GC.Collect();
        }

        private static Image<Bgra, Byte> Overlay(Image<Bgra, Byte> target, Image<Bgra, Byte> overlay)
        {
            Bitmap bmp = target.Bitmap;
            Graphics gra = Graphics.FromImage(bmp);
            gra.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
            gra.DrawImage(overlay.Bitmap, new Point(0, 0));

            return target;
        }
    }
}
