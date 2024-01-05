using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using ACadSharp.IO;
using ACadSharp;
using ACadSharp.Attributes;
using ACadSharp.IO.Templates;
using ACadSharp.Tables;
using ACadSharp.Entities;
using ACadSharp.Tables.Collections;
using static Autodesk.Revit.DB.SpecTypeId;

namespace modlessTest
{
    public class Mylib
    {
        public static string Getfilepath()
        {
            string filename = "";
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filename = dialog.FileName;
            }
            return filename;
        }

        public static double converDouble(string s)
        {
            return Convert.ToDouble(s);
        }

        public static List<FamilyInstance> getinstance(Document doc)
        {
            List<FamilyInstance> getfa = new List<FamilyInstance>();

            FilteredElementCollector col = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilyInstance));

            foreach (Element item in col)
            {
                getfa.Add(item as FamilyInstance);
            }

            return getfa;
        }

        public static XYZ GetmidPoint(FamilyInstance c)
        {
            LocationCurve lc = c.Location as LocationCurve;
            Curve cc = lc.Curve;

            XYZ sp = cc.GetEndPoint(0);
            XYZ ep = cc.GetEndPoint(1);

            XYZ mid = (sp + ep) / 2;

            return mid;
        }

        public static List<string> GetDWGLink(Document doc)
        {
            List<string> filepath = new List<string>();

            FilteredElementCollector linkedDWG = new FilteredElementCollector(doc).OfClass(typeof(CADLinkType));
            try
            {
                foreach (CADLinkType linked in linkedDWG)
                {
                    if (linked.IsExternalFileReference() == true)
                    {
                        ExternalFileReference exFiRef = linked.GetExternalFileReference();
                        string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(exFiRef.GetAbsolutePath());
                        filepath.Add(path);
                    }
                }
            }

            catch (Exception ex)
            {
                TaskDialog.Show("경고", ex.Message);
            }

            return filepath;
        }

        public static List<string> ReadDWGLayers(string filepath)
        {
            List<string> ReadDWGLayers = new List<string>();
            try
            {
                if (filepath == "") return null;
                CadDocument docCad = DwgReader.Read(filepath);
                LayersTable lts = docCad.Layers;
                foreach (Layer item in lts)
                {
                    string layername = item.Name;
                    ReadDWGLayers.Add(layername);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("경고", ex.Message);
            }

            return ReadDWGLayers;
        }

        public static Dictionary<string, List<object>> GetTextInfor(string filepath)
        {
            CadDocument docCad = DwgReader.Read(filepath);
            List<Entity> entities = new List<Entity>(docCad.Entities);

            Dictionary<string, List<object>> dix = new Dictionary<string, List<object>>();

            try
            {
                List<object> strings = new List<object>();
                List<object> pts = new List<object>();
                List<object> layername = new List<object>();
                List<object> Colorname = new List<object>();

                foreach (Entity item in entities)
                {
                    ACadSharp.ObjectType obj = item.ObjectType;
                    if (obj == ACadSharp.ObjectType.TEXT)
                    {
                        TextEntity tn = item as TextEntity;
                        if (tn != null)
                        {
                            string textname = tn.Value;
                            XYZ revitPT = new XYZ(tn.InsertPoint.X, tn.InsertPoint.Y, tn.InsertPoint.Z);
                            strings.Add(textname);
                            pts.Add(revitPT);
                            layername.Add(tn.Layer);
                            Colorname.Add(tn.Color);
                        }
                    }

                    else if (obj == ACadSharp.ObjectType.MTEXT)
                    {
                        MText tn = item as MText;
                        if (tn != null)
                        {
                            string textname = tn.Value;
                            XYZ revitPT = new XYZ(tn.InsertPoint.X, tn.InsertPoint.Y, tn.InsertPoint.Z);
                            strings.Add(textname);
                            pts.Add(revitPT);
                            layername.Add(tn.Layer);
                            Colorname.Add(tn.Color);
                        }
                    }
                }

                dix.Add("Text", strings);
                dix.Add("XYZ", pts);
                dix.Add("Layer", layername);
                dix.Add("Color", Colorname);
            }

            catch (Exception ex)
            {
                TaskDialog.Show("경고", ex.Message);
            }

            return dix;
        }


        public static List<string> linedata(string filepath, double limits)
        {
            limits = 801;

            List<string> returnstring = new List<string>();
            CadDocument docCad = DwgReader.Read(filepath);
            List<Entity> entities = new List<Entity>(docCad.Entities);


            List<Autodesk.Revit.DB.Line> LongCurve = new List<Autodesk.Revit.DB.Line>();
            List<Autodesk.Revit.DB.Line> shortCurve = new List<Autodesk.Revit.DB.Line>();

            List<Polyline> PolylineCurve = new List<Polyline>();
            

            foreach (Entity item in entities)
            {
                ACadSharp.ObjectType obj = item.ObjectType;
                if (obj == ACadSharp.ObjectType.TEXT)
                {
                    continue;
                }
                else if(obj == ACadSharp.ObjectType.LINE)
                {
                    ACadSharp.Entities.Line line = item as ACadSharp.Entities.Line;
                    Autodesk.Revit.DB.Line getLine = ConvertRevitLine(line);

                    if (getLine.Length * 304.8 > limits)
                    {
                        LongCurve.Add(getLine);
                    }

                    else if ((getLine.Length * 304.8 <= limits))
                    {
                        shortCurve.Add(getLine);
                    }
                }             

                else if(obj == ACadSharp.ObjectType.LWPOLYLINE)
                {
                    PolylineCurve.Add(item as ACadSharp.Entities.Polyline);
                }
            }

            return returnstring;
        }


        public static void linedataText(Document doc, string filepath, double limits)
        {
            limits = 801;

            List<string> returnstring = new List<string>();
            CadDocument docCad = DwgReader.Read(filepath);
            List<Entity> entities = new List<Entity>(docCad.Entities);

            List<Autodesk.Revit.DB.Line> LongCurve = new List<Autodesk.Revit.DB.Line>();
            Dictionary<XYZ, string> keyValuePairs = new Dictionary<XYZ, string>();


            foreach (Entity item in entities)
            {
                ACadSharp.ObjectType obj = item.ObjectType;
                if (obj == ACadSharp.ObjectType.TEXT)
                {
                    TextEntity tn = item as TextEntity;
                    if (tn != null)
                    {
                        string textname = tn.Value;
                        XYZ revitPT = new XYZ(tn.InsertPoint.X, tn.InsertPoint.Y, tn.InsertPoint.Z)/304.8;
                        keyValuePairs.Add(revitPT, textname);
                    }
                }

                else if (obj == ACadSharp.ObjectType.LINE)
                {
                    ACadSharp.Entities.Line line = item as ACadSharp.Entities.Line;
                    Autodesk.Revit.DB.Line getLine = ConvertRevitLine(line);

                    if(getLine == null)
                    {
                        continue;
                    }

                    if (getLine.Length * 304.8 > limits)
                    {
                        LongCurve.Add(getLine);
                    }
                }
            }

            foreach (Autodesk.Revit.DB.Line item in LongCurve)
            {
                XYZ mid = item.Evaluate(0.5, true);
                var dd = from n in keyValuePairs orderby n.Key.DistanceTo(mid) ascending select n;
                string symbolname = dd.First().Value;

                FamilySymbol fs = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).FirstOrDefault(q => q.Name == symbolname) as FamilySymbol;
                if (fs == null) continue;

                using(Transaction trans = new Transaction(doc, "STrst"))
                {
                    trans.Start();
                    if (fs != null)
                    {
                        fs.Activate();
                        FamilyInstance fss = doc.Create.NewFamilyInstance(item, fs, doc.ActiveView.GenLevel, StructuralType.Beam);
                    }
                    trans.Commit();
                }
            }
        }


        public static double getlenth(ACadSharp.Entities.Line line)
        {
            CSMath.XYZ p1 = line.StartPoint;
            CSMath.XYZ p2 = line.EndPoint;

            XYZ pp1 = new XYZ(p1.X, p1.Y, p1.Z);
            XYZ pp2 = new XYZ(p2.X, p2.Y, p2.Z);

            return pp1.DistanceTo(pp2);
        }

        public static Autodesk.Revit.DB.Line ConvertRevitLine(ACadSharp.Entities.Line line)
        {
            Autodesk.Revit.DB.Line line1 = null;

            try
            {
                CSMath.XYZ p1 = line.StartPoint;
                CSMath.XYZ p2 = line.EndPoint;

                XYZ pp1 = new XYZ(p1.X, p1.Y, p1.Z) / 304.8;
                XYZ pp2 = new XYZ(p2.X, p2.Y, p2.Z) / 304.8;

                line1 = Autodesk.Revit.DB.Line.CreateBound(pp1, pp2);
            }
            catch (Exception e) 
            {
                MessageBox.Show(e.Message);
            }

            return line1;
        }

        public static Autodesk.Revit.DB.Line GetExtendCurve(Curve c, double length)
        {
            XYZ vector = (c.GetEndPoint(1) - c.GetEndPoint(0)).Normalize();
            XYZ pp1 = c.GetEndPoint(0);
            XYZ pp2 = c.GetEndPoint(1);

            XYZ exp1 = new XYZ(pp1.X + (length * -vector.X), pp1.Y + (length * -vector.Y), pp1.Z);
            XYZ exp2 = new XYZ(pp2.X + (length * vector.X), pp2.Y + (length * vector.Y), pp2.Z);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(exp1, exp2);
            return line;
        }
    }


    public class linedata
    {
        
    }
}
