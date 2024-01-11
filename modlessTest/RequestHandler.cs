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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Security.Cryptography;

namespace modlessTest
{
    public class RequestHandler : IExternalEventHandler
    {
        public static UIDocument m_uidoc = null;
        public static UIApplication m_uiapp = null;
        public static Document m_doc = null;

        private Request m_request = new Request();
        public Request Request
        {
            get { return m_request; }
        }
        public String GetName()
        {
            return "testing";
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        {
                            CreateBOJ(uiapp);
                            break;
                        }
                    case RequestId.Test:
                        {
                            testexcution(uiapp);
                            break;
                        }
                    case RequestId.gangpin:
                        {
                            _gangpin(uiapp);
                            break;
                        }
                    case RequestId.readCAD:
                        {
                            _readCAD(uiapp);
                            break;
                        }
                    case RequestId.CreateBeams:
                        {
                            _CreateBeams(uiapp);
                            break;
                        }
                    case RequestId.deckslab:
                        {
                            _deckslab(uiapp);
                            break;
                        }
                    case RequestId.deckSlab2:
                        {
                            _deckSlab2(uiapp);
                            break;
                        }
                    case RequestId.gang:
                        {
                            _gang(uiapp);
                            break;
                        }
                }
            }
            finally
            {
                App.thisApp.WakeFormUp();
            }
            return;
        }

        private void _gang(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            List<string> filepath = Mylib.GetDWGLink(m_doc);

            CadDocument docCad = DwgReader.Read(filepath[0]);
            List<Entity> entities = new List<Entity>(docCad.Entities);

            FilteredElementCollector col = new FilteredElementCollector(m_doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilyInstance));

            Dictionary<FamilyInstance, XYZ> dixdix = new Dictionary<FamilyInstance, XYZ>();

            foreach (FamilyInstance item in col)
            {
                if(item.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    LocationCurve lc = item.Location as LocationCurve;
                    Curve c = lc.Curve;
                    dixdix.Add(item, c.Evaluate(0.5, true));
                }
            }

            List<compareData> cd = new List<compareData>();

            foreach (Entity item in entities)
            {
                ACadSharp.ObjectType obj = item.ObjectType;
                if (obj == ACadSharp.ObjectType.LINE)
                {
                    Autodesk.Revit.DB.Line line = Mylib.ConvertRevitLine(item as ACadSharp.Entities.Line);
                    if (line == null) continue;

                    if (line.Length * 304.8 < 801)
                    {
                        compareData cd1 = new compareData();
                        cd1.m_mid = line.Evaluate(0.5, true);
                        cd1.m_infor = "Pin";
                        cd.Add(cd1);
                    }
                }

                else if (obj == ACadSharp.ObjectType.LWPOLYLINE)
                {
                    ACadSharp.Entities.LwPolyline ee = item as ACadSharp.Entities.LwPolyline;
                    IEnumerable<Entity> eee = ee.Explode();
                    Entity en = eee.First();
                    Autodesk.Revit.DB.Line line = Mylib.ConvertRevitLine(en as ACadSharp.Entities.Line);
                    if (line == null) continue;

                    compareData cd1 = new compareData();
                    cd1.m_mid = line.Evaluate(0.5, true);
                    cd1.m_infor = "Moment";
                    cd.Add(cd1);
                }
            }

            using(Transaction trans = new Transaction(m_doc, "Change Moment"))
            {
                trans.Start();
                foreach (Entity item in entities)
                {
                    ACadSharp.ObjectType obj = item.ObjectType;
                    if (obj == ACadSharp.ObjectType.LINE)
                    {
                        Autodesk.Revit.DB.Line line = Mylib.ConvertRevitLine(item as ACadSharp.Entities.Line);
                        if (line == null) continue;

                        if (line.Length * 304.8 >= 801)
                        {
                            var sort = from n in dixdix orderby line.Evaluate(0.5, true).DistanceTo(n.Value) ascending select n;
                            FamilyInstance findinstance = sort.First().Key;

                            XYZ sp = line.GetEndPoint(0);
                            var sortsp = from n in cd orderby sp.DistanceTo(n.m_mid) ascending select n;
                            string infor = sortsp.First().m_infor;

                            XYZ ep = line.GetEndPoint(1);
                            var sortsp1 = from n in cd orderby ep.DistanceTo(n.m_mid) ascending select n;
                            string infor1 = sortsp1.First().m_infor;


                            Parameter param = findinstance.get_Parameter(BuiltInParameter.STRUCT_CONNECTION_BEAM_START);
                            Parameter param1 = findinstance.get_Parameter(BuiltInParameter.STRUCT_CONNECTION_BEAM_END);

                            StructuralConnectionType st = null;
                            FilteredElementCollector cococo = new FilteredElementCollector(m_doc).OfClass(typeof(StructuralConnectionType));

                            LocationCurve lc = findinstance.Location as LocationCurve;
                            Curve cc = lc.Curve;

                            XYZ fsp = cc.GetEndPoint(0);
                            XYZ fep = cc.GetEndPoint(1);

                            double t = sp.DistanceTo(fsp);
                            double t1 = sp.DistanceTo(fep);

                            if (t > t1)
                            {
                                foreach (StructuralConnectionType eee in cococo)
                                {
                                    if (eee.Name == infor1)
                                    {
                                        st = eee;
                                        break;
                                    }
                                }

                                string NewId = st.Id.ToString();
                                int idInt = Convert.ToInt32(NewId);
                                param1.Set(new ElementId(idInt));

                                foreach (StructuralConnectionType eee in cococo)
                                {
                                    if (eee.Name == infor)
                                    {
                                        st = eee;
                                        break;
                                    }
                                }

                                string NewId1 = st.Id.ToString();
                                int idInt1 = Convert.ToInt32(NewId1);
                                param.Set(new ElementId(idInt1));
                            }

                            else
                            {
                                foreach (StructuralConnectionType eee in cococo)
                                {
                                    if (eee.Name == infor)
                                    {
                                        st = eee;
                                        break;
                                    }
                                }

                                string NewId = st.Id.ToString();
                                int idInt = Convert.ToInt32(NewId);
                                param.Set(new ElementId(idInt));

                                foreach (StructuralConnectionType eee in cococo)
                                {
                                    if (eee.Name == infor1)
                                    {
                                        st = eee;
                                        break;
                                    }
                                }

                                string NewId1 = st.Id.ToString();
                                int idInt1 = Convert.ToInt32(NewId1);
                                param1.Set(new ElementId(idInt1));

                            }
                        }
                    }
                }
                trans.Commit();
            }
        }

        private void CreateBOJ(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
        }

        private void testexcution(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            IList<Reference> references = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element);
            List<FamilyInstance> cols = new List<FamilyInstance>();

            foreach (Reference item in references)
            {
                Element e = m_doc.GetElement(item.ElementId);
                if(e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
                {
                    cols.Add(e as FamilyInstance);
                }
            }


            Dictionary<FamilyInstance, double> keyValuePairs = new Dictionary<FamilyInstance, double>();

            foreach (FamilyInstance item in cols)
            {
                Parameter levelparamBase = item.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                Parameter levelparamtop = item.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);

                if (levelparamBase != null)
                {
                    ElementId idx = levelparamBase.AsElementId();
                    Level level = m_doc.GetElement(idx) as Level;
                    double el = level.Elevation;
                    keyValuePairs.Add(item, el);
                }
            }

            var sort = from n in keyValuePairs orderby n.Value ascending select n;

            FamilyInstance firstElement = sort.First().Key;
            FamilyInstance lastElement = sort.Last().Key;

            string baselevelname = levelname(firstElement, true);
            string Toplevelname = levelname(lastElement, false);

            LocationPoint lp = firstElement.Location as LocationPoint;
            XYZ p1 = lp.Point;
            string final = baselevelname + "-" + Toplevelname;

            createT(m_doc, p1, final);
        }

        public static string levelname(FamilyInstance f, bool b)
        {
            string levelname = "";
            if(b == true)
            {
                Parameter levelparamBase1 = f.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                ElementId idx1 = levelparamBase1.AsElementId();
                Level level1 = m_doc.GetElement(idx1) as Level;
                levelname = level1.Name;
            }

            else if(b == false) 
            {
                Parameter levelparamBase1 = f.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                ElementId idx1 = levelparamBase1.AsElementId();
                Level level1 = m_doc.GetElement(idx1) as Level;
                levelname = level1.Name;
            }

            return levelname;
        }

        public void createT(Document doc, XYZ p1, string text)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            TextNoteType textNoteType = collector.OfClass(typeof(TextNoteType)).FirstOrDefault() as TextNoteType;

            using(Transaction trans = new Transaction(doc, "Create Text"))
            {
                trans.Start();
                TextNote tn = TextNote.Create(doc, doc.ActiveView.Id, p1, text, textNoteType.Id);
                trans.Commit();
            }
        }

        public void CreateTextElement(Document doc, XYZ position, string text)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            TextNoteType textNoteType = collector.OfClass(typeof(TextNoteType)).FirstOrDefault() as TextNoteType;

            if (textNoteType == null)
            {
                TaskDialog.Show("Error", "No TextNoteType found.");
                return;
            }

            // TextNote 생성
            using (Transaction t = new Transaction(doc, "Create TextNote"))
            {
                t.Start();

                TextNote textNote = TextNote.Create(doc, doc.ActiveView.Id, position, text, textNoteType.Id);

                t.Commit();
            }
        }

        private void _gangpin(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            string filepath = Mylib.Getfilepath();
            string[] lines = File.ReadAllLines(filepath);

            Dictionary<FamilyInstance, XYZ> dix = new Dictionary<FamilyInstance, XYZ>();
            List<FamilyInstance> beams = Mylib.getinstance(m_doc);

            foreach (FamilyInstance b in beams)
            {
                LocationCurve lc = b.Location as LocationCurve;
                Curve curve = lc.Curve;
                XYZ mid = curve.Evaluate(0.5, true);

                dix.Add(b, mid);
            }

            if (lines.Length > 0)
            {
                foreach (string item in lines)
                {
                    string[] strs = item.Split(',');
                    double x = Mylib.converDouble(strs[0]);
                    double y = Mylib.converDouble(strs[1]);
                    double z = Mylib.converDouble(strs[2]);

                    XYZ midp = new XYZ(x, y, z);

                    var sort = from n in dix orderby midp.DistanceTo(n.Value) ascending select n;
                    FamilyInstance f = sort.Last().Key;

                    double x1 = Mylib.converDouble(strs[3]);
                    double y1 = Mylib.converDouble(strs[4]);
                    double z1 = Mylib.converDouble(strs[5]);

                    XYZ p1 = new XYZ(x1, y1, z1);
                    string h = strs[6];
                    string hh = strs.Last();

                    LocationCurve lc = f.Location as LocationCurve;
                    Curve curve = lc.Curve;
                    XYZ sp = curve.GetEndPoint(0);
                    XYZ ep = curve.GetEndPoint(1);

                    double t = p1.DistanceTo(sp);
                    double t1 = p1.DistanceTo(ep);

                    XYZ pp = new XYZ();

                    if(t < t1)
                    {
                        pp = sp;
                    }
                    else
                    {
                        pp = ep;
                    }

                    using(Transaction trans = new Transaction(m_doc, "fff"))
                    {
                        trans.Start();
                        Parameter param = f.get_Parameter(BuiltInParameter.STRUCT_CONNECTION_BEAM_START);
                        StructuralConnectionType st = null;
                        FilteredElementCollector cococo = new FilteredElementCollector(m_doc).OfClass(typeof(StructuralConnectionType));

                        foreach (StructuralConnectionType eee in cococo)
                        {
                            if(eee.Name == h)
                            {
                                st = eee;
                                break;
                            }
                        }

                        string NewId = st.Id.ToString();
                        int idInt = Convert.ToInt32(NewId);
                        param.Set(new ElementId(idInt));

                        Parameter param1 = f.get_Parameter(BuiltInParameter.STRUCT_CONNECTION_BEAM_END);
                        StructuralConnectionType st1 = null;
                        FilteredElementCollector cococo1 = new FilteredElementCollector(m_doc).OfClass(typeof(StructuralConnectionType));

                        foreach (StructuralConnectionType eee in cococo1)
                        {
                            if (eee.Name == hh)
                            {
                                st1 = eee;
                                break;
                            }
                        }

                        string NewId1 = st1.Id.ToString();
                        int idInt1 = Convert.ToInt32(NewId1);
                        param1.Set(new ElementId(idInt1));

                        trans.Commit();
                    }
                }
            }
        }

        private void _readCAD(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            List<string> filepath = Mylib.GetDWGLink(m_doc);

            CadDocument docCad = DwgReader.Read(filepath[0]);
            List<Entity> entities = new List<Entity>(docCad.Entities);

            FilteredElementCollector col = new FilteredElementCollector(m_doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilyInstance));


            Dictionary<XYZ, string> dix = new Dictionary<XYZ, string>();

            foreach (Entity item in entities)
            {
                ACadSharp.ObjectType obj = item.ObjectType;
                if (obj == ACadSharp.ObjectType.LINE)
                {
                    Autodesk.Revit.DB.Line line = Mylib.ConvertRevitLine(item as ACadSharp.Entities.Line);
                    if(line.Length * 304.8 < 801)
                    {
                        dix.Add(line.Evaluate(0.5, true), "PIN");
                    }
                }

                else if(obj == ACadSharp.ObjectType.LWPOLYLINE)
                {
                    ACadSharp.Entities.LwPolyline ee = item as ACadSharp.Entities.LwPolyline;
                    
                }
            }

            foreach (Entity item in entities)
            {
                ACadSharp.ObjectType obj = item.ObjectType;
                if (obj == ACadSharp.ObjectType.LINE)
                {
                    Autodesk.Revit.DB.Line line = Mylib.ConvertRevitLine(item as ACadSharp.Entities.Line);
                    XYZ sp = line.GetEndPoint(0);

                }
            }
        }

        private void _CreateBeams(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            List<string> filepath = Mylib.GetDWGLink(m_doc);
            Mylib.linedataText(m_doc, filepath[0], 800);

        }

        private void _deckslab(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            IList<Reference> refs = m_uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element);
            List<Curve> curves = new List<Curve>();

            foreach (Reference item in refs)
            {
                Element e = m_doc.GetElement(item.ElementId);
                if(e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    LocationCurve lc = (e as FamilyInstance).Location as LocationCurve;
                    Curve c = lc.Curve;
                    curves.Add(c);
                }
            }

            CurveArray excurves = new CurveArray();
            foreach (Curve item in curves)
            {
                Autodesk.Revit.DB.Line line = Mylib.GetExtendCurve(item, 2000 / 304.8);
                excurves.Append(line);
            }


            List<IList<CurveLoop>> curves1 = new List<IList<CurveLoop>>();

            using (TransactionGroup transGroup = new TransactionGroup(m_doc))
            {
                transGroup.Start();

                using (Transaction trans = new Transaction(m_doc, "create"))
                {
                    trans.Start();
                    //SketchPlane sp = SketchPlane.Create(m_doc, m_doc.ActiveView.GenLevel.Id);

                    ModelCurveArray mca = m_doc.Create.NewRoomBoundaryLines(m_doc.ActiveView.SketchPlane, excurves, m_doc.ActiveView);
                    trans.Commit();
                }

                Level level = m_doc.ActiveView.GenLevel;
                using (Transaction t = new Transaction(m_doc, "Topologie"))
                {
                    t.Start();
                    PlanTopology pt = m_doc.get_PlanTopology(level);
                    foreach (PlanCircuit pc in pt.Circuits)
                    {
                        if (!pc.IsRoomLocated)
                        {
                            Room r = m_doc.Create.NewRoom(null, pc);
                            CurveLoop cl = new CurveLoop();
                            IList<IList<BoundarySegment>> loops = r.GetBoundarySegments(new SpatialElementBoundaryOptions());
                            if (loops.Count == 0) continue;

                            IList<BoundarySegment> pp = loops.First();
                            foreach (BoundarySegment item in pp)
                            {
                                cl.Append(item.GetCurve());
                            }

                            IList<CurveLoop> lll = new List<CurveLoop>();
                            lll.Add(cl);
                            curves1.Add(lll);
                        }
                    }

                    t.Commit();
                }

                transGroup.RollBack();
            }

            using(Transaction t = new Transaction(m_doc, "Create slab"))
            {
                t.Start();
                foreach (IList<CurveLoop> item in curves1)
                {
                    Floor f = Floor.Create(m_doc, item, new FilteredElementCollector(m_doc).OfCategory(BuiltInCategory.OST_Floors).OfClass(typeof(FloorType)).FirstElementId(), m_doc.ActiveView.GenLevel.Id);
                }

                t.Commit();
            }
        }

        private void _deckSlab2(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;


            // 보를 선택한다.
            IList<Reference> refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element);

            CurveArray ca = new CurveArray();

            foreach (Reference item in refs)
            {
                Element e = m_doc.GetElement(item.ElementId);
                LocationCurve lc = e.Location as LocationCurve;
                Curve curve = lc.Curve;

                Autodesk.Revit.DB.Line getExLine = Mylib.GetExtendCurve(curve, 2000/304.8);
                ca.Append(getExLine);
            }

            List<IList<CurveLoop>> curves1 = new List<IList<CurveLoop>>();

            using (TransactionGroup transGroup = new TransactionGroup(m_doc))
            {
                transGroup.Start();

                using (Transaction trans = new Transaction(m_doc, "create"))
                {
                    trans.Start();
                    //SketchPlane sp = SketchPlane.Create(m_doc, m_doc.ActiveView.GenLevel.Id);

                    ModelCurveArray mca = m_doc.Create.NewRoomBoundaryLines(m_doc.ActiveView.SketchPlane, ca, m_doc.ActiveView);
                    trans.Commit();
                }

                Level level = m_doc.ActiveView.GenLevel;
                using (Transaction t = new Transaction(m_doc, "Topologie"))
                {
                    t.Start();
                    PlanTopology pt = m_doc.get_PlanTopology(level);
                    foreach (PlanCircuit pc in pt.Circuits)
                    {
                        if (!pc.IsRoomLocated)
                        {
                            Room r = m_doc.Create.NewRoom(null, pc);
                            CurveLoop cl = new CurveLoop();
                            IList<IList<BoundarySegment>> loops = r.GetBoundarySegments(new SpatialElementBoundaryOptions());
                            if (loops.Count == 0) continue;

                            IList<BoundarySegment> pp = loops.First();
                            foreach (BoundarySegment item in pp)
                            {
                                cl.Append(item.GetCurve());
                            }

                            IList<CurveLoop> lll = new List<CurveLoop>();
                            lll.Add(cl);
                            curves1.Add(lll);
                        }
                    }

                    t.Commit();
                }

                transGroup.RollBack();
            }

            using (Transaction t = new Transaction(m_doc, "Create slab"))
            {
                t.Start();
                foreach (IList<CurveLoop> item in curves1)
                {
                    Floor f = Floor.Create(m_doc, item, new FilteredElementCollector(m_doc).OfCategory(BuiltInCategory.OST_Floors).OfClass(typeof(FloorType)).FirstElementId(), m_doc.ActiveView.GenLevel.Id);
                }

                t.Commit();
            }
        }
    }

    public class linedata
    {
        public Autodesk.Revit.DB.Line m_cadLine { get; set; }
        public XYZ sp {  get; set; }
        public XYZ ep { get; set; }

        public string spInfor { get; set;}
        public string epInfor { get; set; }

    }

    public class compareData
    {
        public XYZ m_mid { get; set; }
        public string m_infor {  get; set; }
    }
}
