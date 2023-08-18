using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace PipeSleeve
{

    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View actiView = doc.ActiveView;

            //fileter pipes
            FilteredElementCollector collectorPipe = new FilteredElementCollector(doc, actiView.Id);
            var listPipe = collectorPipe.OfClass(typeof(Pipe))
                .Cast<Pipe>().ToList();

            //filter rvt File
            //FilteredElementCollector collectorRvt = new FilteredElementCollector(doc);
            //var listRvt = collectorRvt.OfCategory(BuiltInCategory.OST_RvtLinks).
            //    Where(x => x.Name.Contains("Structure Model")).ToList();
            
            //pick file link
            Reference r = uidoc.Selection.PickObject(ObjectType.Element);
            Element ele = doc.GetElement(r);
            RevitLinkInstance instancelink = ele as RevitLinkInstance;
            Document linkdoc = instancelink.GetLinkDocument();
            FilteredElementCollector collectorLink = new FilteredElementCollector(linkdoc);
            IList<Element> ListLinkedEle = collectorLink.OfClass(typeof(Floor)).ToElements();

            //setup all list
            List<Floor> listFloorFileLink = new List<Floor>();
            List<Solid> listSolitPipeCheck = new List<Solid>();
            List<Solid> listSolitFloorCheck = new List<Solid>();
            
            List<double> listDPipe = new List<double>();
            List<Element> listPipeCopy = new List<Element>();

            //setup count var
            int k1 = 0;
            Dictionary<double, double> valueDiamerter = new Dictionary<double, double>();
            //add DH Pipe
            valueDiamerter.Add(15.0, 15.0);
            valueDiamerter.Add(18.0, 15.0);
            valueDiamerter.Add(22.0, 20.0);
            valueDiamerter.Add(28.0, 20.0);
            valueDiamerter.Add(32.0, 25.0);
            valueDiamerter.Add(38.0, 32.0);
            valueDiamerter.Add(42.0, 40.0);
            valueDiamerter.Add(48.0, 40.0);
            valueDiamerter.Add(60.0, 50.0);
            valueDiamerter.Add(76.0, 65.0);
            valueDiamerter.Add(89.0, 80.0);
            valueDiamerter.Add(114.0, 100.0);
            valueDiamerter.Add(140.0, 125.0);
            valueDiamerter.Add(165.0, 150.0);
            valueDiamerter.Add(216.0, 200.0);
            valueDiamerter.Add(267.0, 250.0);
            valueDiamerter.Add(318.0, 300.0);

            //add PB Pipe
            valueDiamerter.Add(10.1, 15.0);
            valueDiamerter.Add(16.0, 15.0);
            valueDiamerter.Add(16.3, 15.0);
            valueDiamerter.Add(20.3, 20.0);
            valueDiamerter.Add(22.3, 25.0);
            valueDiamerter.Add(25.3, 25.0);
            valueDiamerter.Add(28.1, 32.0);
            valueDiamerter.Add(32.2, 32.0);
            valueDiamerter.Add(35.2, 35.0);
            
            //add SPP Pipe
            valueDiamerter.Add(89.1, 80.0);
            valueDiamerter.Add(114.3, 100.0);
            valueDiamerter.Add(139.8, 125.0);
            valueDiamerter.Add(165.2, 150.0);
            valueDiamerter.Add(216.3, 200.0);
            valueDiamerter.Add(267.4, 250.0);
            valueDiamerter.Add(318.5, 300.0);
            valueDiamerter.Add(355.5, 350.0);
            valueDiamerter.Add(406.4, 400.0);
            valueDiamerter.Add(457.2, 450.0);
            valueDiamerter.Add(508.0, 500.0);

            //add SUS Pipe
            valueDiamerter.Add(9.52, 15.0);
            valueDiamerter.Add(12.7, 15.0);
            valueDiamerter.Add(15.9, 15.0);
            valueDiamerter.Add(22.2, 20.0);
            valueDiamerter.Add(28.6, 25.0);
            valueDiamerter.Add(34.0, 32.0);
            valueDiamerter.Add(42.7, 40.0);
            valueDiamerter.Add(48.6, 50.0);
            valueDiamerter.Add(60.5, 65.0);
            valueDiamerter.Add(76.3, 80.0);

            using(Transaction trans = new Transaction(doc,"Get Count Pipe"))
            {
                trans.Start();
                foreach (Element item in ListLinkedEle)
                {
                    Floor floorItem = item as Floor;
                    listFloorFileLink.Add(floorItem);
                }
                
                foreach (Pipe item in listPipe)
                {
                    Element eleItem = item as Element;
                    List<Solid> listSolidPipe = GetSolidTest(eleItem);
                    
                    foreach (Solid solidItem in listSolidPipe)
                    {
                        listSolitPipeCheck.Add(solidItem);
                        
                        Parameter para = eleItem.LookupParameter("Outside Diameter");
                        listDPipe.Add(para.AsDouble()*304.8);
                        listPipeCopy.Add(eleItem);
                    }

                }

                //get bounding box solid Pipes
                foreach (Solid solid in listSolitPipeCheck)
                {
                    //bouding box intersec element
                    BoundingBoxXYZ boundingbox = solid.GetBoundingBox();
                    Transform t = boundingbox.Transform;
                    Outline outLine = new Outline(t.OfPoint(boundingbox.Min), t.OfPoint(boundingbox.Max));
                    
                    BoundingBoxIntersectsFilter boundingBoxFilter = new BoundingBoxIntersectsFilter(outLine);
                    FilteredElementCollector collector = new FilteredElementCollector(linkdoc);
                    var listIntersecEle = collector.WherePasses(boundingBoxFilter).ToList();

                    //filter Floor Intersec
                    List<Element> floorIntersec = new List<Element>();
                    if(listIntersecEle !=null)
                    foreach (Element item in listIntersecEle)
                        {
                            foreach (Floor itemFloor in listFloorFileLink)
                            {
                                if (item.Id == itemFloor.Id)
                                {
                                    floorIntersec.Add(item);
                                }
                            }
                        }

                    //get solid Floor 
                    foreach(Element item in floorIntersec)
                    {
                        List<Solid> listSolidFloorLinkFile = GetSolidTest(item);
                        foreach (Solid s in listSolidFloorLinkFile)
                        {
                            Solid interSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid, s, BooleanOperationsType.Intersect);
                            if (interSolid.Volume> 0)
                            {
                                //place sleeve
                                XYZ pointSleeve = interSolid.ComputeCentroid();
                                FamilyInstance instanceSleeve =  PlaceInstance(doc, pointSleeve);

                                //set L sleeve
                                BoundingBoxXYZ box = interSolid.GetBoundingBox();
                                var cuboidHeight = box.Max.Z - box.Min.Z;
                                var cuboidWidth = box.Max.Y - box.Min.Y;
                                Parameter paraSleeveL = instanceSleeve.LookupParameter("슬리브_L");
                                paraSleeveL.Set(cuboidHeight);
                                
                                //set D sleeve
                                Parameter paraSleeveD = instanceSleeve.LookupParameter("D");
                                double tg = listDPipe[k1];
                                double dSleeve = valueDiamerter[tg];
                                paraSleeveD.Set(dSleeve/304.8);

                                //get set others parameter 
                                GetSetParam("공종코드", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("공구코드", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("동코드", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("중분류", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("소분류", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("층코드", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("세대", listPipeCopy[k1], instanceSleeve);
                                GetSetParam("추가자재", listPipeCopy[k1], instanceSleeve);
                            }
                        }
                    }

                    k1++;

                }
                
                trans.Commit();
            }

            return Result.Succeeded;
        }

        public List<Solid> GetSolidEle(Element ele)
        {
            List<Solid> listSolid = new List<Solid>();
            Options opt = new Options();
            GeometryElement geoEle = ele.get_Geometry(opt);
            foreach( GeometryObject geomObj in geoEle)
            {
                GeometryInstance instance = geomObj as GeometryInstance;
                if(null != instance)
                {
                    foreach(GeometryObject instObj in instance.SymbolGeometry)
                    {
                        Solid solid = instObj as Solid;
                        listSolid.Add(solid);
                        if (null != solid || 0 != solid.Faces.Size || 0 != solid.Edges.Size)
                        {
                            //listSolid.Add(solid);
                        }
                        
                    }
                }

            }
            return listSolid;


        }

        public List<Solid> GetSolidTest(Element ele)
        {
            List<Solid> listSolid = new List<Solid>();
            Options opt = new Options();
            GeometryElement geoEle = ele.get_Geometry(opt);
            foreach (GeometryObject geomObj in geoEle)
            {
                Solid solid = geomObj as Solid;
                if(solid != null)
                {
                    listSolid.Add(solid);
                }
            }
            return listSolid;

        }

        public FamilyInstance PlaceInstance(Document doc, XYZ point)
        {
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> lst =  collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_PipeAccessory)
                .Where(x => x.Name == "강관 내화충진(벽체)").ToList();
                
            FamilySymbol symbol = lst.First() as FamilySymbol;
            View activeView = doc.ActiveView;
            FamilyInstance instance = doc.Create.NewFamilyInstance(point, symbol,Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            return instance;

        }

        public void GetSetParam(string paraName, Element eleSource, Element eleTo)
        {
            
            Parameter paraSouce = eleSource.LookupParameter(paraName);
            Parameter paraTo = eleTo.LookupParameter(paraName);
            paraTo.Set(paraSouce.AsString());
        }
        

    }
}
