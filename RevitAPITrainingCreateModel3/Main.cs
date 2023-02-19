using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Autodesk.Revit.ApplicationServices;

namespace RevitAPITrainingCreateModel3
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Wall> walls = CreateWalls(doc, 10000, 5000, GetLevel(doc, "Уровень 1"), GetLevel(doc, "Уровень 2"));
            AddDoor(doc, GetLevel(doc, "Уровень 1"), walls[0]);
            for (int i = 1; i < 4; i++)
            {
                AddWindow(doc, GetLevel(doc, "Уровень 1"), walls[i], 1000);
            }
            //AddRoof(doc, GetLevel(doc, "Уровень 2"), walls);
            AddExtrusionRoof(doc, GetLevel(doc, "Уровень 2"), walls);
            return Result.Succeeded;
        }

        private void AddDoor(Document doc, Level level, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            Transaction transaction = new Transaction(doc, "Добавление двери");
            transaction.Start();
            if (!doorType.IsActive) doorType.Activate();
            doc.Create.NewFamilyInstance(point, doorType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            transaction.Commit();
        }


        private void AddExtrusionRoof(Document doc, Level level, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Типовой - 400мм"))
                .Where(x => x.FamilyName.Equals("Базовая крыша"))
                .FirstOrDefault();
            
            double wallWidth = walls[0].Width;
            double wallLength = walls[0].get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
            double wallLengthSecond = walls[1].get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
            double levelHeight = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();

            Transaction transaction = new Transaction(doc, "Добавление крыши");
            transaction.Start();         

            CurveArray curveArray = new CurveArray();
            curveArray.Append(Line.CreateBound(new XYZ(0, - wallLengthSecond/2 - wallWidth / 2, levelHeight), new XYZ(0, 0, levelHeight + 10)));
            curveArray.Append(Line.CreateBound(new XYZ(0, 0, levelHeight + 10), new XYZ(0, wallLengthSecond / 2 + wallWidth / 2, levelHeight)));

            ReferencePlane plane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 10), new XYZ(0, 20, 0), doc.ActiveView);
                          
            doc.Create.NewExtrusionRoof(curveArray, plane, level, roofType, - wallLength/2 - wallWidth/2, wallLength/2 + wallWidth/2);

            transaction.Commit();
        }


        private void AddRoof(Document doc, Level level, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Типовой - 400мм"))
                .Where(x => x.FamilyName.Equals("Базовая крыша"))
                .FirstOrDefault();

            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;
            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dt, -dt, 0));
            points.Add(new XYZ(dt, -dt, 0));
            points.Add(new XYZ(dt, dt, 0));
            points.Add(new XYZ(-dt, dt, 0));
            points.Add(new XYZ(-dt, -dt, 0));

            Transaction transaction = new Transaction(doc, "Добавление крыши");
            transaction.Start();

            Application application = doc.Application;
            CurveArray footprint = application.Create.NewCurveArray();
            for (int i = 0; i < 4; i++)
            {
                LocationCurve curve = walls[i].Location as LocationCurve;
                XYZ p1 = curve.Curve.GetEndPoint(0);
                XYZ p2 = curve.Curve.GetEndPoint(1);
                Line line = Line.CreateBound(p1 + points[i], p2 + points[i + 1]);
                footprint.Append(line);
            }
            ModelCurveArray foorPrintToModelCurveMapping = new ModelCurveArray();
            FootPrintRoof footPrintRoof = doc.Create.NewFootPrintRoof(footprint, level, roofType, out foorPrintToModelCurveMapping);
            ModelCurveArrayIterator iterator = foorPrintToModelCurveMapping.ForwardIterator();
            iterator.Reset();
            while (iterator.MoveNext())
            {
                ModelCurve modelCurve = iterator.Current as ModelCurve;
                footPrintRoof.set_DefinesSlope(modelCurve, true);
                footPrintRoof.set_SlopeAngle(modelCurve, 0.5);
            }
            transaction.Commit();
        }


        private void AddWindow(Document doc, Level level, Wall wall, double height)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            Transaction transaction = new Transaction(doc, "Добавление окна");
            transaction.Start();
            if (!windowType.IsActive) windowType.Activate();
            FamilyInstance wind = doc.Create.NewFamilyInstance(point, windowType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            //высота нижнего бруса
            wind.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(UnitUtils.ConvertToInternalUnits(height, UnitTypeId.Millimeters));
            transaction.Commit();
        }


        public Level GetLevel(Document doc, string inputString)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
            .OfClass(typeof(Level))
            .OfType<Level>()
            .ToList();

            Level level = listLevel
               .Where(x => x.Name.Equals(inputString))
               .FirstOrDefault();
            return level;
        }
        public List<Wall> CreateWalls(Document doc, double widthInput, double depthInput, Level level1, Level level2)
        {

            double width = UnitUtils.ConvertToInternalUnits(widthInput, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(depthInput, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(doc, "Построение стен");
            transaction.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }
            transaction.Commit();

            return walls;
        }

    }
}