using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using Autodesk.Revit.DB.Architecture;
using System.IO;
using NPOI.SS.UserModel;
using Autodesk.Revit.DB.Plumbing;

namespace R_3_4_2
//{
//    [Transaction(TransactionMode.Manual)]
//    public class Main : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiapp = commandData.Application;
//            UIDocument uidoc = uiapp.ActiveUIDocument;
//            Document doc = uidoc.Document;

//            string roomInfo = string.Empty;

//            var rooms = new FilteredElementCollector(doc)
//                .OfCategory(BuiltInCategory.OST_Rooms)
//                .Cast<Room>()
//                .ToList();

//            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");
//            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
//            { 
//                IWorkbook workbook = new XSSFWorkbook();
//                ISheet sheet = workbook.CreateSheet("Лист1");

//                int rowIndex = 0;
//                foreach (var room in rooms)
//                {
//                    sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name);
//                }
//            }


//            TaskDialog.Show("Сообщение", "Тест");
//            return Result.Succeeded;
//        }
//    }
//}

{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string roomInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc, doc.ActiveView.Id)
            .OfClass(typeof(Pipe))
            .Cast<Pipe>()
            .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    if (pipe is Pipe)
                    {
                        Parameter nameParametr = pipe.get_Parameter(BuiltInParameter.COVER_TYPE_NAME);
                        Parameter rMaxParametr = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                        Parameter rMinParametr = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                        Parameter lengthParametr = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);

                        if (rMaxParametr.StorageType == StorageType.Double || rMinParametr.StorageType == StorageType.Double || lengthParametr.StorageType == StorageType.Double)
                        {
                            double rMaxhValue = UnitUtils.ConvertFromInternalUnits(rMaxParametr.AsDouble(), DisplayUnitType.DUT_METERS);
                            double rMax = rMaxhValue;

                            double rMinhValue = UnitUtils.ConvertFromInternalUnits(rMinParametr.AsDouble(), DisplayUnitType.DUT_METERS);
                            double rMin = rMinhValue;

                            double lengtValue = UnitUtils.ConvertFromInternalUnits(lengthParametr.AsDouble(), DisplayUnitType.DUT_METERS);
                            double lengt = lengtValue;
                            sheet.SetCellValue(rowIndex, columnIndex: 0, nameParametr);
                            sheet.SetCellValue(rowIndex, columnIndex: 1, lengt);
                            sheet.SetCellValue(rowIndex, columnIndex: 2, rMax);
                            sheet.SetCellValue(rowIndex, columnIndex: 3, rMin);
                        }
                    }
                }
                return Result.Succeeded;
            }
        }
    }
}
