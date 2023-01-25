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

namespace R_3_4_2
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

            var rooms = new FiltredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            { 
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var room in rooms)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name);
                }
            }


            TaskDialog.Show("Сообщение", "Тест");
            return Result.Succeeded;
        }
    }
}
