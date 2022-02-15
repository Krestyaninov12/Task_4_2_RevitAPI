using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task_4_2
{//
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var pipes = new FilteredElementCollector(doc)
            .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pippes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    var pipeSymbolName = pipe.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    var outerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    var innerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    var lenght = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);

                    var pipeSymbolNameValue = pipeSymbolName.AsValueString();
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipeSymbolNameValue);

                    if (outerDiameter.StorageType == StorageType.Double && innerDiameter.StorageType == StorageType.Double && lenght.StorageType == StorageType.Double)
                    {
                        double outerDiameterValue = UnitUtils.ConvertFromInternalUnits(outerDiameter.AsDouble(), UnitTypeId.Millimeters);
                        double innerDiameterValue = UnitUtils.ConvertFromInternalUnits(innerDiameter.AsDouble(), UnitTypeId.Millimeters);
                        double lenghtValue = UnitUtils.ConvertFromInternalUnits(lenght.AsDouble(), UnitTypeId.Meters);

                        sheet.SetCellValue(rowIndex, columnIndex: 1, outerDiameterValue);
                        sheet.SetCellValue(rowIndex, columnIndex: 2, innerDiameterValue);
                        sheet.SetCellValue(rowIndex, columnIndex: 3, lenghtValue);
                    }
                    rowIndex++;
                }
                workbook.Write(stream);
                workbook.Close();
            }

            System.Diagnostics.Process.Start(excelPath);
            return Result.Succeeded;
        }
    }
}
