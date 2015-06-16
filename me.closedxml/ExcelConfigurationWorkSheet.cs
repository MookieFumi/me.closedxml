using System.Linq;
using ClosedXML.Excel;

namespace me.closedxml
{
    public class ExcelConfigurationWorksheet
    {
        public static IXLWorksheet WorkSheet;
        private readonly ExcelStyler _excelStyler;

        public ExcelConfigurationWorksheet(XLWorkbook workbook, ExcelStyler excelStyler)
        {
            if (workbook.Worksheets.All(p => p.Name != strings.Configuration))
            {
                workbook.Worksheets.Add(strings.Configuration);
            }
            WorkSheet = workbook.Worksheets.Single(p => p.Name == strings.Configuration);
            _excelStyler = excelStyler;
            AddHeaderValues();
            SetHeaderStyle();
        }

        private void AddHeaderValues()
        {
            WorkSheet.Cell("A1").Value = "WorksheetName";
            WorkSheet.Cell("B1").Value = "ConfigurationTypeName";
            WorkSheet.Cell("C1").Value = "TypeName";
            WorkSheet.Cell("D1").Value = "HeaderRange";
            WorkSheet.Cell("E1").Value = "DataRange";
        }

        private void SetHeaderStyle()
        {
            _excelStyler.SetHeaderStyle(WorkSheet.Range("A1:E1"));
        }

        public void Write(ExcelConfigurationWorksheetRow excelConfigurationWorksheetRow)
        {
            var lastRowNumber = WorkSheet.LastCellUsed().Address.RowNumber;
            var currentRowNumber = lastRowNumber + 1;

            WorkSheet.Cell(currentRowNumber, ExcelConfigurationWorksheetColumnNumber.WorkSheetName).Value = excelConfigurationWorksheetRow.WorksheetName;
            WorkSheet.Cell(currentRowNumber, ExcelConfigurationWorksheetColumnNumber.ConfigurationTypeName).Value = excelConfigurationWorksheetRow.ConfigurationTypeName;
            WorkSheet.Cell(currentRowNumber, ExcelConfigurationWorksheetColumnNumber.TypeName).Value = excelConfigurationWorksheetRow.TypeName;
            WorkSheet.Cell(currentRowNumber, ExcelConfigurationWorksheetColumnNumber.HeaderRange).Value = excelConfigurationWorksheetRow.HeaderRange;
            WorkSheet.Cell(currentRowNumber, ExcelConfigurationWorksheetColumnNumber.DataRange).Value = excelConfigurationWorksheetRow.DataRange;
        }
    }
}