using ClosedXML.Excel;

namespace me.closedxml
{
    public class ExcelConfigurationWorkSheet
    {
        private readonly IXLWorksheet _configurationWorkSheet;

        public ExcelConfigurationWorkSheet(XLWorkbook workbook)
        {
            _configurationWorkSheet = workbook.Worksheets.Add("Configuration");
            AddConfigurationWorkSheet();
        }

        private void AddConfigurationWorkSheet()
        {
            _configurationWorkSheet.Cell("A1").Value = "Worksheet Name";
            _configurationWorkSheet.Cell("B1").Value = "Type";
            _configurationWorkSheet.Range("A1:B1").Style
                .Font.SetFontSize(11)
                .Font.SetBold(true)
                .Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(XLColor.Gray);
        }

        public void WriteInConfigurationWorkSheet(string workSheetName, string type)
        {
            var lastRowNumber = _configurationWorkSheet.LastCellUsed().Address.RowNumber;
            var currentRowNumber = lastRowNumber + 1;

            _configurationWorkSheet.Cell(currentRowNumber, 1).Value = workSheetName;
            _configurationWorkSheet.Cell(currentRowNumber, 2).Value = type;
        }

        public void AdjustToContentsAndProtect()
        {
            _configurationWorkSheet.Columns().AdjustToContents();
            _configurationWorkSheet.Protect("1234");
        }
    }
}