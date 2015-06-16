using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml.Writer
{
    public class ExcelWriter
    {
        private readonly string _filePath;
        private readonly IEnumerable<IExcelData<IQueryResult>> _items;
        private readonly XLWorkbook _workbook;

        public ExcelWriter(string filePath, IEnumerable<IExcelData<IQueryResult>> items)
        {
            _workbook = new XLWorkbook();
            _filePath = filePath;
            _items = items;
        }

        public void Write()
        {
            foreach (var item in _items)
            {
                item.Write(_workbook);
            }
            
            SaveWorkBook();
        }

        private void SaveWorkBook()
        {
            var configurationWorksheet = ExcelConfigurationWorksheet.WorkSheet;
            var lastRowUsed = configurationWorksheet.LastRowUsed();
            var lastColumnUsed = configurationWorksheet.LastColumnUsed();
            var range = configurationWorksheet.Range(2, 1, lastRowUsed.RowNumber(), lastColumnUsed.ColumnNumber());

            foreach (var row in range.RowsUsed())
            {
                var workSheetName = row.Cell(ExcelConfigurationWorksheetColumnNumber.WorkSheetName).GetValue<string>();
                var dataRange = row.Cell(ExcelConfigurationWorksheetColumnNumber.DataRange).GetValue<string>();
                var worksheet = _workbook.Worksheets.Single(p => p.Name == workSheetName);

                worksheet.Columns().AdjustToContents();
                worksheet.Protect().FormatColumns = true;
                worksheet.Range(dataRange).Style.Protection.SetLocked(false);
            }

            configurationWorksheet.Columns().AdjustToContents();
            configurationWorksheet.Protect().FormatColumns = true;
            _workbook.Worksheets.Single(p=>p.Name==strings.Configuration).Position = 1;
            _workbook.SaveAs(_filePath);
        }
    }
}
