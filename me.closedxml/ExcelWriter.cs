using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace closedxml
{
    public class ExcelWriter
    {
        private readonly IEnumerable<ExcelData> _items;
        readonly XLWorkbook _workbook;
        private IXLWorksheet _configurationWorkSheet;

        public ExcelWriter(IEnumerable<ExcelData> items)
        {
            _items = items;
            _workbook = new XLWorkbook();
        }

        public void Write()
        {
            AddConfigurationWorkSheet();
            foreach (var item in _items)
            {
                AddData(item);
            }
            AdjustToContentsConfigurationWorkSheet();
            _workbook.SaveAs(@"c:\temp\excel.xlsx");
        }

        private void AddData(ExcelData item)
        {
            var worksheet = _workbook.Worksheets.Add(item.Name);
            Type type = item.Data.First().GetType();
            PropertyInfo[] propertyInfos = type.GetProperties();
            for (int i = 0; i < propertyInfos.Count(); i++)
            {
                worksheet.Cell(1, i + 1).Value = propertyInfos[i].Name;
            }

            WriteInConfigurationWorkSheet(item.Name, type.FullName);

            worksheet.Cell("A2").Value = item.Data;

            worksheet.Range(1, 1, 1, propertyInfos.Count()).Style
                .Font.SetFontSize(11)
                .Font.SetBold(true)
                .Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(XLColor.Gray);
            
            worksheet.Columns().AdjustToContents();
            worksheet.Protect("1234");
        }

        private void AddConfigurationWorkSheet()
        {
            _configurationWorkSheet = _workbook.Worksheets.Add("Configuration");
            _configurationWorkSheet.Cell("A1").Value = "Worksheet Name";
            _configurationWorkSheet.Cell("B1").Value = "Type";
            _configurationWorkSheet.Range("A1:B1").Style
                .Font.SetFontSize(11)
                .Font.SetBold(true)
                .Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(XLColor.Gray);

            _configurationWorkSheet.Columns().AdjustToContents();
            _configurationWorkSheet.Protect("1234");
        }

        private void WriteInConfigurationWorkSheet(string workSheetName, string type)
        {
            var lastRowNumber = _configurationWorkSheet.LastCellUsed().Address.RowNumber;
            var currentRowNumber = lastRowNumber+1;

            _configurationWorkSheet.Cell(currentRowNumber, 1).Value = workSheetName;
            _configurationWorkSheet.Cell(currentRowNumber, 2).Value = type;
        }

        private void AdjustToContentsConfigurationWorkSheet()
        {
            _configurationWorkSheet.Columns().AdjustToContents();
        }
    }

    public class ExcelData
    {
        public string Name { get; set; }
        public IEnumerable<object> Data { get; set; }
    }

    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; }
    }

    public class Company
    {
        public int CompanyId { get; set; }
        public string Name { get; set; }
    }
}
