using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace me.closedxml
{
    public class ExcelWriter
    {
        private readonly ExcelConfigurationWorkSheet _configuration;
        private readonly IEnumerable<ExcelData> _items;
        readonly XLWorkbook _workbook;

        public ExcelWriter(IEnumerable<ExcelData> items)
        {
            _items = items;
            _workbook = new XLWorkbook();
            _configuration = new ExcelConfigurationWorkSheet(_workbook);
        }

        public void Write()
        {
            foreach (var item in _items)
            {
                AddData(item);
            }
            SaveWorkBook();
        }

        private void AddData(ExcelData item)
        {
            var worksheet = _workbook.Worksheets.Add(item.Name);
            Type type = item.Data.First().GetType();
            PropertyInfo[] propertyInfos = type.GetProperties();
            for (var i = 0; i < propertyInfos.Count(); i++)
            {
                worksheet.Cell(1, i + 1).Value = propertyInfos[i].Name;
            }

            _configuration.WriteInConfigurationWorkSheet(item.Name, type.FullName);

            worksheet.Cell("A2").Value = item.Data;

            worksheet.Range(1, 1, 1, propertyInfos.Count()).Style
                .Font.SetFontSize(11)
                .Font.SetBold(true)
                .Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(XLColor.Gray);

            worksheet.Columns().AdjustToContents();
            worksheet.Protect("1234");
        }

        private void SaveWorkBook()
        {
            _configuration.AdjustToContentsAndProtect();
            _workbook.SaveAs(@"c:\temp\excel.xlsx");
        }
    }

    
}
