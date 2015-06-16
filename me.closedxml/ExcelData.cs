using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml
{
    public class ExcelData<T> : IExcelData<IQueryResult> where T : ExcelConfigurationWorksheetRow
    {
        private const string FirstCellAddressInRange = "A2";
        private const int FirstCellRow = 2;
        private const int MaxCellRow = 999;
        private readonly ExcelStyler _excelStyler;

        public ExcelData(string worksheetName, IEnumerable<IQueryResult> data)
        {
            WorksheetName = worksheetName;
            ConfigurationTypeName = typeof(T).FullName;
            Data = data;
            _excelStyler = new ExcelStyler();
        }

        private string ConfigurationTypeName { get; set; }

        #region IExcelData<IQueryResult> members

        public string WorksheetName { get; set; }
        public IEnumerable<IQueryResult> Data { get; set; }

        public void Write(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add(WorksheetName);
            var type = Data.First().GetType();
            var propertyInfos = type.GetProperties();
            var configurationType = Type.GetType(ConfigurationTypeName);
            var dataRange = worksheet.Range(2, 1, Data.Count() + 1, propertyInfos.Count()).RangeAddress.ToString();
            var headerRange = worksheet.Range(1, 1, 1, propertyInfos.Count()).RangeAddress.ToString();

            AddDataInConfigurationWorkSheet(workbook, configurationType, WorksheetName, type, dataRange, headerRange);
            AddHeader(worksheet, propertyInfos);
            AddData(worksheet, propertyInfos);
        }

        #endregion

        private void AddData(IXLWorksheet worksheet, PropertyInfo[] propertyInfos)
        {
            worksheet.Cell(FirstCellAddressInRange).Value = Data;
            for (var i = 0; i < propertyInfos.Count(); i++)
            {
                var cellColumn = i + 1;
                var range = worksheet.Range(FirstCellRow, cellColumn, MaxCellRow, cellColumn);
                var dataValidation = range.SetDataValidation();
                switch (Type.GetTypeCode(propertyInfos[i].PropertyType))
                {
                    case TypeCode.DateTime:
                        range.DataType = XLCellValues.DateTime;
                        range.Style.DateFormat.SetFormat("dd-mm-yyyy");
                        dataValidation.AllowedValues = XLAllowedValues.Date;
                        dataValidation.ErrorMessage = "Only dates";
                        break;
                    case TypeCode.Int32:
                        range.DataType = XLCellValues.Number;
                        dataValidation.WholeNumber.Between(Int32.MinValue, Int32.MaxValue);
                        dataValidation.ErrorMessage = "Only integers";
                        break;
                    case TypeCode.Decimal:
                        range.DataType = XLCellValues.Number;
                        dataValidation.Decimal.Between(Int32.MinValue, Int32.MaxValue);
                        dataValidation.ErrorMessage = "Only numbers";
                        break;
                }
            }
        }

        private void AddDataInConfigurationWorkSheet(XLWorkbook workbook, Type configurationType, string worksheetName, Type type, string dataRange, string headerRange)
        {
            var dataConfiguration = (ExcelConfigurationWorksheetRow)Activator.CreateInstance(configurationType, worksheetName, configurationType.FullName, type.FullName, headerRange, dataRange);
            var excelConfigurationWorksheet = new ExcelConfigurationWorksheet(workbook, _excelStyler);
            excelConfigurationWorksheet.Write(dataConfiguration);
        }

        private void AddHeader(IXLWorksheet worksheet, PropertyInfo[] properties)
        {
            for (var i = 0; i < properties.Count(); i++)
            {
                worksheet.Cell(1, i + 1).Value = properties[i].Name;
            }

            _excelStyler.SetHeaderStyle(worksheet.Range(1, 1, 1, properties.Count()));
        }
    }
}
