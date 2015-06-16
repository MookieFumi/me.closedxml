using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Linq.Expressions;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml.Reader
{
    public class ExcelReader
    {
        private readonly XLWorkbook _workbook;

        public ExcelReader(string file)
        {
            _workbook = new XLWorkbook(file);
        }

        public IEnumerable<IExcelData<IQueryResult>> Read()
        {
            var returnValue = new Collection<IExcelData<IQueryResult>>();

            var configurationWorksheetRows = GetConfigurationWorksheetRows();
            foreach (var configurationWorksheetRow in configurationWorksheetRows)
            {
                var worksheet = _workbook.Worksheet(configurationWorksheetRow.WorksheetName);
                var data = configurationWorksheetRow.Read(worksheet);
                var excelData = new ExcelData<ExcelConfigurationWorksheetRow>(worksheet.Name, data);
                returnValue.Add(excelData);
            }

            return returnValue;
        }

        private IEnumerable<ExcelConfigurationWorksheetRow> GetConfigurationWorksheetRows()
        {
            var configurationWorksheet = _workbook.Worksheets.Single(p => p.Name == strings.Configuration);
            var lastRowUsed = configurationWorksheet.LastRowUsed();
            var lastColumnUsed = configurationWorksheet.LastColumnUsed();
            var range = configurationWorksheet.Range(2, 1, lastRowUsed.RowNumber(), lastColumnUsed.ColumnNumber());
            var rows = new Collection<ExcelConfigurationWorksheetRow>();
            foreach (var row in range.Rows())
            {
                var worksheetName = row.Cell(1).GetValue<string>();
                var configurationTypeName = row.Cell(2).GetValue<string>();
                var typeName = row.Cell(3).GetValue<string>();
                var headerRange = row.Cell(4).GetValue<string>();
                var dataRange = row.Cell(5).GetValue<string>();
                var type = Type.GetType(configurationTypeName);

                var dataConfiguration = (ExcelConfigurationWorksheetRow)Activator.CreateInstance(type, worksheetName, configurationTypeName, typeName, headerRange, dataRange);
                rows.Add(dataConfiguration);
            }
            return rows;
        }
    }
}