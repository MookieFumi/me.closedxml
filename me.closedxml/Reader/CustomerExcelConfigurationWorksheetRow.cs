using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml.Reader
{
    public class CustomerExcelConfigurationWorksheetRow : ExcelConfigurationWorksheetRow
    {
        public CustomerExcelConfigurationWorksheetRow(string worksheetName, string configurationTypeName, string typeName, string headerRange, string dataRange)
            : base(worksheetName, configurationTypeName, typeName, headerRange, dataRange)
        {
        }

        public override IEnumerable<IQueryResult> Read(IXLWorksheet worksheet)
        {
            var items = new Collection<CustomerQueryResult>();
            for (var i = 2; i <= worksheet.LastRowUsed().RowNumber(); i++)
            {
                var customerId = worksheet.Row(i).Cell(1).GetValue<int>();
                var name = worksheet.Row(i).Cell(2).GetValue<string>();
                var birthDate = worksheet.Row(i).Cell(3).GetValue<DateTime>();
                var lastInvoice = worksheet.Row(i).Cell(4).GetValue<decimal>();
                var removed = worksheet.Row(i).Cell(5).GetValue<bool>();
                var companyQueryResult = new CustomerQueryResult(customerId, name, birthDate, lastInvoice, removed);
                items.Add(companyQueryResult);
            }
            return items;
        }
    }
}