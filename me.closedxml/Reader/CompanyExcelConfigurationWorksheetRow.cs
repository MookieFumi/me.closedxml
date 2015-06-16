using System.Collections.Generic;
using System.Collections.ObjectModel;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml.Reader
{
    public class CompanyExcelConfigurationWorksheetRow : ExcelConfigurationWorksheetRow
    {
        public CompanyExcelConfigurationWorksheetRow(string worksheetName, string configurationTypeName, string typeName, string headerRange, string dataRange)
            : base(worksheetName, configurationTypeName, typeName, headerRange, dataRange)
        {
        }

        public override IEnumerable<IQueryResult> Read(IXLWorksheet worksheet)
        {
            var items = new Collection<CompanyQueryResult>();
            for (var i = 2; i <= worksheet.LastRowUsed().RowNumber(); i++)
            {
                var companyId = worksheet.Row(i).Cell(1).GetValue<int>();
                var name = worksheet.Row(i).Cell(2).GetValue<string>();
                var companyQueryResult = new CompanyQueryResult(companyId, name);
                items.Add(companyQueryResult);
            }
            return items;
        }
    }
}