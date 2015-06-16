using System.Collections.Generic;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml
{
    public abstract class ExcelConfigurationWorksheetRow
    {
        protected ExcelConfigurationWorksheetRow(string worksheetName, string configurationTypeName, string typeName, string headerRange, string dataRange)
        {
            WorksheetName = worksheetName;
            ConfigurationTypeName = configurationTypeName;
            TypeName = typeName;
            HeaderRange = headerRange;
            DataRange = dataRange;
        }

        public string ConfigurationTypeName { get; set; }
        public string DataRange { get; set; }
        public string HeaderRange { get; set; }
        public string TypeName { get; set; }
        public string WorksheetName { get; set; }

        public abstract IEnumerable<IQueryResult> Read(IXLWorksheet worksheet);
    }
}