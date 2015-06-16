using System.Collections.Generic;
using ClosedXML.Excel;
using me.closedxml.Queries.QueryResult;

namespace me.closedxml
{
    public interface IExcelData<T> where T : IQueryResult
    {
        string WorksheetName { get; set; }
        IEnumerable<T> Data { get; set; }
        
        void Write(XLWorkbook workbook);
    }
}