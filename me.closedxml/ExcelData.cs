using System.Collections.Generic;

namespace me.closedxml
{
    public class ExcelData
    {
        public string Name { get; set; }
        public IEnumerable<object> Data { get; set; }
    }
}
