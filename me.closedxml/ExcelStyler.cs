using ClosedXML.Excel;

namespace me.closedxml
{
    public class ExcelStyler
    {
        private readonly XLColor _headerFontColor;
        private readonly XLColor _backgroundColor;
        private int FontSize { get; set; }

        public ExcelStyler(int fontSize = 11, XLColor headerFontColor = null, XLColor backgroundColor = null)
        {
            FontSize = fontSize;
            _headerFontColor = headerFontColor ?? XLColor.White;
            _backgroundColor = backgroundColor ?? XLColor.Black;
        }

        public void SetHeaderStyle(IXLRange range)
        {
            range.Style
                .Font.SetFontSize(FontSize)
                .Font.SetBold(true)
                .Font.SetFontColor(_headerFontColor)
                .Fill.SetBackgroundColor(_backgroundColor);
        }
    }
}