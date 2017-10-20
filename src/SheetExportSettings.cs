using System.Globalization;

namespace Firefly.SimpleXls
{
    public class SheetExportSettings
    {
        public string SheetName { get; set; }
        public CultureInfo UseCulture { get; set; } = CultureInfo.CurrentCulture;
        public bool OmitEmptyColumns { get; set; }
    }
}