using System.Globalization;
using Microsoft.Extensions.Localization;

namespace Firefly.SimpleXls
{
    public class SheetExportSettings
    {
        public string SheetName { get; set; }
        public CultureInfo UseCulture { get; set; } = CultureInfo.CurrentCulture;
        public bool OmitEmptyColumns { get; set; }
        public IStringLocalizer Localizer { get; set; }
        public bool TranslateColumnHeaders { get; set; } = true;
        internal bool HasLocalizer => Localizer != null;

        internal IStringLocalizer GetLocalizer()
        {
            if (!HasLocalizer)
            {
                return null;
            }
            return Localizer.WithCulture(UseCulture);
        }
    }
}