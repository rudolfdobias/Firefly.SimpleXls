using System;
using System.Globalization;
using System.Reflection;

namespace Firefly.SimpleXls.Converters
{
    /// <inheritdoc />
    public class DateTimeXlsValueConverter : IValueConverter
    {
        /// <inheritdoc />
        public object Write(object item, Type itemType, CultureInfo culture = null)
        {
            return typeof(DateTime).IsAssignableFrom(itemType)
                ? ((DateTime) item).ToString(culture ?? CultureInfo.CurrentCulture)
                : item;
        }

        /// <inheritdoc />
        public object Read(object item)
        {
            if (DateTime.TryParse((string) item, out var val))
            {
                return val;
            }

            throw new ArgumentException("Cannot parse datetime.", nameof(item));
        }
    }
}