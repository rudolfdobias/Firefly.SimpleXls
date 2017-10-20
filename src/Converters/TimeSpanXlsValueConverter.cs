using System;
using System.Globalization;
using System.Reflection;

namespace Firefly.SimpleXls.Converters
{
    /// <inheritdoc />
    public class TimeSpanXlsValueConverter : IValueConverter
    {
        /// <inheritdoc />
        public object Write(object item, Type itemType, CultureInfo culture = null)
        {
            return typeof(TimeSpan).IsAssignableFrom(itemType) ? ((TimeSpan) item).ToString() : item;
        }

        /// <inheritdoc />
        public object Read(object item)
        {
            if (TimeSpan.TryParse((string) item, out var val))
            {
                return val;
            }

            throw new ArgumentException("Cannot parse timespan.", nameof(item));
        }
    }
}