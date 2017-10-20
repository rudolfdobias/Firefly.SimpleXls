using System;
using System.Collections.Generic;

namespace Firefly.SimpleXls.Converters
{
    public static class XlsConverters
    {
        internal static Dictionary<Type, IValueConverter> Converters { get; }

        static XlsConverters()
        {
            Converters = new Dictionary<Type, IValueConverter>
            {
                {typeof(DateTime), new DateTimeXlsValueConverter()},
                {typeof(TimeSpan), new TimeSpanXlsValueConverter()}
            };
        }

        /// <summary>
        /// Declares a custom type converter
        /// </summary>
        /// <param name="type"></param>
        /// <param name="converter"></param>
        public static void UseConveter(Type type, IValueConverter converter)
        {
            Converters[type] = converter;
        }
    }
}