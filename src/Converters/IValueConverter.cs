using System;
using System.Globalization;

namespace Firefly.SimpleXls.Converters
{
    /// <summary>
    /// Implements conversion of values between xls and data models
    /// </summary>
    public interface IValueConverter
    {
        /// <summary>
        /// Writes cell value to XLS
        /// </summary>
        /// <param name="item"></param>
        /// <param name="itemType"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        object Write(object item, Type itemType, CultureInfo culture = null);

        /// <summary>
        /// Reads cell value from XLS
        /// </summary>
        /// <param name="item"></param>
        object Read(object item);
    }
}