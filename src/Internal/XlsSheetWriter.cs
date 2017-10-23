using System.Collections.Generic;
using System.Linq;
using Firefly.SimpleXls.Exceptions;
using OfficeOpenXml;

namespace Firefly.SimpleXls.Internal
{
    /// <summary>
    /// Writes data to excel sheet
    /// </summary>
    internal static class XlsSheetWriter
    {
        /// <summary>
        /// Creates new sheet in workbook and pours data in it, then saves into the stream.
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="items"></param>
        /// <param name="settings"></param>
        /// <typeparam name="T"></typeparam>
        /// <exception cref="SimpleXlsException"></exception>
        public static void WriteSheet<T>(ExcelPackage excel, IEnumerable<T> items, SheetExportSettings settings)
            where T : class, new()
        {
            var descriptor = ModelDescriber.DescribeModel<T>();

            var worksheet = CreateSheet(excel, descriptor, settings);
            CreateWorksheetHeader(worksheet, descriptor, settings);

            var columnUsages = new int[descriptor.Columns.Count + 1];
            columnUsages.Initialize();

            var row = 2; // 1 = table header, 2 = first row. Excel ...
            foreach (var i in items)
            {
                var col = 1;
                foreach (var d in descriptor.Columns)
                {
                    if (d.Attributes.Ignore)
                    {
                        continue;
                    }

                    object value;
                    if (d.CustomValueConverter != null)
                    {
                        value =
                            d.CustomValueConverter.Write(d.Property.GetValue(i),
                                d.Property.PropertyType, settings.UseCulture);
                    }
                    else
                    {
                        value = d.Property.GetValue(i);
                    }

                    if (value != null)
                    {
                        columnUsages[col]++;
                    }

                    if (d.Attributes.TranslateValue)
                    {
                        value = TranslateValue(value, d, settings);
                    }

                    worksheet.Cells[row, col].Value = value;
                    col++;
                }
                row++;
            }


            if (settings.OmitEmptyColumns)
            {
                OmitUnusedColumns(worksheet, descriptor.Columns.Count, columnUsages);
            }

            excel.Save();
        }

        /// <summary>
        /// Creates (and translates name of) a new sheet
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="descriptor"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        private static ExcelWorksheet CreateSheet(ExcelPackage excel, SheetDescriptor descriptor,
            SheetExportSettings settings)
        {
            // name from settings has priority
            var name = string.IsNullOrEmpty(settings.SheetName) ? descriptor.Name : settings.SheetName;

            // if nothing set, model type name will be used
            if (string.IsNullOrEmpty(name))
            {
                name = descriptor.ModelType.Name;
            }

            // name will be translated if localizer is present
            if (settings.HasLocalizer && (settings.Translate))
            {
                name = settings.Localizer[descriptor.GetTranslationKeyForColumn(name)];
            }

            if (excel.Workbook.Worksheets.Any(s => s.Name == name))
            {
                name = string.Format("{0} {1}", name, excel.Workbook.Worksheets.Count);
            }

            var worksheet = excel.Workbook.Worksheets.Add(name);
            worksheet.OutLineApplyStyle = true;

            return worksheet;
        }

        /// <summary>
        /// Translates value by localizer
        /// </summary>
        /// <param name="value"></param>
        /// <param name="descriptor"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        /// <exception cref="SimpleXlsException"></exception>
        private static object TranslateValue(object value, ColumnDescriptor descriptor, SheetExportSettings settings)
        {
            if (settings.HasLocalizer == false)
            {
                throw new SimpleXlsException("Property " + descriptor.Key +
                                             " has XlsTranslateAttribute but no Localizer was defined in export settings.");
            }
            if (value == null)
            {
                return null;
            }
            var str = (string) value;
            if (string.IsNullOrEmpty(str))
            {
                return value;
            }

            return settings.GetLocalizer()[descriptor.Attributes.DictionaryPrefix + str];
        }

        /// <summary>
        /// Deletes unused rows
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="totalCols"></param>
        /// <param name="columnUsages"></param>
        private static void OmitUnusedColumns(ExcelWorksheet worksheet, int totalCols, int[] columnUsages)
        {
            var deleted = 0;
            for (var col = 1; col < totalCols; col++)
            {
                if (columnUsages[col] != 0) continue;
                worksheet.DeleteColumn(col - deleted);
                deleted++;
            }
        }

        /// <summary>
        /// Creates worksheet header
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="descriptor"></param>
        /// <param name="settings"></param>
        private static void CreateWorksheetHeader(ExcelWorksheet worksheet,
            SheetDescriptor descriptor, SheetExportSettings settings)
        {
            var cntr = 1;
            foreach (var d in descriptor.Columns)
            {
                if (d.Attributes.Ignore)
                {
                    continue;
                }

                var value = d.Attributes.Heading;
                if (settings.HasLocalizer && settings.Translate)
                {
                    value = settings.GetLocalizer()[descriptor.GetTranslationKeyForColumn(d.Attributes.Heading)];
                }
                worksheet.Cells[1, cntr].Value = value;
                cntr++;
            }
        }
    }
}