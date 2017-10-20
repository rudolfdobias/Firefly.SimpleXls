﻿using System.Collections.Generic;
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
            var name = settings.SheetName;
            if (string.IsNullOrEmpty(name))
            {
                name = typeof(T).Name;
            }
            if (excel.Workbook.Worksheets.Any(s => s.Name == name))
            {
                throw new SimpleXlsException("A sheet named " + name + " already exists in this document.");
            }

            var worksheet = excel.Workbook.Worksheets.Add(name);
            worksheet.OutLineApplyStyle = true;

            var descriptors = ModelDescriptor.DescribeModel<T>();
            CreateWorksheetHeader(worksheet, descriptors);

            var columnUsages = new int[descriptors.Count + 1];
            columnUsages.Initialize();


            var row = 2; // 1 = table header, 2 = first row. Excel ...
            foreach (var i in items)
            {
                var col = 1;
                foreach (var d in descriptors.Values)
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

                    worksheet.Cells[row, col].Value = value;
                    col++;
                }
                row++;
            }


            if (settings.OmitEmptyColumns)
            {
                OmitUnusedColumns(worksheet, descriptors.Count, columnUsages);
            }

            excel.Save();
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
        /// <param name="descriptors"></param>
        private static void CreateWorksheetHeader(ExcelWorksheet worksheet,
            Dictionary<string, ColumnDescriptor> descriptors)
        {
            var cntr = 1;
            foreach (var d in descriptors.Values)
            {
                if (d.Attributes.Ignore)
                {
                    continue;
                }
                worksheet.Cells[1, cntr].Value = d.Attributes.Heading;
                cntr++;
            }
        }
    }
}