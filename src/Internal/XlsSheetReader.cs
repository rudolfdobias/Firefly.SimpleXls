using System;
using System.Collections.Generic;
using System.Linq;
using Firefly.SimpleXls.Exceptions;
using OfficeOpenXml;

namespace Firefly.SimpleXls.Internal
{
    /// <summary>
    /// Imports excel sheets. (At least trying.)
    /// </summary>
    internal static class XlsSheetReader
    {
        /// <summary>
        /// Raw import of excel sheet
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="idx"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        /// <exception cref="SimpleXlsValueReadException"></exception>
        public static RawTable ReadSheet(ExcelPackage excel, int idx,
            SheetImportSettings settings)
        {
            var data = new RawTable();
            var sheet = GetSheet(excel, idx);
            var totalCols = sheet.Dimension.Columns;
            var totalRows = sheet.Dimension.Rows;
            if (totalRows == 0 || totalCols == 0)
            {
                return data;
            }

            var startingRow = 1;
            if (settings.HasHeader)
            {
                data.Headers = ReadHeader(sheet, totalCols);
                startingRow = 2;
            }

            for (var row = startingRow; row <= totalRows; row++)
            {
                var line = new object[totalCols];
                for (var col = 1; col <= totalCols; col++)
                {
                    try
                    {
                        line[col] = sheet.Cells[row, col].Value;
                    }
                    catch (Exception crap)
                    {
                        if (settings.BreakOnError)
                        {
                            throw new SimpleXlsValueReadException(row, col, null,
                                "Cannot import value at " + row + ":" + col + " / " + crap.Message);
                        }
                    }
                }
                data.Values.Add(line);
            }

            return data;
        }

        /// <summary>
        /// Reads sheet and tries to bind data to specified model.
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="idx"></param>
        /// <param name="settings"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <exception cref="SimpleXlsValueReadException"></exception>
        public static List<T> ReadSheet<T>(ExcelPackage excel, int idx,
            SheetImportSettings settings)
            where T : class, new()
        {
            var data = new List<T>();
            var sheet = GetSheet(excel, idx);
            var totalCols = sheet.Dimension.Columns;
            var totalRows = sheet.Dimension.Rows;
            if (totalRows == 0 || totalCols == 0)
            {
                return data;
            }

            var startingRow = settings.HasHeader ? 2 : 1;
            var descriptor = ModelDescriber.DescribeModel<T>();
            var maxCol = Math.Min(totalCols, descriptor.Columns.Count);
            var descriptorValues = descriptor.Columns.Where(d => d.Attributes.Ignore == false).ToArray();

            for (var row = startingRow; row <= totalRows; row++)
            {
                var item = new T();
                for (var col = 1; col <= maxCol; col++)
                {
                    var info = descriptorValues[col - 1];
                    try
                    {
                        // todo: Speed it up with cached expressions
                        var value = sheet.Cells[row, col].Value;

                        if (info.CustomValueConverter != null)
                        {
                            info.Property.SetValue(item, info.CustomValueConverter.Read(value), null);
                        }
                        else
                        {
                            if (value == null)
                            {
                                continue;
                            }

                            info.Property.SetValue(item,
                                Convert.ChangeType(value, info.Property.PropertyType), null);
                        }
                    }
                    catch (Exception crap)
                    {
                        if (settings.BreakOnError)
                        {
                            throw new SimpleXlsValueReadException(row, col, info.Key,
                                "Cannot parse value: " + crap.Message);
                        }
                    }
                }
                data.Add(item);
            }

            return data;
        }

        /// <summary>
        /// Returns header of table
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        private static string[] ReadHeader(ExcelWorksheet sheet, int cols)
        {
            var headers = new string[cols];
            for (var i = 1; i <= cols; i++)
            {
                headers[i] = (string) sheet.Cells[1, i].Value;
            }

            return headers;
        }

        /// <summary>
        /// Gets sheet from excel by idx
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="idx"></param>
        /// <returns></returns>
        /// <exception cref="SimpleXlsException"></exception>
        private static ExcelWorksheet GetSheet(ExcelPackage excel, int idx)
        {
            if (excel.Workbook.Worksheets.Count == 0)
            {
                throw new SimpleXlsException("The document is empty!");
            }

            try
            {
                return excel.Workbook.Worksheets[idx];
            }
            catch (IndexOutOfRangeException)
            {
                throw new SimpleXlsException("Sheet with index " + idx + " does not exist.");
            }
        }
    }
}