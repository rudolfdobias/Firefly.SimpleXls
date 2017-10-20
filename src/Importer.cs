using System;
using System.Collections.Generic;
using System.IO;
using Firefly.SimpleXls.Exceptions;
using Firefly.SimpleXls.Internal;
using OfficeOpenXml;

namespace Firefly.SimpleXls
{
    /// <summary>
    /// Imports data from XLS(X) files.
    /// </summary>
    public class Importer : IDisposable
    {
        /// <summary>
        /// The main excel
        /// </summary>
        private ExcelPackage Excel { get; }

        /// <summary>
        /// Protected ctor. Opens document from stream.
        /// </summary>
        /// <param name="stream"></param>
        /// <exception cref="SimpleXlsException"></exception>
        protected Importer(Stream stream)
        {
            if (stream.CanRead == false)
            {
                throw new SimpleXlsException("Cannot read from the stream!");
            }
            Excel = new ExcelPackage(stream);
        }

        /// <summary>
        /// Opens document from filename
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static Importer Open(string filename)
        {
            var info = new FileInfo(filename);
            return Open(info);
        }

        /// <summary>
        /// Opens document from FileInfo
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        /// <exception cref="SimpleXlsException"></exception>
        public static Importer Open(FileInfo info)
        {
            if (info.Exists == false)
            {
                throw new SimpleXlsException("Cannot read file " + info.FullName);
            }

            return new Importer(info.OpenRead());
        }

        /// <summary>
        /// Opens document from stream
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static Importer Open(Stream stream)
        {
            return new Importer(stream);
        }

        /// <summary>
        /// Gets internal ExcelWorkbook for custom purposes.
        /// </summary>
        /// <returns></returns>
        public ExcelWorkbook ImportWorkbook()
        {
            return Excel.Workbook;
        }

        /// <summary>
        /// EXPERIMENTAL - 
        /// Imports sheet by its index as a generic type.
        /// Sheet index starts by 1.
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        /// <exception cref="SimpleXlsException"></exception>
        public RawTable ImportAsRaw(int sheetIndex = 1,
            Action<SheetImportSettings> settings = null)
        {
            var importSettings = new SheetImportSettings();
            settings?.Invoke(importSettings);
            try
            {
                return XlsSheetReader.ReadSheet(Excel, sheetIndex, importSettings);
            }
            catch (SimpleXlsException)
            {
                throw;
            }
            catch (Exception crap)
            {
                throw new SimpleXlsException("Import failed: " + crap.Message, crap);
            }
        }

        /// <summary>
        /// EXPERIMENTAL - 
        /// Tries to map excel sheet to a model.
        /// The excel data MUST be in correct order and form.
        /// NO columns must be omitted.
        /// Use only for known data.
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="settings"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <exception cref="SimpleXlsException"></exception>
        public List<T> ImportAs<T>(int sheetIndex = 1, Action<SheetImportSettings> settings = null)
            where T : class, new()
        {
            var importSettings = new SheetImportSettings();
            settings?.Invoke(importSettings);
            try
            {
                return XlsSheetReader.ReadSheet<T>(Excel, sheetIndex, importSettings);
            }
            catch (SimpleXlsException)
            {
                throw;
            }
            catch (SimpleXlsValueReadException)
            {
                throw;
            }
            catch (Exception crap)
            {
                throw new SimpleXlsException("Import failed: " + crap.Message, crap);
            }
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Excel?.Dispose();
        }
    }
}