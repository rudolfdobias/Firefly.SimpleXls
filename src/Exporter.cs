using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Firefly.SimpleXls.Exceptions;
using Firefly.SimpleXls.Internal;
using OfficeOpenXml;

namespace Firefly.SimpleXls
{
    /// <summary>
    /// Manages exporting data models to XLS
    /// </summary>
    public class Exporter : IDisposable
    {
        /// <summary>
        /// String used as key hidden in xls for metadata
        /// </summary>
        internal const string ExporterXmlKey = "Firefly.SimpleXls";

        /// <summary>
        /// Where the XLS is located
        /// </summary>
        private MemoryStream Buffer { get; } = new MemoryStream();

        /// <summary>
        /// The main excel
        /// </summary>
        private ExcelPackage Excel { get; }

        /// <summary>
        /// Protected constructor
        /// </summary>
        protected Exporter()
        {
            Excel = new ExcelPackage(Buffer);
            Excel.Workbook.Properties.CustomPropertiesXml.CreateElement(ExporterXmlKey);
        }

        /// <summary>
        /// Creates new exporting instance
        /// </summary>
        /// <returns></returns>
        public static Exporter CreateNew()
        {
            return new Exporter();
        }

        /// <summary>
        /// Adds named sheet to the file filled with the data from export
        /// </summary>
        /// <param name="models"></param>
        /// <param name="settings"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public Exporter AddSheet<T>(IEnumerable<T> models, Action<SheetExportSettings> settings = null)
            where T : class, new()
        {
            var exportSettings = new SheetExportSettings();
            settings?.Invoke(exportSettings);

            try
            {
                XlsSheetWriter.WriteSheet(Excel, models, exportSettings);
            }
            catch (SimpleXlsException)
            {
                throw;
            }
            catch (Exception crap)
            {
                throw new SimpleXlsException(crap.Message, crap);
            }
            return this;
        }

        /// <summary>
        /// Exports excel to file.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="overwrite"></param>
        /// <exception cref="SimpleXlsException"></exception>
        public void Export(string filename, bool overwrite = true)
        {
            CheckBeforeExport();
            var file = new FileInfo(filename);
            if (file.Exists)
            {
                if (!overwrite)
                    throw new SimpleXlsException("File " + filename +
                                                 " already exists. Use overwrite = true for replacing.");

                file.Delete();
            }
            using (var stream = file.Create())
            {
                Excel.SaveAs(stream);
            }
        }

        /// <summary>
        /// Exports excel to file.
        /// </summary>
        /// <param name="file"></param>
        public void Export(FileInfo file)
        {
            CheckBeforeExport();
            Excel.SaveAs(file);
        }

        /// <summary>
        /// Writes excel to stream
        /// </summary>
        /// <param name="stream"></param>
        public void Export(Stream stream)
        {
            CheckBeforeExport();
            Excel.SaveAs(stream);
        }

        /// <summary>
        /// Throws if export cannot be done.
        /// </summary>
        /// <exception cref="SimpleXlsException"></exception>
        private void CheckBeforeExport()
        {
            if (Excel.Workbook.Worksheets.Any() == false)
            {
                throw new SimpleXlsException("Document does not contain any sheets. Use AddSheet() before.");
            }
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Buffer?.Dispose();
            Excel?.Dispose();
        }
    }
}