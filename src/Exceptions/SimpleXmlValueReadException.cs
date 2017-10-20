using System;

namespace Firefly.SimpleXls.Exceptions
{
    public class SimpleXlsValueReadException : Exception
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public string ModelName { get; set; }

        public SimpleXlsValueReadException(int row, int col, string modelName, string message) : base(message)
        {
            Row = row;
            Col = col;
            ModelName = modelName;
        }

        public SimpleXlsValueReadException(string message) : base(message)
        {
        }

        public SimpleXlsValueReadException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}