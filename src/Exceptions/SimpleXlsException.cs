using System;

namespace Firefly.SimpleXls.Exceptions
{
    public class SimpleXlsException : Exception
    {
        public SimpleXlsException(string message) : base(message)
        {
        }

        public SimpleXlsException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}