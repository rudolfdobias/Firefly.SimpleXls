using System;

namespace Firefly.SimpleXls.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class XlsSheetAttribute : Attribute
    {
        public string Name { get; set; }
        public string DictionaryPrefix { get; set; } = "";
    }
}