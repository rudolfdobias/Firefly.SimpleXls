using System;

namespace Firefly.SimpleXls.Attributes
{
    public class XlsTranslateAttribute : Attribute
    {
        public string DictPrefix { get; set; }
    }
}