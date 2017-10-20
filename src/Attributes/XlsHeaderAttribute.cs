using System;

namespace Firefly.SimpleXls.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class XlsHeaderAttribute : Attribute
    {
        public string Name { get; set; }
    }
}