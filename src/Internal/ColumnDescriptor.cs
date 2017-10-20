using System.Reflection;
using Firefly.SimpleXls.Converters;

namespace Firefly.SimpleXls.Internal
{
    internal struct ColumnDescriptor
    {
        public string Key { get; set; }
        public PropertyInfo Property { get; set; }
        public ColumnAttributeInfo Attributes { get; set; }
        public IValueConverter CustomValueConverter { get; set; }
    }
}