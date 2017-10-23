using System;
using System.Collections.Generic;

namespace Firefly.SimpleXls.Internal
{
    internal class SheetDescriptor
    {
        public List<ColumnDescriptor> Columns { get; set; } = new List<ColumnDescriptor>();
        public string Name { get; set; }
        public Type ModelType { get; set; }
        public string DictionaryPrefix { get; set; }

        public string GetTranslationKeyForColumn(string something)
            => string.Format("{0}{1}", DictionaryPrefix, something);
    }
}