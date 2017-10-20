namespace Firefly.SimpleXls.Internal
{
    internal struct ColumnAttributeInfo
    {
        public bool Ignore { get; set; }
        public string Heading { get; set; }
        public bool TranslateValue { get; set; }
        public string DictionaryPrefix { get; set; }
    }
}