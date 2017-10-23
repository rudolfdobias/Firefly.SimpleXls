namespace Firefly.SimpleXls.Internal
{
    internal class ColumnAttributeInfo
    {
        public bool Ignore { get; set; }
        public string Heading { get; set; }
        public bool TranslateValue { get; set; }
        public string DictionaryPrefix { get; set; } = "";
    }
}