using System.Collections.Generic;

namespace Firefly.SimpleXls
{
    public class RawTable
    {
        public string[] Headers { get; set; } = new string[0];
        public List<object[]> Values { get; } = new List<object[]>();
    }
}