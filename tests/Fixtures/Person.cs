using System;
using System.Reflection;
using Firefly.SimpleXls;
using Firefly.SimpleXls.Attributes;

namespace Firefly.SimpleXlsTests.Fixtures
{
    public class Person
    {
        [XlsHeader(Name = "Name column")]
        public string Name { get; set; }

        public int Age { get; set; }
        public DateTime Birthday { get; set; }
        public float Height { get; set; }
        public decimal Money { get; set; }
        public TimeSpan WorkHours { get; set; }
        public string ThisFieldIsAlwaysNull { get; set; } = null;
        public DateTime? AlwaysNullDateTime { get; set; } = null;
        public string ThisStringIsSometimesNull { get; set; }
        public bool IsAlcoholic { get; set; }

        [XlsIgnore]
        public Guid FacebookId { get; set; } = Guid.NewGuid();

        public static Person Create1()
        {
            return new Person
            {
                Age = 27,
                Birthday = new DateTime(1990, 02, 14),
                Height = 179.1f,
                Money = 123456.78m,
                Name = "Theodor Roosevelt",
                WorkHours = TimeSpan.FromHours(14),
                IsAlcoholic = true
            };
        }

        public static Person Create2()
        {
            return new Person
            {
                Age = 24,
                Birthday = new DateTime(1990, 03, 07),
                Height = 169.1f,
                Money = 163456.78m,
                Name = "Alice vas Neverland",
                WorkHours = TimeSpan.FromHours(8)
            };
        }
    }
}