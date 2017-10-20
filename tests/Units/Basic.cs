using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Firefly.SimpleXls;
using Firefly.SimpleXlsTests.Fixtures;
using Xunit;

namespace Firefly.SimpleXlsTests.Units
{
    public class Basic
    {
        private const string TempFile = "./testing.xlsx";

        [Fact]
        public void FileBasic()
        {
            var firstPerson = Person.Create1();
            var secondPerson = Person.Create2();
            var data = new[]
            {
                firstPerson,
                secondPerson
            };

            Exporter.CreateNew().AddSheet(data,
                    settings => { settings.OmitEmptyColumns = false; }
                )
                .Export(TempFile);

            var check = Importer.Open(TempFile).ImportAs<Person>(1, settings => { settings.BreakOnError = true; });
            Assert.NotNull(check);
            CompareResults(check, firstPerson, secondPerson);
        }

        [Fact]
        public void StreamBasic()
        {
            var firstPerson = Person.Create1();
            var secondPerson = Person.Create2();
            var data = new[]
            {
                firstPerson,
                secondPerson
            };

            using (var s = new MemoryStream())
            {
                Exporter.CreateNew().AddSheet(data,
                        settings => { settings.OmitEmptyColumns = false; }
                    )
                    .Export(s);

                var check = Importer.Open(s).ImportAs<Person>(1, settings => { settings.BreakOnError = true; });
                Assert.NotNull(check);
                CompareResults(check, firstPerson, secondPerson);
            }
        }

        [Fact]
        public void FileInfoBasic()
        {
            var firstPerson = Person.Create1();
            var secondPerson = Person.Create2();
            var data = new[]
            {
                firstPerson,
                secondPerson
            };

            var info = new FileInfo(TempFile);

            Exporter.CreateNew().AddSheet(data,
                    settings => { settings.OmitEmptyColumns = false; }
                )
                .Export(info);

            var check = Importer.Open(info).ImportAs<Person>(1, settings => { settings.BreakOnError = true; });
            Assert.NotNull(check);
            CompareResults(check, firstPerson, secondPerson);
        }

        private static void CompareResults(IReadOnlyCollection<Person> check, Person firstPerson, Person secondPerson)
        {
            Assert.Equal(2, check.Count);
            var p1 = check.First();
            Assert.Equal(firstPerson.Age, p1.Age);
            Assert.Null(firstPerson.AlwaysNullDateTime);
            Assert.Equal(firstPerson.Birthday, p1.Birthday);
            Assert.Equal(firstPerson.Height, p1.Height);
            Assert.Equal(firstPerson.Money, p1.Money);
            Assert.Equal(firstPerson.Name, p1.Name);
            Assert.Null(firstPerson.ThisFieldIsAlwaysNull);
            Assert.Equal(firstPerson.WorkHours, p1.WorkHours);
            Assert.Equal(firstPerson.IsAlcoholic, p1.IsAlcoholic);

            var p2 = check.Last();
            Assert.Equal(secondPerson.Age, p2.Age);
            Assert.Null(secondPerson.AlwaysNullDateTime);
            Assert.Equal(secondPerson.Birthday, p2.Birthday);
            Assert.Equal(secondPerson.Height, p2.Height);
            Assert.Equal(secondPerson.Money, p2.Money);
            Assert.Equal(secondPerson.Name, p2.Name);
            Assert.Null(secondPerson.ThisFieldIsAlwaysNull);
            Assert.Equal(secondPerson.WorkHours, p2.WorkHours);
            Assert.Equal(secondPerson.IsAlcoholic, p2.IsAlcoholic);
        }
    }
}