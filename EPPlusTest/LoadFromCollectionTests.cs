using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlusTest
{
    [TestClass]
    public class LoadFromCollectionTests
    {
        internal abstract class BaseClass
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }

        internal class Implementation : BaseClass
        {
            public int Number { get; set; }
        }

        internal class Aclass
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public int Number { get; set; }
        }

        internal class Competitor
        {
            [System.ComponentModel.Description("Competitor Name")]
            public string Name { get; set; }
            
            [System.ComponentModel.Description("Competitor Rank")]
            public Rank Rank { get; set; }
        }

        public enum Rank
        {
            [System.ComponentModel.Description("1st")]
            First,

            [System.ComponentModel.Description("2nd")]
            Second,

            [System.ComponentModel.Description("3rd")]
            Third
        }

        [TestMethod]
        public void ShouldUseAclassProperties()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseBaseClassProperties()
        {
            var items = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseAnonymousProperties()
        {
            var objs = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void ShouldThrowInvalidCastExceptionIf()
        {
            var objs = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, BindingFlags.Public | BindingFlags.Instance, typeof(string).GetMembers());

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseDisplayOnEnumValues()
        {
            var competitors = new List<Competitor>();
            competitors.Add(new Competitor { Name = "Carlsen, Magnus", Rank = Rank.First });
            competitors.Add(new Competitor { Name = "Caruana, Fabiano", Rank = Rank.Second });
            competitors.Add(new Competitor { Name = "Ding, Liren", Rank = Rank.Third });


            using (var package = new ExcelPackage(new MemoryStream()))
            {
                var sheet = package.Workbook.Worksheets.Add("sheet");
                sheet.Cells["A1"].LoadFromCollection(competitors, true, TableStyles.Dark1);

                Assert.AreEqual("Competitor Name", sheet.Cells["A1"].Value);
                Assert.AreEqual("1st", sheet.Cells["B2"].Value);
            }
        }
    }
}
