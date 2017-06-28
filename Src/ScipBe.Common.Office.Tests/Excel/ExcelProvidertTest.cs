using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScipBe.Common.Office.Excel;
using System;
using System.Linq;

namespace ScipBe.Common.Office.Tests
{
    [TestClass]
    public class ExcelProvidertTest
    {
        [TestMethod]
        public void LoadXlsxFile()
        {
            var fileName = $@"{AppDomain.CurrentDomain.BaseDirectory}\Excel\Persons.xlsx";
            LoadFile(fileName, "Persons");
        }

        [TestMethod]
        public void LoadXlsFile()
        {
            var fileName = $@"{AppDomain.CurrentDomain.BaseDirectory}\Excel\Persons.xls";
            LoadFile(fileName, "Persons");
        }

        [TestMethod]
        public void LoadCsvSemicolumnFile()
        {
            var fileName = $@"{AppDomain.CurrentDomain.BaseDirectory}\Excel\PersonsSemicolumn.csv";
            LoadFile(fileName);
        }

        [TestMethod]
        public void LoadCsvCommaFile()
        {
            var fileName = $@"{AppDomain.CurrentDomain.BaseDirectory}\Excel\PersonsComma.csv";
            LoadFile(fileName);
        }

        [TestMethod]
        public void LoadCsvTabFile()
        {
            var fileName = $@"{AppDomain.CurrentDomain.BaseDirectory}\Excel\PersonsTab.csv";
            LoadFile(fileName);
        }

        private void LoadFile(string fileName, string workSheetName = null)
        {            
            var excel = new ExcelProvider(fileName, workSheetName);

            Assert.AreEqual(5, excel.Columns.Count());

            Assert.AreEqual(1, excel.Columns.First().Index);
            Assert.AreEqual("A", excel.Columns.First().Header);
            Assert.AreEqual("ID", excel.Columns.First().Name);

            Assert.AreEqual(5, excel.Columns.Last().Index);
            Assert.AreEqual("E", excel.Columns.Last().Header);
            Assert.AreEqual("BirthDate", excel.Columns.Last().Name);

            Assert.AreEqual(5, excel.Rows.Count());

            Assert.AreEqual(5, excel.Rows.Last().Index);
            Assert.AreEqual(5, excel.Rows.Last().Get<int>(1));
            Assert.AreEqual(5, excel.Rows.Last().Get<int>("A"));
            Assert.AreEqual(5, excel.Rows.Last().GetByName<int>("ID"));

            Assert.AreEqual("Peter", excel.Rows.Last().Get<string>(2));
            Assert.AreEqual("Peter", excel.Rows.Last().Get<string>("B"));
            Assert.AreEqual("Peter", excel.Rows.Last().GetByName<string>("FirstName"));

            Assert.AreEqual(DateTime.Parse("3/05/1979 0:00:00"), excel.Rows.Last().GetByName<DateTime>("BirthDate"));
        }
    }
}
