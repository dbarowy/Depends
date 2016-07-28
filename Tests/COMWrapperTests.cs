using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using COMWrapper;

namespace ExceLintTests
{
    [TestClass]
    public class COMWrapperTests
    {
        [TestMethod]
        public void WorkbookIndexTest()
        {
            var path = "../../../../../data/analyses/CUSTODES/custodes/example/input_spreadsheets";
            var wb1_name = "01-38-PK_tables-figures.xls";
            var wb2_name = "chartssection2.xls";
            var wb3_name = "gradef03-sec3.xls";
            var wb4_name = "ribimv001.xls";
            var wb5_name = "p36.xls";

            var app = new Application();

            var wb1 = app.OpenWorkbook(System.IO.Path.Combine(path, wb1_name));
            Assert.AreEqual(wb1.WorkbookName, wb1_name);

            var wb2 = app.OpenWorkbook(System.IO.Path.Combine(path, wb2_name));
            Assert.AreEqual(wb2.WorkbookName, wb2_name);

            app.CloseWorkbook(wb2);

            var wb3 = app.OpenWorkbook(System.IO.Path.Combine(path, wb3_name));
            Assert.AreEqual(wb3.WorkbookName, wb3_name);

            var wb4 = app.OpenWorkbook(System.IO.Path.Combine(path, wb4_name));
            Assert.AreEqual(wb4.WorkbookName, wb4_name);

            app.CloseWorkbook(wb1);

            var wb5 = app.OpenWorkbook(System.IO.Path.Combine(path, wb5_name));
            Assert.AreEqual(wb5.WorkbookName, wb5_name);
        }
    }
}
