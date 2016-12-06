using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace COMWrapperTests
{
    [TestClass]
    public class BasicTests
    {
        [TestMethod]
        public void TestMagicBytesXLS()
        {
            var xls_f = System.IO.Path.GetFullPath("../../testfiles/AnXLS.xls");
            var ft = COMWrapper.Application.MagicBytes(xls_f);
            Assert.AreEqual(COMWrapper.Application.CWFileType.XLS, ft);
        }

        [TestMethod]
        public void TestMagicBytesXLSX()
        {
            var xlsx_f = System.IO.Path.GetFullPath("../../testfiles/AnXLSX.xlsx");
            var ft = COMWrapper.Application.MagicBytes(xlsx_f);
            Assert.AreEqual(COMWrapper.Application.CWFileType.XLSX, ft);
        }
    }
}
