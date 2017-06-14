using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Tests
{
    [TestClass]
    public class BasicTests
    {
        [TestMethod]
        public void checkForVLOOKUP()
        {
            // (1,1)    =VLOOKUP(B1,$C$1:$D$5,2,FALSE)
            var filename = @"..\..\TestData\VHLOOKUP.xlsx";
            var app = new COMWrapper.Application();
            var wb = app.OpenWorkbook(filename);
            var dag = wb.buildDependenceGraph();

            var addr_a1 = AST.Address.fromA1withMode(1, "A", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wb.WorksheetName(1), wb.WorkbookName, wb.Path);
            var frm_a1 = dag.getFormulaAtAddress(addr_a1);

            Assert.AreEqual("=VLOOKUP(B1,$C$1:$D$5,2,FALSE)", frm_a1);
        }

        [TestMethod]
        public void checkForHLOOKUP()
        {
            // (1,6)    =HLOOKUP(B1,G1:K2,2,FALSE)
            var filename = @"..\..\TestData\VHLOOKUP.xlsx";
            var app = new COMWrapper.Application();
            var wb = app.OpenWorkbook(filename);
            var dag = wb.buildDependenceGraph();

            var addr_a1 = AST.Address.fromA1withMode(1, "F", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wb.WorksheetName(1), wb.WorkbookName, wb.Path);
            var frm_a1 = dag.getFormulaAtAddress(addr_a1);

            Assert.AreEqual("=HLOOKUP(B1,G1:K2,2,FALSE)", frm_a1);
        }
    }
}
