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

        [TestMethod]
        public void levelComputation()
        {
            var filename = @"..\..\TestData\LevelsDAG.xlsx";
            var app = new COMWrapper.Application();
            var wb = app.OpenWorkbook(filename);
            var dag = wb.buildDependenceGraph();

            Func<string,int,AST.Address> fastCell =
                (string col, int row) =>
                AST.Address.fromA1withMode(row, col, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wb.WorksheetName(1), wb.WorkbookName, wb.Path);

            Func<AST.Address, AST.Address, int[], bool> expectedDistances =
                    (AST.Address from, AST.Address to, int[] distances) =>
                    {
                        var tbl = dag.AllRefDistancesFromInput(from);
                        HashSet<int> expected = new HashSet<int>(distances);
                        HashSet<int> actual = tbl.ContainsKey(to) ? tbl[to] : new HashSet<int>();
                        Func<int,int> identity = (int i) => i;
                        return actual.OrderBy(identity).SequenceEqual(expected);
                    };


            var a1 = fastCell("A", 1);
            var b1 = fastCell("B", 1);
            var c1 = fastCell("C", 1);
            var d1 = fastCell("D", 1);
            var e1 = fastCell("E", 1);
            var f1 = fastCell("F", 1);
            var g1 = fastCell("G", 1);
            var h1 = fastCell("H", 1);
            var i1 = fastCell("I", 1);
            var j1 = fastCell("J", 1);

            // from A1
            int[] a1_to_a1 = { };
            Assert.IsTrue(expectedDistances(a1, a1, a1_to_a1));

            // from B1
            int[] b1_to_a1 = { 1 };
            Assert.IsTrue(expectedDistances(b1, a1, b1_to_a1));

            // from C1
            int[] c1_to_a1 = { 1 };
            Assert.IsTrue(expectedDistances(c1, a1, c1_to_a1));

            // from D1
            int[] d1_to_a1 = { 1 };
            Assert.IsTrue(expectedDistances(d1, a1, d1_to_a1));

            // from E1
            int[] e1_to_a1 = { };
            Assert.IsTrue(expectedDistances(e1, a1, e1_to_a1));

            // from F1
            int[] f1_to_a1 = { 2 };
            Assert.IsTrue(expectedDistances(f1, a1, f1_to_a1));
            int[] f1_to_c1 = { 1 };
            Assert.IsTrue(expectedDistances(f1, c1, f1_to_c1));

            // from G1
            int[] g1_to_a1 = { 2 };
            Assert.IsTrue(expectedDistances(g1, a1, g1_to_a1));
            int[] g1_to_c1 = { 1 };
            Assert.IsTrue(expectedDistances(g1, c1, g1_to_c1));

            // from H1
            int[] h1_to_a1 = { 2 };
            Assert.IsTrue(expectedDistances(h1, a1, h1_to_a1));
            int[] h1_to_d1 = { 1 };
            Assert.IsTrue(expectedDistances(h1, d1, h1_to_d1));

            // from I1
            int[] i1_to_a1 = { 2, 3 };
            Assert.IsTrue(expectedDistances(i1, a1, i1_to_a1));
            int[] i1_to_c1 = { 2 };
            Assert.IsTrue(expectedDistances(i1, c1, i1_to_c1));
            int[] i1_to_d1 = { 1, 2 };
            Assert.IsTrue(expectedDistances(i1, d1, i1_to_d1));
            int[] i1_to_e1 = { 2 };
            Assert.IsTrue(expectedDistances(i1, e1, i1_to_e1));
            int[] i1_to_g1 = { 1 };
            Assert.IsTrue(expectedDistances(i1, g1, i1_to_g1));
            int[] i1_to_h1 = { 1 };
            Assert.IsTrue(expectedDistances(i1, h1, i1_to_h1));

            // from J1
            int[] j1_to_a1 = { 3 };
            Assert.IsTrue(expectedDistances(j1, a1, j1_to_a1));
            int[] j1_to_d1 = { 2 };
            Assert.IsTrue(expectedDistances(j1, d1, j1_to_d1));
            int[] j1_to_e1 = { 1, 2 };
            Assert.IsTrue(expectedDistances(j1, e1, j1_to_e1));
            int[] j1_to_h1 = { 1 };
            Assert.IsTrue(expectedDistances(j1, h1, j1_to_h1));
        }
    }
}
