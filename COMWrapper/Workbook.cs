using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Depends;
using System.Collections.Generic;

namespace COMWrapper
{
    public class Workbook : IDisposable
    {
        private Excel.Application _app;
        private Excel.Workbook _wb;
        private String _wb_name;
        private bool _fetched_graph = false;
        private DAG.RawGraph _raw_graph;

        public Workbook(Excel.Workbook wb, Excel.Application app)
        {
            _app = app;
            _wb = wb;
            _wb_name = wb.Name;
        }

        public void Dispose()
        {
            if (_wb != null)
            {
                _wb.Close(SaveChanges: false);
                Marshal.ReleaseComObject(_wb);
                _wb = null;
            }
        }

        public DAG buildDependenceGraph()
        {
            return new DAG(_wb, _app, true);
        }

        public Dictionary<AST.Address,string> Formulas
        {
            get
            {
                if (!_fetched_graph)
                {
                    _raw_graph = DAG.FastFormulaRead(_wb);
                    _fetched_graph = true;
                }
                return _raw_graph.formulas;
            }
        }

        public string WorkbookName
        {
            get { return _wb.Name; }
        }

        public string WorksheetName(int index)
        {
            // worksheets are 1-indexed
            Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)_wb.Worksheets[index];
            var name = ws.Name;
            Marshal.ReleaseComObject(ws);
            return name;
        }

        public string Path
        {
            get { return _wb.Path;  }
        }
    }
}
