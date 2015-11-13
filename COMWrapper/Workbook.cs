using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Depends;

namespace COMWrapper
{
    public class Workbook : IDisposable
    {
        private Excel.Application _app;
        private Excel.Workbook _wb;

        public Workbook(Excel.Workbook wb, Excel.Application app)
        {
            _app = app;
            _wb = wb;
        }

        public void Dispose()
        {
            _wb.Close();
            Marshal.ReleaseComObject(_wb);
            _wb = null;
        }

        public DAG buildDependenceGraph()
        {
            return new DAG(_wb, _app, true);
        }

        public string WorkbookName
        {
            get { return _wb.Name; }
        }

        public string WorksheetName(int index)
        {
            // worksheets are 1-indexed
            Microsoft.Office.Interop.Excel.Worksheet ws = _wb.Worksheets[index];
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
