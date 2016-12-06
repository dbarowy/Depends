using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace COMWrapper
{
    public class WorkbookOpenException : Exception { }

    public class Application : IDisposable
    {
        Excel.Application _app;
        List<Workbook> _wbs;

        public Application()
        {
            _app = new Excel.Application();
            _wbs = new List<Workbook>();
        }

        // All of the following private enums are poorly documented
        private enum XlCorruptLoad
        {
            NormalLoad = 0,
            RepairFile = 1,
            ExtractData = 2
        }

        private enum XlUpdateLinks
        {
            Yes = 2,
            No = 0
        }

        private enum XlPlatform
        {
            Macintosh = 1,
            Windows = 2,
            MSDOS = 3
        }

        public Workbook OpenWorkbook(string relpath)
        {
            // get the absolute path
            var abspath = System.IO.Path.GetFullPath(relpath);

            // we need to disable all alerts, e.g., password prompts, etc.
            _app.DisplayAlerts = false;

            // disable macros
            _app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // This call is stupid.  See:
            // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
            _app.Workbooks.Open(abspath,                    // FileName (String)
                               XlUpdateLinks.Yes,           // UpdateLinks (XlUpdateLinks enum)
                               true,                        // ReadOnly (Boolean)
                               Missing.Value,               // Format (int?)
                               "thisisnotapassword",        // Password (String)
                               Missing.Value,               // WriteResPassword (String)
                               true,                        // IgnoreReadOnlyRecommended (Boolean)
                               Missing.Value,               // Origin (XlPlatform enum)
                               Missing.Value,               // Delimiter; if the filetype is txt (String)
                               Missing.Value,               // Editable; not what you think (Boolean)
                               false,                       // Notify (Boolean)
                               Missing.Value,               // Converter(int)
                               false,                       // AddToMru (Boolean)
                               Missing.Value,               // Local; really "use my locale?" (Boolean)
                               XlCorruptLoad.RepairFile);   // CorruptLoad (XlCorruptLoad enum)

            // init wrapped workbook
            var wb_idx = _wbs.Count + 1; // Excel uses 1-based arrays
            var wbref = _app.Workbooks[wb_idx];

            // if the open call above failed, stop now
            if (wbref == null)
            {
                throw new WorkbookOpenException();
            }

            var wb = new Workbook(wbref, _app);

            // add to list
            _wbs.Add(wb);

            return wb;
        }

        public void CloseWorkbookByName(String name)
        {
            var wb = _wbs.Find( (Workbook w) => w.WorkbookName == name);
            CloseWorkbook(wb);
        }

        public void CloseWorkbook(Workbook wb)
        {
            wb.Dispose();
            _wbs = _wbs.Where((Workbook w) => w != wb).ToList();
        }

        public Excel.Application XLApplication()
        {
            return _app;
        }

        public void Dispose()
        {
            foreach (var wb in _wbs)
            {
                wb.Dispose();
            }
            _app.Quit();
            Marshal.ReleaseComObject(_app);
            _app = null;
        }
    }
}
