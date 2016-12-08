using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace COMWrapper
{
    public class WorkbookOpenException : Exception {}

    public class Application : IDisposable
    {
        Excel.Application _app;
        Dictionary<string,Workbook> _wbs;

        public Application()
        {
            _app = new Excel.Application();

            // set app properties once
            _app.AskToUpdateLinks = false;
            _app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            _app.DisplayAlerts = false;

            _wbs = new Dictionary<string,Workbook>();
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

        public enum CWFileType
        {
            XLS,
            XLSX,
            Unknown
        }

        public static CWFileType MagicBytes(string fileabspath)
        {
            byte[] xlsx1 = { 0x50, 0x4B, 0x03, 0x04 };
            byte[] xlsx2 = { 0x50, 0x4B, 0x05, 0x06 };
            byte[] xlsx3 = { 0x50, 0x4B, 0x07, 0x08 };
            byte[] xls = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

            byte[] first8 = System.IO.File.ReadAllBytes(fileabspath).Take(8).ToArray();
            byte[] first4 = first8.Take(4).ToArray();

            if (first8.SequenceEqual(xls))
            {
                return CWFileType.XLS;
            } else if (first4.SequenceEqual(xlsx1)
                       || first4.SequenceEqual(xlsx2)
                       || first4.SequenceEqual(xlsx3))
            {
                return CWFileType.XLSX;
            } else
            {
                return CWFileType.Unknown;
            }
        }

        public Workbook OpenWorkbook(string relpath)
        {
            // get the absolute path
            var abspath = System.IO.Path.GetFullPath(relpath);

            // make sure that this is actually an Excel file
            var ft = MagicBytes(abspath);
            if (ft != CWFileType.XLS && ft != CWFileType.XLSX)
            {
                throw new WorkbookOpenException();
            }

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

            
            // Excel uses 1-based arrays
            var wb_idx = _app.Workbooks.Count;
            var wbref = _app.Workbooks[wb_idx];

            // if the open call above failed, stop now
            if (wbref == null)
            {
                throw new WorkbookOpenException();
            }

            // do not autorecover!
            wbref.EnableAutoRecover = false;

            // if this workbook has links, break them
            var links = (Array) wbref.LinkSources(Excel.XlLink.xlExcelLinks);
            if (links != null)
            {
                for (int i = 1; i <= links.Length; i++)
                {
                    wbref.BreakLink((string)links.GetValue(i), Excel.XlLinkType.xlLinkTypeExcelLinks);
                }
            }

            // create callback to remove workbook from _wbs on Dispose()
            string wbname = wbref.Name;
            Action dcb = () => _wbs.Remove(wbname);  // MUST use copy of wbref.Name here!

            // init wrapped workbook
            var wb = new Workbook(wbref, _app, dcb);

            // add to list
            _wbs.Add(wb.WorkbookName, wb);

            return wb;
        }

        public void CloseWorkbookByName(String name)
        {
            CloseWorkbook(_wbs[name]);
        }

        public void CloseWorkbook(Workbook wb)
        {
            wb.Dispose();
        }

        public Excel.Application XLApplication()
        {
            return _app;
        }

        public void Dispose()
        {
            foreach (var pair in _wbs)
            {
                pair.Value.Dispose();
            }
            _wbs = null;

            _app.Quit();
            Marshal.ReleaseComObject(_app);
            _app = null;
        }
    }
}
