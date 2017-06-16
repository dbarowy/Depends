using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using CellRefDict = Depends.BiDictionary<AST.Address, ParcelCOMShim.COMRef>;
using VectorRefDict = Depends.BiDictionary<AST.Range, ParcelCOMShim.COMRef>;
using FormulaDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using InputDict = System.Collections.Generic.Dictionary<AST.Address, object>;
using Formula2VectDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Range>>;
using Vect2FormulaDict = System.Collections.Generic.Dictionary<AST.Range, System.Collections.Generic.HashSet<AST.Address>>;
using Vect2InputCellDict = System.Collections.Generic.Dictionary<AST.Range, System.Collections.Generic.HashSet<AST.Address>>;
using InputCell2VectDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Range>>;
using Formula2InputCellDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Address>>;
using InputCell2FormulaDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Address>>;
using AddrExpansion = System.Tuple<AST.Address, AST.Range[], AST.Address[]>;
using PathTuple = System.Tuple<string, string, string>;
using PathIndexDict = System.Collections.Generic.Dictionary<System.Tuple<string, string, string>, int>;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using AddrFunList = Microsoft.FSharp.Collections.FSharpList<AST.Address>;

namespace Depends
{
    [Serializable]
    public class DAG
    {
        public static int THIS_VERSION = 7;
        [OptionalField]
        private int _version = THIS_VERSION;
        private DateTime _dagBuilt;
        private string _path;
        private string _wbname;
        private string[] _wsnames;
        private CellRefDict _all_cells = new CellRefDict();                 // maps every cell address (including formulas) to its COMRef
        private VectorRefDict _all_vectors = new VectorRefDict();           // maps every vector to its COMRef
        private FormulaDict _formulas = new FormulaDict();                  // maps every formula to its formula expr
        private InputDict _inputs = new InputDict();                        // maps every cell address (including formulas) to its data
        private Formula2VectDict _f2v = new Formula2VectDict();             // maps every formula to its input vectors
        private Vect2FormulaDict _v2f = new Vect2FormulaDict();             // maps every input vector to its formulas
        private Formula2InputCellDict _f2i = new Formula2InputCellDict();   // maps every formula to its single-cell inputs
        private Vect2InputCellDict _v2i = new Vect2InputCellDict();         // maps every input vector to its component input cells
        private InputCell2VectDict _i2v = new InputCell2VectDict();         // maps every component input cell to its vectors
        private InputCell2FormulaDict _i2f = new InputCell2FormulaDict();   // maps every single-cell input to its formulas
        private Dictionary<AST.Range, bool> _do_not_perturb = new Dictionary<AST.Range, bool>();    // vector perturbability
        private Dictionary<AST.Address, int> _weights = new Dictionary<AST.Address, int>();         // graph node weight
        private readonly long _analysis_time;                               // amount of time to run dependence analysis
        private PathTuple[] _path_closure;                                  // the set of paths referenced by formulas in this DAG
        private PathIndexDict _path_closure_index;                          // the index of a path in the ordered array of closed-over paths

        [OnDeserializing]
        private void SetVersionDefault(StreamingContext sc)
        {
            _version = 1;
        }

        private static string SerializationPath(string dirpath, string wbname)
        {
            string[] paths = { dirpath, "EXCELINT_" + wbname + "." + THIS_VERSION + ".bin" };
            return Path.Combine(paths);
        }

        private class SingleRefDiff
        {
            HashSet<AST.Address> _same;
            HashSet<AST.Address> _deleted;
            HashSet<AST.Address> _added;

            public SingleRefDiff(HashSet<AST.Address> same, HashSet<AST.Address> deleted, HashSet<AST.Address> added)
            {
                _same = same;
                _deleted = deleted;
                _added = added;
            }

            public HashSet<AST.Address> Same
            {
                get
                {
                    return _same;
                }
            }

            public HashSet<AST.Address> Deleted
            {
                get
                {
                    return _deleted;
                }
            }

            public HashSet<AST.Address> Added
            {
                get
                {
                    return _added;
                }
            }
        }

        private HashSet<T> Union<T>(HashSet<T> hs1, HashSet<T> hs2)
        {
            var hsNew = new HashSet<T>(hs1);
            hsNew.UnionWith(hs2);
            return hsNew;
        }

        private HashSet<T> Intersect<T>(HashSet<T> hs1, HashSet<T> hs2)
        {
            var hsNew = new HashSet<T>(hs1);
            hsNew.IntersectWith(hs2);
            return hsNew;
        }

        private HashSet<T> Difference<T>(HashSet<T> hs1, HashSet<T> hs2)
        {
            var hsNew = new HashSet<T>(hs1);
            hsNew.ExceptWith(hs2);
            return hsNew;
        }

        private SingleRefDiff singleRefDiff(AST.Address addr, string fOld, string fNew)
        {
            var ssOld = new HashSet<AST.Address>(Parcel.addrReferencesFromFormula(fOld, addr.Path, addr.WorkbookName, addr.WorksheetName, false));
            var ssNew = new HashSet<AST.Address>(Parcel.addrReferencesFromFormula(fNew, addr.Path, addr.WorkbookName, addr.WorksheetName, false));

            var same = Union(ssOld, ssNew);
            var deleted = Difference(ssOld, ssNew);
            var added = Difference(ssNew, ssOld);

            return new SingleRefDiff(same, deleted, added);
        }

        public DAG CopyWithUpdatedFormulas(KeyValuePair<AST.Address,string>[] formulas, Microsoft.Office.Interop.Excel.Application app, bool ignore_parse_errors, Progress p)
        {
            var dag2 = new DAG(this);

            // copy build time
            dag2._dagBuilt = _dagBuilt;

            // clear graph
            dag2._inputs.Clear();
            dag2._all_vectors.Clear();
            dag2._do_not_perturb.Clear();
            dag2._f2v.Clear();
            dag2._v2f.Clear();
            dag2._i2v.Clear();
            dag2._v2i.Clear();
            dag2._i2f.Clear();
            dag2._f2i.Clear();

            // get all of the open workbooks & worksheets
            var openWBNames = new Dictionary<string, Microsoft.Office.Interop.Excel.Workbook>();
            var openWSNames = new Dictionary<Tuple<string,string>, Microsoft.Office.Interop.Excel.Worksheet>();
            foreach (Microsoft.Office.Interop.Excel.Workbook wb in app.Workbooks)
            {
                openWBNames.Add(wb.Name, wb);
                foreach (Microsoft.Office.Interop.Excel.Worksheet ws in wb.Worksheets)
                {
                    openWSNames.Add(new Tuple<string,string>(wb.Name, ws.Name), ws);
                }
            }

            // replace old formulas with new ones
            foreach (var newfrm in formulas)
            {
                var addr = newfrm.Key;
                var frm = newfrm.Value;
                var x = addr.Col;
                var y = addr.Row;
                var wb = openWBNames[addr.WorkbookName];
                var ws = openWSNames[new Tuple<string, string>(addr.WorkbookName, addr.WorksheetName)];
                var cell = this.getCOMRefForAddress(addr).Range;

                // update DAG formula string
                dag2._formulas[addr] = frm;

                // make a new COMRef
                var kvp2 = makeCOMRef(
                    y,
                    x,
                    addr.WorksheetName,
                    addr.WorkbookName,
                    addr.Path,
                    wb,
                    ws,
                    cell,
                    dag2._formulas);

                // add formula COMRef to cells
                dag2._all_cells[addr] = kvp2.Value;

                // get induced reference diffs
                var diff = singleRefDiff(addr, _formulas[addr], frm);

                // for all deleted references, remove links
                foreach (AST.Address dAddr in diff.Deleted)
                {
                    // remove edge from formula node to input node
                    dag2._f2i[addr].Remove(dAddr);
                    // remove edge from input node to formula node
                    dag2._i2f[dAddr].Remove(addr);
                }

                // for all added references, add links
                foreach (AST.Address aAddr in diff.Deleted)
                {
                    // add edge from formula node to input node
                    dag2._f2i[addr].Add(aAddr);
                    // add edge from input node to formula node
                    dag2._i2f[aAddr].Add(addr);
                }
            }

            foreach (var cell_cr in dag2._all_cells.KeysU)
            {
                if (cell_cr.IsFormula)
                {
                    // reinitialize references
                    var addr = dag2._all_cells[cell_cr];
                    dag2._f2v.Add(addr, new HashSet<AST.Range>());
                    dag2._f2i.Add(addr, new HashSet<AST.Address>());
                }
            }

            // parse formulas and rebuild graph
            ConstructDAG(app, dag2, ignore_parse_errors, p);

            return dag2;
        }

        public void SerializeToDirectory(string dirpath)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(SerializationPath(dirpath, _wbname), FileMode.Create, FileAccess.Write, FileShare.None);
            formatter.Serialize(stream, this);
            stream.Close();
        }

        public static bool CachedDAGExists(string cacheDirPath, string workbookName)
        {
            var fileName = SerializationPath(cacheDirPath, workbookName);
            return File.Exists(fileName);
        }

        public static DAG DAGFromCache(Boolean forceDAGBuild, Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Application app, bool ignore_parse_errors, string cacheDirPath, Progress p, CancellationToken t)
        {
            // get path
            var fileName = SerializationPath(cacheDirPath, wb.Name);

            // return DAG from cache path, otherwise build and serialize to cache path
            if (!forceDAGBuild && CachedDAGExists(cacheDirPath, wb.Name))
            {
                var dag = DeserializeFrom(fileName, app, p, t);

                if (t.IsCancellationRequested)
                {
                    return null;
                }

                if (dag._version != THIS_VERSION)
                {
                    p.Reset();
                    dag = newDAG(wb, app, ignore_parse_errors, cacheDirPath, p, DateTime.Now);
                }

                return dag;
            } else
            {
                return newDAG(wb, app, ignore_parse_errors, cacheDirPath, p, DateTime.Now);
            }
        }

        private static DAG newDAG(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Application app, bool ignore_parse_errors, string cacheDirPath, Progress p, DateTime dagBuilt)
        {
            var dag = new DAG(wb, app, ignore_parse_errors, p, dagBuilt);
            dag.SerializeToDirectory(cacheDirPath);
            return dag;
        }

        private static void reconstituteAddressRefs(DAG dag, Microsoft.Office.Interop.Excel.Application app, Progress p, CancellationToken t)
        {
            var allAddrs = dag.allCells();

            for (int i = 0; i < allAddrs.Length; i++)
            {
                if (t.IsCancellationRequested)
                {
                    return;
                }

                AST.Address addr = allAddrs[i];
                ParcelCOMShim.COMRef oldCR = dag._all_cells[addr];
                ParcelCOMShim.COMRef newCR = oldCR.DeserializationCellFixup(addr, app);
                dag._all_cells[addr] = newCR;

                p.IncrementCounter();
            }
        }

        private static void reconstituteRangeRefs(DAG dag, Microsoft.Office.Interop.Excel.Application app, Progress p, CancellationToken t)
        {
            var allVectors = dag.allVectors();

            for (int i = 0; i < allVectors.Length; i++)
            {
                if (t.IsCancellationRequested)
                {
                    return;
                }

                AST.Range rng = allVectors[i];
                ParcelCOMShim.COMRef oldCR = dag._all_vectors[rng];
                ParcelCOMShim.COMRef newCR = oldCR.DeserializationRangeFixup(rng, app);
                dag._all_vectors[rng] = newCR;

                p.IncrementCounter();
            }
        }

        public static DAG DeserializeFrom(string fileName, Microsoft.Office.Interop.Excel.Application app, Progress p, CancellationToken t)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            DAG obj = (DAG)formatter.Deserialize(stream);
            stream.Close();

            reconstituteAddressRefs(obj, app, p, t);
            reconstituteRangeRefs(obj, app, p, t);

            if (t.IsCancellationRequested)
            {
                return null;
            }

            return obj;
        }

        // for callers who do not need progress bars
        public DAG(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Application app, bool ignore_parse_errors, DateTime dagBuilt)
            : this(wb, app, ignore_parse_errors, new Progress(n => { }, () => { }, 1L), dagBuilt) { }

        public DAG(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Application app, bool ignore_parse_errors, Progress p, DateTime dagBuilt)
            : this(null, wb, app, ignore_parse_errors, p, dagBuilt) { }

        public DAG(Microsoft.Office.Interop.Excel.Worksheet ws, Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Application app, bool ignore_parse_errors, Progress p, DateTime dagBuilt)
        {
            // start stopwatch
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            // set build time
            _dagBuilt = dagBuilt;

            // save application & workbook references
            _path = wb.Path;
            _wbname = wb.Name;
            _wsnames = new string[wb.Worksheets.Count];
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                _wsnames[i - 1] = ((Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[i]).Name;
            }

            // bulk read worksheets & set progress total
            var data = FastFormulaRead(ws, wb);
            _formulas = data.formulas;
            _inputs = data.inputs;
            _f2v = data.f2v;
            _f2i = data.f2i;
            _all_cells = data.allCells;
            p.TotalWorkUnits = data.formulas.Count();

            // construct DAG
            ConstructDAG(app, this, ignore_parse_errors, p);

            // stop stopwatch
            sw.Stop();
            _analysis_time = sw.ElapsedMilliseconds;
        }

        // copy constructor
        public DAG(DAG dag)
        {
            _dagBuilt = dag._dagBuilt;
            _path = dag._path;
            _wbname = dag._wbname;
            _wsnames = dag._wsnames;
            _all_cells = new CellRefDict(dag._all_cells);
            _all_vectors = new VectorRefDict(dag._all_vectors);
            _formulas = new FormulaDict(dag._formulas);
            _inputs = new InputDict(dag._inputs);
            _f2v = new Formula2VectDict(dag._f2v);
            _v2f = new Vect2FormulaDict(dag._v2f);
            _f2i = new Formula2InputCellDict(dag._f2i);
            _v2i = new Vect2FormulaDict(dag._v2i);
            _i2v = new Formula2VectDict(dag._i2v);
            _i2f = new Formula2InputCellDict(dag._i2f);
            _do_not_perturb = new Dictionary<AST.Range, bool>(dag._do_not_perturb);
            _weights = new Dictionary<AST.Address, int>(dag._weights);
            _analysis_time = 0L;
        }

        private Boolean NeedsWorkbookOpen(AST.Range r, HashSet<string> openWBNames)
        {
            foreach (Tuple<AST.Address,AST.Address> tlbr in r.Ranges())
            {
                var tl = tlbr.Item1;
                var br = tlbr.Item2;
                if (NeedsWorkbookOpen(tl, openWBNames) || NeedsWorkbookOpen(tl, openWBNames))
                {
                    return true;
                }
            }
            return false;
        }

        private Boolean NeedsWorkbookOpen(AST.Address a, HashSet<string> openWBNames)
        {
            var result = !openWBNames.Contains(a.WorkbookName);

            return result;
        }

        public AST.Expression getASTofFormulaAt(AST.Address addr)
        {
            return Parcel.parseFormulaAtAddress(addr, this.getFormulaAtAddress(addr));
        }

        /// <summary>
        /// Finds all paths between all vertices in the DAG.  It is strongly advised
        /// that you supply a DAG as input or be prepared to wait awhile for the
        /// answer.
        /// </summary>
        /// <param name="dag">A directed acyclic graph.</param>
        /// <param name="p">Progress object (not presently used)</param>
        /// <returns></returns>
        private static SparseMatrix AllSimplePaths(DAG dag, Progress p)
        {
            var m = new SparseMatrix(dag.allCells().Length);

            Action<AST.Address, AddrFunList> dfs = null;
            dfs = (AST.Address start, AddrFunList antecedents) =>
            {
                if (dag.isFormula(start))
                {
                    // the distance between the start and itself is 0 by definition
                    m.Connect(start, start, 0);

                    var single_cells = dag._f2i[start];
                    var vector_cells = dag._f2v[start].SelectMany((v) => v.Addresses());
                    var all = single_cells.Concat(vector_cells);

                    var antecedents2 = AddrFunList.Cons(start, antecedents);

                    foreach (var child in all)
                    {
                        AST.Address[] ants = antecedents2.ToArray();

                        for (int depth = 0; depth < ants.Length; depth++)
                        {
                            m.Connect(ants[depth], child, depth + 1);
                        }
                        
                        dfs(child, antecedents2);
                    }
                }
            };

            foreach (var f in dag.terminalFormulaNodes(true))
            {
                dfs(f, AddrFunList.Empty);
            }

            return m;
        }

        public static void ConstructDAG(Microsoft.Office.Interop.Excel.Application app, DAG dag, bool ignore_parse_errors, Progress p)
        {
            // run the parser
            var frms = dag.getAllFormulaAddrs();
            var aes = new AddrExpansion[frms.Length];
            System.Threading.Tasks.Parallel.For(0, frms.Length, i =>
            {
                var formula_addr = frms[i];
                var cr = dag.getCOMRefForAddress(formula_addr);
                var vs = Parcel.rangeReferencesFromFormula(cr.Formula, cr.Path, cr.WorkbookName, cr.WorksheetName, ignore_parse_errors);
                var ss = Parcel.addrReferencesFromFormula(cr.Formula, cr.Path, cr.WorkbookName, cr.WorksheetName, ignore_parse_errors);

                aes[i] = new AddrExpansion(formula_addr, vs, ss);
            });

            // get all of the open workbooks
            var openWBNames = new HashSet<string>();
            foreach (Microsoft.Office.Interop.Excel.Workbook wb in app.Workbooks)
            {
                openWBNames.Add(wb.Name);
            }

            // do all the side-effecting stuff (building the graph) last
            foreach (AddrExpansion ae in aes)
            {
                var formula_addr = ae.Item1;
                var vectorRefs = ae.Item2;
                var scalarRefs = ae.Item3;

                foreach (AST.Range vector_rng in vectorRefs)
                {
                    // fetch/create COMRef, as appropriate
                    dag.makeInputVectorCOMRef(vector_rng, app, openWBNames);

                    // link formula and input vector
                    dag.linkInputVector(formula_addr, vector_rng);

                    // link input vector to the vector's single inputs
                    foreach (AST.Address input_single in vector_rng.Addresses())
                    {
                        dag.linkComponentInputCell(vector_rng, input_single);
                    }

                    // if num single inputs = num formulas,
                    // mark vector as non-perturbable
                    dag.markPerturbability(vector_rng);
                    
                }

                foreach (AST.Address input_addr in scalarRefs)
                {
                    dag.linkSingleCellInput(formula_addr, input_addr);
                }
            }
        }

        // this is mostly for diagnostic purposes
        public int numberOfInputCells()
        {
            var v_cells = new HashSet<AST.Address>(_all_vectors.KeysT.SelectMany(rng => rng.Addresses()));
            var sc_cells = new HashSet<AST.Address>(_i2f.Values.SelectMany(addr => addr));
            var all = v_cells.Union(sc_cells);
            return all.Count();
        }

        public struct RawGraph
        {
            public FormulaDict formulas;
            public InputDict inputs;
            public Formula2VectDict f2v;
            public Formula2InputCellDict f2i;
            public CellRefDict allCells;

            public RawGraph(bool yeah)
            {
                formulas = new FormulaDict();
                inputs = new InputDict();
                f2v = new Formula2VectDict();
                f2i = new Formula2InputCellDict();
                allCells = new CellRefDict();
            }
        }

        private struct DataAt<T>
        {
            public int Row;     // row
            public int Column;  // column
            public T Data;      // data

            public DataAt(int row, int column, T data)
            {
                Row = row;
                Column = column;
                Data = data;
            }
        }

        private static List<DataAt<string>> ReadFormulaStringList(Microsoft.Office.Interop.Excel.Range urng)
        {
            // init R1C1 extractor
            var regex = new Regex("^R([0-9]+)C([0-9]+)$", RegexOptions.Compiled);

            // init formula validator
            var fn_filter = new Regex("^=", RegexOptions.Compiled);

            // get dimensions
            var left = urng.Column;                      // 1-based left-hand y coordinate
            var right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
            var top = urng.Row;                          // 1-based top x coordinate
            var bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

            // init
            int width = right - left + 1;
            int height = bottom - top + 1;

            // output
            var fList = new List<DataAt<string>>();

            // if the used range is a single cell, Excel changes the type
            if (left == right && top == bottom)
            {
                var f = (string)urng.Formula;

                if (fn_filter.IsMatch(f))
                {
                    fList.Add(new DataAt<string>(top, left, f));
                }
            }
            else
            {
                // array read of formula cells
                // note that this is a 1-based 2D multiarray
                object[,] formulas = (object[,])urng.Formula;

                // for every cell that is actually a formula, add to 
                // formula dictionary & init formula lookup dictionaries
                for (int c = 1; c <= width; c++)
                {
                    for (int r = 1; r <= height; r++)
                    {
                        var f = (string)formulas[r, c];
                        if (fn_filter.IsMatch(f) && IsReallyAFormula(f, urng))
                        {
                            fList.Add(new DataAt<string>(r + top - 1, c + left - 1, f));
                        }
                    }
                }
            }
            return fList;
        }

        private static bool IsReallyAFormula(string formula, Microsoft.Office.Interop.Excel.Range used_range)
        {
            var wb = WorkbookFromRange(used_range);
            var ast_opt = Parcel.parseFormula(formula, wb.Path, wb.Name, used_range.Worksheet.Name);
            return Microsoft.FSharp.Core.FSharpOption<AST.Expression>.get_IsSome(ast_opt);
        }

        private static Microsoft.Office.Interop.Excel.Workbook WorkbookFromRange(Microsoft.Office.Interop.Excel.Range r)
        {
            return (Microsoft.Office.Interop.Excel.Workbook)r.Worksheet.Parent;
        }

        private static List<DataAt<Microsoft.Office.Interop.Excel.Range>> ReadCOMRefList(Microsoft.Office.Interop.Excel.Range urng)
        {
            // get dimensions
            var left = urng.Column;                      // 1-based left-hand y coordinate
            var right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
            var top = urng.Row;                          // 1-based top x coordinate
            var bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

            // init
            int width = right - left + 1;
            int height = bottom - top + 1;

            // output
            var rList = new List<DataAt<Microsoft.Office.Interop.Excel.Range>>();

            // array read of data cells
            // note that this is a 1-based 2D multiarray
            // we grab this in array form so that we can avoid a COM
            // call for every blank-cell check
            object[,] data;
            // annoyingly, the return type for Value2 changes depending on the size of the range
            if (width == 1 && height == 1)
            {
                int[] lengths = new int[2] { 1, 1 };
                int[] lower_bounds = new int[2] { 1, 1 };
                data = (object[,])Array.CreateInstance(typeof(object), lengths, lower_bounds);
                data[1, 1] = urng.Value2;
            }
            else
            {   // ok, it really is an array
                data = (object[,])urng.Value2;
            }

            // if the worksheet contains nothing, data will be null
            if (data != null)
            {
                // for each COM object in the used range, create an address object
                // WITHOUT calling any methods on the COM object itself
                int x_old = -1;
                int x = -1;
                int y = 0;

                for (int i = 0; i < width * height; i++)
                {
                    // The basic idea here is that we know how Excel iterates over collections
                    // of cells.  The Excel.Range returned by UsedRange is always rectangular.
                    // Thus we can calculate the addresses of each COM cell reference without
                    // needing to incur the overhead of actually asking it for its address.
                    x = (x + 1) % width;
                    // increment y if x wrapped (x < x_old or x == x_old when width == 1)
                    y = x <= x_old ? y + 1 : y;

                    int c = x + left;
                    int r = y + top;

                    // don't track if the cell contains nothing
                    if (data[y + 1, x + 1] != null) // adjust indices to be one-based
                    {
                        rList.Add(new DataAt<Microsoft.Office.Interop.Excel.Range>(r, c, (Microsoft.Office.Interop.Excel.Range)urng.Item[r, c]));
                    }

                    x_old = x;
                }
            }

            return rList;
        }

        private static List<DataAt<object>> ReadInputList(Microsoft.Office.Interop.Excel.Range urng)
        {
            // get dimensions
            var left = urng.Column;                      // 1-based left-hand y coordinate
            var right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
            var top = urng.Row;                          // 1-based top x coordinate
            var bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

            // init
            int width = right - left + 1;
            int height = bottom - top + 1;

            // output
            var dList = new List<DataAt<object>>();

            // array read of data cells
            // note that this is a 1-based 2D multiarray
            // we grab this in array form so that we can avoid a COM
            // call for every blank-cell check
            object[,] data;
            // annoyingly, the return type for Value2 changes depending on the size of the range
            if (width == 1 && height == 1)
            {
                int[] lengths = new int[2] { 1, 1 };
                int[] lower_bounds = new int[2] { 1, 1 };
                data = (object[,])Array.CreateInstance(typeof(object), lengths, lower_bounds);
                data[1,1] = urng.Value2;
            } else
            {   // ok, it really is an array
                data = (object[,])urng.Value2;
            }

            // if the worksheet contains nothing, data will be null
            if (data != null)
            {
                // for each COM object in the used range, create an address object
                // WITHOUT calling any methods on the COM object itself
                int x_old = -1;
                int x = -1;
                int y = 0;

                for (int i = 0; i < width * height; i++)
                {
                    // The basic idea here is that we know how Excel iterates over collections
                    // of cells.  The Excel.Range returned by UsedRange is always rectangular.
                    // Thus we can calculate the addresses of each COM cell reference without
                    // needing to incur the overhead of actually asking it for its address.
                    x = (x + 1) % width;
                    // increment y if x wrapped (x < x_old or x == x_old when width == 1)
                    y = x <= x_old ? y + 1 : y;

                    int c = x + left;
                    int r = y + top;

                    // don't track if the cell contains nothing
                    if (data[y + 1, x + 1] != null) // adjust indices to be one-based
                    {
                        dList.Add(new DataAt<object>(r, c, data[y + 1, x + 1]));
                    }

                    x_old = x;
                }
            }

            return dList;
        }

        public bool Changed(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            // get names once
            var wbfullname = wb.FullName;
            var wbname = wb.Name;
            var path = wb.Path;

            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in wb.Worksheets)
            {
                // get name once
                var wsname = worksheet.Name;

                // get used range
                Microsoft.Office.Interop.Excel.Range urng = worksheet.UsedRange;

                // get formulas
                var formulas = ReadFormulaStringList(urng);

                // short-circuit for formula additions and removals
                var fcnt = _formulas.Where(pair => pair.Key.A1Worksheet() == wsname).Count();
                if (formulas.Count != fcnt)
                {
                    return true;
                }

                // get cell contents
                var inputs = ReadInputList(urng);

                // short-circuit for input additions and removals
                var icnt = _all_cells.WhereT(addr => addr.A1Worksheet() == wsname).Count();
                if (inputs.Count != icnt)
                {
                    return true;
                }

                // check formulas
                foreach (var formula in formulas)
                {
                    var addr = AST.Address.fromR1C1withMode(formula.Row, formula.Column, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
                    if (!_formulas.ContainsKey(addr) || _formulas[addr] != formula.Data)
                    {
                        return true;
                    }
                }

                // check data
                foreach (var input in inputs)
                {
                    var addr = AST.Address.fromR1C1withMode(input.Row, input.Column, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
                    object data = _inputs[addr];
                    
                    // There are only two possibilities: data is a value type or data is a reference type
                    // If data is a value type, then input must also be.
                    // If data is not a value type, then it must be a string, and since both
                    // will be refs to different locations, conversion and comparison is necessary

                    // can't be the same value if they're not even
                    // the same data type
                    if (data.GetType() != input.Data.GetType())
                    {
                        return true;
                    }
                    // double
                    else if (data.GetType() == typeof(double))
                    {
                        if ((double)data != (double)input.Data)
                        {
                            return true;
                        }
                    // int
                    } else if (data.GetType() == typeof(int))
                    {
                        if ((int)data != (int)input.Data)
                        {
                            return true;
                        }
                    // string
                    } else
                    {
                        string d = Convert.ToString(data);
                        string i = Convert.ToString(input.Data);
                        if (d.CompareTo(i) != 0)
                        {
                            // strings differ
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private class AddrCache
        {
            Dictionary<Tuple<int, int>, AST.Address> addrs;

            public AddrCache(int initSz)
            {
                addrs = new Dictionary<Tuple<int, int>, AST.Address>(initSz);
            }

            public AST.Address getAddr(int row, int col, string wsname, string wbname, string path)
            {
                var rc = new Tuple<int, int>(row, col);
                AST.Address addr;
                if (addrs.ContainsKey(rc))
                {
                    addr = addrs[rc];
                } else
                {
                    addr = AST.Address.fromR1C1withMode(row, col, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
                    addrs.Add(rc, addr);
                }
                return addr;
            }
        }

        public static RawGraph FastFormulaRead(Microsoft.Office.Interop.Excel.Worksheet ws, Microsoft.Office.Interop.Excel.Workbook wb)
        {
            // allocate struct, etc.
            var retVal = new RawGraph(true);

            // get names once
            var wbfullname = wb.FullName;
            var wbname = wb.Name;
            var path = wb.Path;

            if (ws != null)
            {
                FastFormulaReadWorksheet(ws, wb, retVal, wbname, path);
            } else
            {
                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in wb.Worksheets)
                {
                    FastFormulaReadWorksheet(worksheet, wb, retVal, wbname, path);
                }
            }

            return retVal;
        }

        public static void FastFormulaReadWorksheet(
                            Microsoft.Office.Interop.Excel.Worksheet worksheet,
                            Microsoft.Office.Interop.Excel.Workbook wb,
                            RawGraph retVal,
                            string wbname,
                            string path
                           )
        {
            // get name once
            var wsname = worksheet.Name;

            // get used range
            Microsoft.Office.Interop.Excel.Range urng = worksheet.UsedRange;

            // make address object cache
            var addrCache = new AddrCache(urng.Count);

            // get formulas
            var formulas = ReadFormulaStringList(urng);

            // get data
            var inputs = ReadInputList(urng);

            // get COM refs
            var refs = ReadCOMRefList(urng);

            // process formulas
            foreach (var formula in formulas)
            {
                var addr = addrCache.getAddr(formula.Row, formula.Column, wsname, wbname, path);
                retVal.formulas.Add(addr, formula.Data);
                retVal.f2v.Add(addr, new HashSet<AST.Range>());
                retVal.f2i.Add(addr, new HashSet<AST.Address>());
            }

            // process data
            foreach (var input in inputs)
            {
                var addr = addrCache.getAddr(input.Row, input.Column, wsname, wbname, path);
                retVal.inputs.Add(addr, input.Data);
            }

            // process COM refs
            foreach (var drng in refs)
            {
                var addr = addrCache.getAddr(drng.Row, drng.Column, wsname, wbname, path);
                var formula = retVal.formulas.ContainsKey(addr) ? new Microsoft.FSharp.Core.FSharpOption<string>(retVal.formulas[addr]) : Microsoft.FSharp.Core.FSharpOption<string>.None;
                var cr = new ParcelCOMShim.LocalCOMRef(wb, worksheet, drng.Data, path, wbname, wsname, formula, 1, 1);
                retVal.allCells.Add(addr, cr);
            }
        }

        // This seriously ugly method exists because we need to call it from several places,
        // one of which is very hot.  Computing many of these parameters from COM objects
        // is expensive, so we expand them into the parameter list.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static KeyValuePair<AST.Address, ParcelCOMShim.LocalCOMRef> makeCOMRef(int r, int c, string wsname, string wbname, string path, Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Worksheet ws, Microsoft.Office.Interop.Excel.Range cell, Dictionary<AST.Address, string> formulas)
        {
            var addr = AST.Address.fromR1C1withMode(r, c, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var formula = formulas.ContainsKey(addr) ? new Microsoft.FSharp.Core.FSharpOption<string>(formulas[addr]) : Microsoft.FSharp.Core.FSharpOption<string>.None;
            var cr = new ParcelCOMShim.LocalCOMRef(wb, ws, cell, path, wbname, wsname, formula, 1, 1);

            return new KeyValuePair<AST.Address, ParcelCOMShim.LocalCOMRef>(addr, cr);
        }

        public string readCOMValueAtAddress(AST.Address addr)
        {
            // null values become the empty string
            var s = System.Convert.ToString(this.getCOMRefForAddress(addr).Range.Value2);
            if (s == null)
            {
                return "";
            }
            else
            {
                return s;
            }
        }

        public long AnalysisMilliseconds
        {
            get { return _analysis_time; }
        }

        public ParcelCOMShim.COMRef getCOMRefForAddress(AST.Address addr)
        {
            return _all_cells[addr];
        }

        public ParcelCOMShim.COMRef getCOMRefForRange(AST.Range rng)
        {
            return _all_vectors[rng];
        }

        public string getFormulaAtAddress(AST.Address addr)
        {
            return _formulas[addr];
        }

        public AST.Address[] getAllFormulaAddrs()
        {
            return _formulas.Keys.ToArray();
        }

        public void makeInputVectorCOMRef(AST.Range rng, Microsoft.Office.Interop.Excel.Application app, HashSet<string> openWBNames)
        {
            // check for the range in the dictionary
            ParcelCOMShim.COMRef c;

            // if it's not in the dict, create it
            if (!_all_vectors.TryGetValue(rng, out c))
            {
                // is it a local reference?
                if (NeedsWorkbookOpen(rng, openWBNames))
                {
                    // no
                    string path = rng.GetPathName();
                    string wbname = rng.GetWorkbookName();
                    string wsname = rng.GetWorksheetName();

                    c = new ParcelCOMShim.NonLocalComRef(path, wbname, wsname, Microsoft.FSharp.Core.FSharpOption<string>.None);
                } else
                {
                    // yes
                    Microsoft.Office.Interop.Excel.Range com = ParcelCOMShim.Range.GetCOMObject(rng, app);
                    Microsoft.Office.Interop.Excel.Worksheet ws = com.Worksheet;
                    Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)ws.Parent;
                    string wsname = ws.Name;
                    string wbname = wb.Name;
                    var path = wb.Path;
                    int width = com.Columns.Count;
                    int height = com.Rows.Count;
                    c = new ParcelCOMShim.LocalCOMRef(wb, ws, com, path, wbname, wsname, Microsoft.FSharp.Core.FSharpOption<string>.None, width, height);
                }

                // cache it
                _all_vectors.Add(rng, c);
                _do_not_perturb.Add(rng, true);    // initially mark as not perturbable
            }
        }

        public void linkInputVector(AST.Address formula_addr, AST.Range vector_rng)
        {
            // add range to range-lookup-by-formula_addr dictionary
            // (initialized in DAG constructor)
            _f2v[formula_addr].Add(vector_rng);
            // add formula_addr to faddr-lookup-by-range dictionary,
            // initializing bucket if necessary
            if (!_v2f.ContainsKey(vector_rng))
            {
                _v2f.Add(vector_rng, new HashSet<AST.Address>());
            }
            if (!_v2f[vector_rng].Contains(formula_addr))
            {
                _v2f[vector_rng].Add(formula_addr);
            }
        }

        public void linkComponentInputCell(AST.Range input_range, AST.Address input_addr)
        {
            // add input_addr to iaddr-lookup-by-input_range dictionary,
            // initializing bucket if necessary
            if (!_v2i.ContainsKey(input_range))
            {
                _v2i.Add(input_range, new HashSet<AST.Address>());
            }
            if (!_v2i[input_range].Contains(input_addr))
            {
                _v2i[input_range].Add(input_addr);
            }
            // add input_range to irng-lookup-by-iaddr dictionary,
            // initializing bucket if necessary
            if (!_i2v.ContainsKey(input_addr))
            {
                _i2v.Add(input_addr, new HashSet<AST.Range>());
            }
            if (!_i2v[input_addr].Contains(input_range))
            {
                _i2v[input_addr].Add(input_range);
            }
        }

        public void linkSingleCellInput(AST.Address formula_addr, AST.Address input_addr)
        {
            // add address to input_addr-lookup-by-formula_addr dictionary
            // (initialzied in DAG constructor)
            _f2i[formula_addr].Add(input_addr);
            
            // add formula_addr to faddr-lookup-by-iaddr dictionary,
            // initializing bucket if necessary
            if (!_i2f.ContainsKey(input_addr))
            {
                _i2f.Add(input_addr, new HashSet<AST.Address>());
            }
            if (!_i2f[input_addr].Contains(formula_addr))
            {
                _i2f[input_addr].Add(formula_addr);
            }
        }

        public void markPerturbability(AST.Range vector_rng)
        {
            // get inputs
            var inputs = _v2i[vector_rng];

            // count inputs that are formulas
            int fcnt = inputs.Count(iaddr => _formulas.ContainsKey(iaddr));

            // If there is at least one input that is not a formula
            // mark the whole vector as perturbable.
            // Note: all vectors marked as non-perturbable by default.
            if (fcnt != inputs.Count)
            {
                _do_not_perturb[vector_rng] = false;
            }
        }

        public bool containsLoop()
        {
            var OK = true;
            var visited_from = new Dictionary<AST.Address, AST.Address>();
            foreach (AST.Address addr in _formulas.Keys)
            {
                OK = OK && !traversalHasLoop(addr, visited_from, null);
            }
            return !OK;
        }

        private bool traversalHasLoop(AST.Address current_addr, Dictionary<AST.Address, AST.Address> visited, AST.Address from_addr)
        {
            // base case 1: loop check
            if (visited.ContainsKey(current_addr))
            {
                return true;
            }
            // base case 2: an input cell
            if (!_formulas.ContainsKey(current_addr))
            {
                return false;
            }
            // recursive case (it's a formula)
            // check both single inputs and the inputs of any vector inputs
            bool OK = true;
            HashSet<AST.Address> single_inputs = _f2i[current_addr];
            HashSet<AST.Address> vector_inputs = new HashSet<AST.Address>(_f2v[current_addr].SelectMany(addrs => addrs.Addresses()));
            foreach (AST.Address input_addr in vector_inputs.Union(single_inputs))
            {
                if (OK)
                {
                    // new dict to mark visit
                    var visited2 = new Dictionary<AST.Address, AST.Address>(visited);
                    // mark visit
                    visited2.Add(current_addr, from_addr);
                    // recurse
                    OK = OK && !traversalHasLoop(input_addr, visited2, from_addr);
                }
            }
            return !OK;
        }

        public string ToDOT()
        {
            var visited = new HashSet<AST.Address>();
            StringBuilder sb = new StringBuilder();
            sb.Append("digraph spreadsheet {\n");
            foreach (AST.Address formula_addr in _formulas.Keys)
            {
                ToDOT(formula_addr, visited, sb);
            }
            sb.Append("\n}\n");
            return sb.ToString();
        }

        private string DOTEscapedFormulaString(string formula)
        {
            return formula.Replace("\"", "\\\"");
        }

        private string DOTNodeName(AST.Address addr)
        {
            return "\"" + addr.A1Local() + "[" + (_formulas.ContainsKey(addr) ? DOTEscapedFormulaString(_formulas[addr]) : readCOMValueAtAddress(addr)) + "]\"";
        }

        private void ToDOT(AST.Address current_addr, HashSet<AST.Address> visited, StringBuilder sb)
        {
            // base case 1: loop protection
            if (visited.Contains(current_addr))
            {
                return;
            }
            // base case 2: an input
            if (!_formulas.ContainsKey(current_addr))
            {
                return;
            }
            // case 3: a formula

            var ca_name = DOTNodeName(current_addr);

            // 3a. single-cell input 
            HashSet<AST.Address> single_inputs = _f2i[current_addr];
            foreach (AST.Address input_addr in single_inputs)
            {
                var ia_name = DOTNodeName(input_addr);

                // print
                sb.Append(ia_name).Append(" -> ").Append(ca_name).Append(";\n");

                // mark visit
                visited.Add(input_addr);

                // recurse
                ToDOT(input_addr, visited, sb);
            }

            // 3b. vector input
            HashSet<AST.Range> vector_inputs = _f2v[current_addr];
            foreach (AST.Range v_addr in vector_inputs)
            {
                var rng_name = "\"" + v_addr.A1Local() + "\"";

                // print
                sb.Append(rng_name).Append(" -> ").Append(ca_name).Append(";\n");

                // recurse
                foreach (AST.Address input_addr in v_addr.Addresses())
                {
                    var ia_name = DOTNodeName(input_addr);

                    // print
                    sb.Append(ia_name).Append(" -> ").Append(rng_name).Append(";\n");

                    // mark visit
                    visited.Add(input_addr);

                    // recurse
                    ToDOT(input_addr, visited, sb);
                }
            }
        }

        /// <summary>
        /// Returns all formula addresses that are not referenced
        /// by any other formula, unless <paramref name="all_outputs"/>
        /// is true, in which case all known formula addresses are
        /// returned.
        /// </summary>
        /// <param name="all_outputs">If true, return all known formula addresses</param>
        /// <returns></returns>
        public AST.Address[] terminalFormulaNodes(bool all_outputs)
        {
            // return only the formula nodes that do not serve
            // as input to another cell and that are also not
            // in our list of excluded functions
            if (all_outputs)
            {
                return getAllFormulaAddrs();
            }
            else
            {
                // get all formula addresses
                return getAllFormulaAddrs().Where(addr =>
                    // such that the number of formulas consuming this formula == 0
                    (!_i2f.ContainsKey(addr) || _i2f[addr].Count == 0) &&
                    // and the number of vectors containing this formula == 0
                    (!_i2v.ContainsKey(addr) || _i2v[addr].Count == 0)
                ).ToArray();
            }
        }

        public void setWeight(AST.Address node, int weight)
        {
            if (!_weights.ContainsKey(node))
            {
                _weights.Add(node, weight);
            }
            else
            {
                _weights[node] = weight;
            }
        }

        public int getWeight(AST.Address node)
        {
            return _weights[node];
        }

        public HashSet<AST.Range> getFormulaInputVectors(AST.Address f)
        {
            // no need to check for key existence; empty
            // HashSet initialized in DAG constructor
            return _f2v[f];
        }

        public bool isFormula(AST.Address node)
        {
            return _formulas.ContainsKey(node);
        }

        public HashSet<AST.Address> getFormulaSingleCellInputs(AST.Address node)
        {
            // no need to check for key existence; empty
            // HashSet initialized in DAG constructor
            return _f2i[node];
        }

        public AST.Range[] terminalInputVectors()
        {
            return _do_not_perturb
                .Where(pair => !pair.Value)
                .Select(pair => pair.Key).ToArray();
        }

        public AST.Address[] allInputs()
        {
            // get all of the input ranges for all of the functions
            var inputs = _f2v.Values.SelectMany(rngs => rngs.SelectMany(rng => rng.Addresses())).Distinct();

            // get all of the single-cell inputs for all of the functions
            var scinputs = _f2i.Values.SelectMany(rngs => rngs).Distinct();

            // concat all together and return
            return inputs.Concat(scinputs).Distinct().ToArray();
        }

        public AST.Address[] allComputationCells()
        {
            // get all inputs
            var inputs = allInputs();

            // get all formulas
            var formulas = getAllFormulaAddrs();

            // concat all together and return
            return inputs.Concat(formulas).Distinct().ToArray();
        }

        public AST.Address[] terminalInputCells()
        {
            // this folds all of the inputs for all of the
            // outputs into a set of distinct data-containing cells
            var iecells = terminalFormulaNodes(true).Aggregate(
                              Enumerable.Empty<AST.Address>(),
                              (acc, node) => acc.Union<AST.Address>(getChildCellsRec(node))
                          );
            return iecells.ToArray<AST.Address>();
        }

        /// <summary>
        /// Gets all cells transitively referenced by the formula at cell_addr,
        /// both those cells referenced by single-cell references and cells referenced
        /// by vector references, including the formula itself.
        /// </summary>
        /// <param name="formula"></param>
        /// <returns>A sequence of addresses.</returns>
        private IEnumerable<AST.Address> getChildCellsRec(AST.Address formula)
        {
            // recursive case
            if (_formulas.ContainsKey(formula))
            {
                // recursively get vector inputs
                var vector_children = _f2v[formula].SelectMany(rng => getVectorChildCellsRec(rng));

                // recursively get single-cell inputs
                var sc_children = _f2i[formula].SelectMany(cell => getChildCellsRec(cell));

                return vector_children.Concat(sc_children);
                // base case
            }
            else
            {
                return new List<AST.Address> { formula };
            }
        }


        private IEnumerable<AST.Address> getVectorChildCellsRec(AST.Range vector_addr)
        {
            // get single-cell inputs (vectors only consist of single cells)
            return _v2i[vector_addr].SelectMany(rng => getChildCellsRec(rng));
        }

        public AST.Range[] allVectors()
        {
            return _all_vectors.KeysT.ToArray();
        }

        public AST.Address[] allCells()
        {
            return _all_cells.KeysT.ToArray();
        }

        public AST.Address[] getFormulasThatRefCell(AST.Address cell)
        {
            if (_i2f.ContainsKey(cell))
            {
                return _i2f[cell].ToArray();
            }
            else
            {
                return new AST.Address[] { };
            }
        }

        public AST.Address[] getFormulasThatRefVector(AST.Range rng)
        {
            if (_v2f.ContainsKey(rng))
            {
                return _v2f[rng].ToArray();
            }
            else
            {
                return new AST.Address[] { };
            }
        }

        public AST.Range[] getVectorsThatRefCell(AST.Address cell)
        {
            if (_i2v.ContainsKey(cell))
            {
                return _i2v[cell].ToArray();
            } else
            {
                return new AST.Range[] { };
            }
        }

        public string getWorkbookDirectory()
        {
            return _path;
        }

        public string getWorkbookName()
        {
            return _wbname;
        }

        public string getWorkbookPath()
        {
            string[] paths = { _path, _wbname };
            return System.IO.Path.Combine(paths);
        }

        public string[] getWorksheetNames()
        {
            return _wsnames;
        }

        // returns the set of all paths (directory, workbook, worksheet)
        // referenced by any formula in this DAG, lexicographically ordered.
        // we evaluate this lazily since it is not always needed
        public Tuple<string,string,string>[] getPathClosure()
        {
            if (_path_closure == null)
            {
                var paths = new HashSet<Tuple<string, string, string>>();

                // single-cell references
                foreach (HashSet<AST.Address> cells in _f2i.Values)
                {
                    foreach (AST.Address cell in cells)
                    {
                        var dir = cell.Path;
                        var wbname = cell.WorkbookName;
                        var wsname = cell.WorksheetName;
                        paths.Add(new Tuple<string, string, string>(dir, wbname, wsname));
                    }
                }

                // vector references
                foreach (HashSet<AST.Range> ranges in _f2v.Values)
                {
                    foreach (AST.Range range in ranges)
                    {
                        var dir = range.GetPathName();
                        var wbname = range.GetWorkbookName();
                        var wsname = range.GetWorksheetName();
                        paths.Add(new Tuple<string, string, string>(dir, wbname, wsname));
                    }
                }

                // all cells-- this covers paths to cells not referenced by formulas
                foreach (AST.Address cell in _all_cells.KeysT)
                {
                    var dir = cell.Path;
                    var wbname = cell.WorkbookName;
                    var wsname = cell.WorksheetName;
                    paths.Add(new Tuple<string, string, string>(dir, wbname, wsname));
                }

                _path_closure = paths.OrderBy(key => key.Item1 + key.Item2 + key.Item3).ToArray();
            }

            return _path_closure;
        }

        public int getPathClosureIndex(Tuple<string,string,string> path)
        {
            if (_path_closure_index == null)
            {
                var pci = new Dictionary<Tuple<string, string, string>,int>();
                var pc = getPathClosure();

                for (int i = 0; i < pc.Length; i++)
                {
                    pci.Add(pc[i], i);
                }

                _path_closure_index = pci;
            }

            return _path_closure_index[path];
        }

        public DateTime Built
        {
            get { return _dagBuilt; }
        }
    }
}
