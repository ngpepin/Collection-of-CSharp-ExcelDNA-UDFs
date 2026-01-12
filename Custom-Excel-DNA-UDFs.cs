/*
 *
 * ExcelDNA User-Defined Functions 
 *
 * This collection provides powerful worksheet functions that extend Excel's native capabilities. All functions are thread-safe and designed for high performance 
 * in large spreadsheets. Stateful functions (like INJECTVALUE) maintain state between calculations and as such violate Excel's "no side effects" rule 
 * (intentionally of course, because by doing so they allow state machines to be created in spreadsheet models!)
 *
 * VERCELDNA, SETTARGETVERSION, GETTARGETVERSION, RECALCALL, GETITERATIONSTATUS, SETITERATION, ISVISIBLE, DESCRIBE, INJECTVALUE, FINDPOS, 
 * PUTOBJECT, GETOBJECT, PURGEOBJECTS,TRUESPLIT, ISMEMBEROF, GETTHREADS, SETTHREADS, HASHARRAY, ISLOCALIP
 *
 * New in version 3.2.0 (needs documenting):
 * 
 * Summary of Functions:
 *
 * 1. VEXCELDNA()
 *    - Returns the current version of the UDF collection
 *    - Usage: =vExcelDNA()
 *    - Returns: String with the version number
 *
 * 2. SETTARGETVERSION(version)
 *    - Sets the target version for backward compatibility
 *    - Usage: =SetTargetVersion("2.0.0")
 *    - Returns: Confirmation string with the previous and current target version
 *
 * 3. GETTARGETVERSION()
 *    - Gets the current target version for backward compatibility
 *    - Usage: =GetTargetVersion()
 *    - Returns: String with the current target version
 *
 * 4. RECALCALL()
 *    - Triggers a full recalculation of the workbook
 *    - Usage: =RECALCALL()
 *    - Returns: "TRUE" on success
 *
 * 5. GETITERATIONSTATUS()
 *    - Returns Excel's iterative calculation settings
 *    - Usage: =GETITERATIONSTATUS()
 *    - Returns: String with status (ON/OFF), max iterations, and max change
 *
 * 6. SETITERATION(IterationOn, [maxIterations], [maxChange])
 *    - Configures Excel's iterative calculation settings
 *    - Usage: =SETITERATION(TRUE, 100, 0.001)
 *    - Returns: Confirmation string with current settings
 *
 * 7. ISVISIBLE([cachingTime])
 *    - Checks if a cell is visible (not hidden by rows/columns)
 *    - Usage: =ISVISIBLE(10)  (10 second cache duration)
 *    - Returns: "TRUE" if visible, "FALSE" if hidden
 *
 * 8. DESCRIBE(cell_reference)
 *    - Returns a description of the cell's content type
 *    - Usage: =DESCRIBE(A1)
 *    - Returns: String describing the value type
 *
 * 9. INJECTVALUE(cell_reference, value)
 *    - Injects a value into a cell (stateful operation)
 *    - Usage: =INJECTVALUE(B2, "Test Value")
 *    - Returns: The injected value
 *
 * 10.FINDPOS(text, substring, instance)
 *    - Finds positions of substrings (case-insensitive)
 *    - Usage: =FINDPOS("Hello World", "o", 1)
 *    - Returns: Position number or error if not found
 *
 * 11.PUTOBJECT(name, value, [force], [debug])
 *    - Stores an object in temporary storage
 *    - Usage: =PUTOBJECT("temp1", A1:A10, TRUE)
 *    - Returns: The stored object
 *
 * 12.GETOBJECT(name, [debug])
 *    - Retrieves an object from temporary storage
 *    - Usage: =GETOBJECT("temp1")
 *    - Returns: The stored object or error
 *
 * 13. PURGEOBJECTS()
 *     - Clears all objects from temporary storage
 *     - Usage: =PURGEOBJECTS()
 *     - Returns: "TRUE" on success
 *
 * 14. TRUESPLIT(input_array, delimiter)
 *     - Splits strings into dynamic arrays
 *     - Usage: =TRUESPLIT(A1:A3, ",")
 *     - Returns: 2D array of split components
 *
 * 15. ISMEMBEROF(array1, array2)
 *     - Checks for common elements between arrays
 *     - Usage: =ISMEMBEROF(A1:A10, B1:B20)
 *     - Returns: TRUE if any match found
 *
 * 16. GETTHREADS()
 *     - Returns Excel's current thread count for calculations
 *     - Usage: =GETTHREADS()
 *     - Returns: Integer thread count
 *
 * 17. SETTHREADS(threadCount)
 *     - Configures Excel's calculation thread count
 *     - Usage: =SETTHREADS(4)  (Use 4 threads)
 *              =SETTHREADS(0)  (Use all processors)
 *     - Returns: Actual thread count set
 *
 * 18. HASHARRAY(input_array, [hashLength])
 *     - Computes a consistent hash value for an array of values
 *     - Usage: =HASHARRAY(A1:A10, 8)
 *     - Returns: Hash string (default length 8, range 4-32)
 *
 * 19. ISLOCALIP(ipAddress_string)
 *    - Checks if an IP address is a local IP (private or loopback)
 *    - Usage: =ISLOCALIP(ipAddress_string)
 *    - Returns: TRUE if local IP, FALSE otherwise or #N/A if invalid input
 *
 * 20. ARRAYSUBTRACT(arrayA, arrayB)
 *    - Subtracts values in arrayB from arrayA, preserving the shape of arrayA where possible
 *    - Usage: =ARRAYSUBTRACT(A1:A10, B1:B3)
 *    - Returns: Dynamic array of values from arrayA that are not present in arrayB
 *
 * 21. EXTRACTSUBSTR(inputString, startMarker, [endMarker])
 *    - Extracts a substring between start and end markers
 *    - Usage: =EXTRACTSUBSTR("A=[123] Z", "A=[", "]")
 *    - Returns: The extracted substring or #N/A if markers are not found
 *
 * 22. STRING_COMMON(s1, s2, minLength)
 *    - Returns maximal common substrings with a minimum length
 *    - Usage: =STRING_COMMON("Hello there, how are you","Hello there how are you",5)
 *    - Returns: Dynamic array of common substrings (empty if none meet minLength)
 *
 * 23. STRING_DIFF(s1, s2, minLength)
 *    - Returns maximal differing substrings with a minimum length
 *    - Usage: =STRING_DIFF("Hello there, how are you","Hello there how are you",1)
 *    - Returns: Dynamic array of differing substrings from s1 (empty if none meet minLength)
 *
 * Notes:
 * - Functions marked as volatile recalculate when any cell changes
 * - Stateful functions (like INJECTVALUE) maintain state between calculations
 * - Temporary storage persists until PURGEOBJECTS() is called or workbook closes
 * - Thread management requires Excel 2007 or later
 *

 *  Notes re: GLOBAL VOLATILITY SWITCH 
 *
 *  Why this exists
 *  ---------------
 *  Many of the original Excel-DNA functions were decorated with [IsVolatile = true] (or called
 *  Application.Volatile in VBA).  On large models that can cause Excel to recalculate *every*
 *  instance of those functions whenever **any** cell changes, interrupting editing and killing
 *  performance.  
 *
 *  This file introduces an **opt-in, workbook-wide toggle** that lets the modeller decide:
 *      *  Volatility ON   -  behave exactly like the old code (recalc on every change)
 *      *  Volatility OFF  -  behave like ordinary non-volatile formulas (only recalc when an
 *                            argument changes or the user presses F9 / Shift+F9)
 *
 *  How it works
 *  ------------
 *
 *      New UDFs
 *      --------------------------------------------------------------------
 *      NAME              USAGE                              EFFECT
 *      --------------------------------------------------------------------
 *      SetVolatility()   =SetVolatility(TRUE/FALSE)         Enables or disables volatility
 *      GetVolatility()   =GetVolatility()                   Returns "ENABLED" / "DISABLED"
 *
 *      Helper inside the code
 *      internal static void MaybeVolatile()
 *         - Called at the top of any function that *used* to be marked volatile.
 *         - Internally calls `XlCall.Excel(xlfVolatile, true)` **only when** the global flag is ON.
 *
 *      Code change pattern
 *      OLD:   [ExcelFunction(IsVolatile = true)] public static object Foo(...) { ... }
 *      NEW:   [ExcelFunction] public static object Foo(...) { MaybeVolatile(); ... }
 *
 *  Behavioural impact
 *  ------------------
 *  * Built-in Excel volatile functions (NOW, RAND, OFFSET, INDIRECT, etc.) are **unaffected.**
 *  * Any *other* add-ins remain untouched unless they explicitly reference the same switch.
 *  * Pressing F9 or Shift+F9 still forces a recalculation of everything, as usual.
 *  * Models that relied on "tick-every-calculation" side-effects (e.g. INJECTVALUE driving a
 *    state machine) should either keep volatility ON or accept an explicit trigger argument.
 *
 *  Default state & persistence
 *  ---------------------------
 *  * The flag defaults to **TRUE** (legacy behaviour) when the add-in loads.
 *  * It persists only for the current Excel session; store `=SetVolatility(FALSE)` in a cell
 *    or run it from VBA/Auto-open if you want it off by default for a workbook.
 *
*/

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

public class C
{
    //--------------------------------------------------------------------
    // Version info
    //--------------------------------------------------------------------
    private const string VERSION_MAJOR = "3";
    private const string VERSION_MINOR = "4";
    private const string VERSION_PATCH = "0";
    private const string CurrentVersion = VERSION_MAJOR + "." + VERSION_MINOR + "." + VERSION_PATCH;
    private static string _targetVersion = CurrentVersion;

    public static string Version { get { return CurrentVersion; } }
    public static string TargetVersion
    {
        get { return _targetVersion; }
        set { if (System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d+\.\d+\.\d+$")) _targetVersion = value; }
    }

    //--------------------------------------------------------------------
    // Global volatility switch
    //--------------------------------------------------------------------
    private static bool _enableVolatility = false; // default is FALSE to avoid performance issues

    [ExcelFunction(Name = "SetVolatility", Description = "Enable (TRUE) or disable (FALSE) volatility for all UDFs", Category = "ExcelDNA Utilities", IsMacroType = true)]
    public static string SetVolatility([ExcelArgument(Description = "TRUE to enable, FALSE to disable")] bool enable)
    {
        _enableVolatility = enable;
        return "Volatility " + (_enableVolatility ? "ENABLED" : "DISABLED");
    }

    [ExcelFunction(Name = "GetVolatility", Description = "Returns current volatility status", Category = "ExcelDNA Utilities")]
    public static string GetVolatility()
    {
        return _enableVolatility ? "ENABLED" : "DISABLED";
    }

    internal static void MaybeVolatile()
    {
        if (_enableVolatility)
        {
            try { XlCall.Excel(XlCall.xlfVolatile, true); } catch { }
        }
    }

    //--------------------------------------------------------------------
    // State dictionaries
    //--------------------------------------------------------------------
    private static readonly Dictionary<string, object> objectStore = new Dictionary<string, object>();
    private static readonly Dictionary<string, object> injectedCells = new Dictionary<string, object>();
    private static readonly Dictionary<string, Tuple<object, object>> invocationCache = new Dictionary<string, Tuple<object, object>>();
    private static readonly Dictionary<string, Tuple<object, object>> visibilityCache = new Dictionary<string, Tuple<object, object>>();

    private static Excel.Application _excelApp;
    private static Excel.Application _app;
    private const int defCachingTime = 10; // seconds

    //--------------------------------------------------------------------
    // Excel helpers
    //--------------------------------------------------------------------
    public static void AttachEvents()
    {
        _excelApp = (Excel.Application)ExcelDnaUtil.Application;
        if (_excelApp != null) _excelApp.WorkbookBeforeClose += WorkbookBeforeClose;
    }

    private static void WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel) { Cleanup(); }

    public static void DetachEvents() { if (_excelApp != null) _excelApp.WorkbookBeforeClose -= WorkbookBeforeClose; }

    public static Excel.Application App
    {
        get
        {
            if (_app == null) _app = (Excel.Application)ExcelDnaUtil.Application;
            return _app;
        }
    }

    public static void Cleanup()
    {
        if (_app != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
            _app = null;
        }
    }

    //--------------------------------------------------------------------
    // 1. Version helpers
    //--------------------------------------------------------------------
    [ExcelFunction(Name = "vExcelDNA", Description = "Returns the version of this UDF collection", Category = "ExcelDNA Utilities")]
    public static string GetExcelDnaVersion() { return CurrentVersion; }

    [ExcelFunction(Name = "SetTargetVersion", Description = "Sets the target version for backward compatibility", Category = "ExcelDNA Utilities", IsMacroType = true)]
    public static string SetTargetVersion(string version)
    {
        string prev = TargetVersion;
        TargetVersion = version;
        return "Target version changed from " + prev + " to " + TargetVersion;
    }

    [ExcelFunction(Name = "GetTargetVersion", Description = "Gets the current target version", Category = "ExcelDNA Utilities")]
    public static string GetTargetVersionFunction() { return TargetVersion; }

    //--------------------------------------------------------------------
    // 2. RecalcAll
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Triggers a full workbook recalculation", Category = "ExcelDNA Utilities")]
    public static object RecalcAll()
    {
        try
        {
            Excel.Application xl = (Excel.Application)ExcelDnaUtil.Application;
            if (xl == null) return ExcelError.ExcelErrorValue;
            ExcelAsyncUtil.QueueAsMacro(delegate { try { xl.CalculateFull(); } catch { } });
            return "TRUE";
        }
        catch { return ExcelError.ExcelErrorValue; }
    }

    //--------------------------------------------------------------------
    // 3. Iteration settings
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Returns Excel iterative calculation settings", Category = "ExcelDNA Utilities")]
    public static string GetIterationStatus()
    {
        MaybeVolatile();
        try
        {
            bool on = App.Iteration;
            return "Status: " + (on ? "ON" : "OFF") + "  Max Iterations: " + App.MaxIterations + "  Max Change: " + App.MaxChange;
        }
        catch (Exception ex) { return ex.Message; }
    }

    [ExcelFunction(Description = "Enable/disable iterative calculation and set parameters", Category = "ExcelDNA Utilities")]
    public static string SetIteration(bool IterationOn, int maxIterations, double maxChange)
    {
        try
        {
            App.Iteration = IterationOn;
            App.MaxIterations = (maxIterations < 1 ? 100 : maxIterations);
            App.MaxChange = (maxChange > 0.0 && maxChange < 1.0) ? maxChange : 0.001;
        }
        catch (Exception ex) { return ex.Message; }
        return GetIterationStatus();
    }

    //--------------------------------------------------------------------
    // 4. IsVisible
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "TRUE if caller cell is visible (row/col not hidden)", Category = "ExcelDNA Utilities", IsMacroType = true)]
    public static object IsVisible(int cachingTime)
    {
        MaybeVolatile();
        try
        {
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (caller == null) return ExcelError.ExcelErrorRef;
            string address = (string)XlCall.Excel(XlCall.xlfReftext, caller, true);

            Tuple<object, object> tup;
            if (visibilityCache.TryGetValue(address, out tup))
            {
                DateTime ts = (DateTime)tup.Item1;
                bool vis = (bool)tup.Item2;
                if ((DateTime.Now - ts).TotalSeconds < cachingTime) return vis ? "TRUE" : "FALSE";
            }

            Excel.Range rng = App.Range[address];
            bool rowHidden = rng.EntireRow.Hidden is bool ? (bool)rng.EntireRow.Hidden : false;
            bool columnHidden = rng.EntireColumn.Hidden is bool ? (bool)rng.EntireColumn.Hidden : false;
            bool visible = !(rowHidden || columnHidden);
            visibilityCache[address] = new Tuple<object, object>(DateTime.Now, visible);
            return visible ? "TRUE" : "FALSE";
        }
        catch (Exception ex) { return ex.Message; }
    }

    //--------------------------------------------------------------------
    // 5. Describe
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Describes a value or reference", Category = "ExcelDNA Utilities", IsMacroType = true)]
    public static string Describe(object arg)
    {
        if (arg is double) return "Double: " + (double)arg;
        if (arg is string) return "String: " + (string)arg;
        if (arg is bool) return "Boolean: " + ((bool)arg);
        if (arg is ExcelError) return "ExcelError: " + arg.ToString();
        if (arg is object[,])
        {
            object[,] arr = (object[,])arg;
            return "Array[" + arr.GetLength(0) + "," + arr.GetLength(1) + "]";
        }
        if (arg is ExcelMissing) return "Missing";
        if (arg is ExcelEmpty) return "Empty";
        if (arg is ExcelReference) return "Reference: " + XlCall.Excel(XlCall.xlfReftext, arg, true);
        return "!?Unheard Of";
    }

    //--------------------------------------------------------------------
    // 6. InjectValue
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Injects a value into a cell (stateful)", Category = "ExcelDNA Utilities")]
    public static object InjectValue([ExcelArgument(AllowReference = true)] object potentialRef, object value)
    {
        MaybeVolatile();
        if (potentialRef == null || value == null) return ExcelError.ExcelErrorValue;
        ExcelReference cellRef = potentialRef as ExcelReference;
        if (cellRef == null) return "Error: first argument must be a cell reference.";

        string address = (string)XlCall.Excel(XlCall.xlfAddress, 1 + cellRef.RowFirst, 1 + cellRef.ColumnFirst);
        string key = cellRef.SheetId + "!" + address;

        object[,] box = new object[1, 1]; box[0, 0] = value;
        object prev;
        if (injectedCells.TryGetValue(key, out prev) && Equals(prev, value)) return box;

        ExcelAsyncUtil.QueueAsMacro(delegate { try { cellRef.SetValue(box); injectedCells[key] = value; } catch { } });
        return box;
    }

    //--------------------------------------------------------------------
    // 7. FINDPOS
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Returns the Nth (or last=-1) position of substring (case-insensitive)", Category = "ExcelDNA Utilities")]
    public static object FindPos(string text, string substring, int instance)
    {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(substring)) return ExcelError.ExcelErrorValue;
        string t = text.ToLower();
        string sub = substring.ToLower();
        List<int> idx = new List<int>();
        int p = t.IndexOf(sub, StringComparison.Ordinal);
        while (p != -1)
        {
            idx.Add(p + 1); // 1‑based for Excel
            p = t.IndexOf(sub, p + 1, StringComparison.Ordinal);
        }
        if (instance == -1)
        {
            if (idx.Count == 0) return ExcelError.ExcelErrorValue;
            return idx[idx.Count - 1];
        }
        if (instance > 0 && instance <= idx.Count) return idx[instance - 1];
        return ExcelError.ExcelErrorValue;
    }

    //--------------------------------------------------------------------
    // 8. PutObject / GetObject / PurgeObjects
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Stores an object in a temporary cache", Category = "ExcelDNA Utilities")]
    public static object PutObject(string name, object value, bool force, bool debug)
    {
        MaybeVolatile();
        if (string.IsNullOrWhiteSpace(name)) return debug ? "Error: name empty" : (object)ExcelError.ExcelErrorValue;

        ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        string callerAddr = (string)XlCall.Excel(XlCall.xlfAddress, 1 + caller.RowFirst, 1 + caller.ColumnFirst);
        string cacheKey = callerAddr + ":" + name;

        Tuple<object, object> tup;
        if (invocationCache.TryGetValue(cacheKey, out tup))
        {
            if (Equals(tup.Item2, value)) return value; // redundant write
        }
        invocationCache[cacheKey] = new Tuple<object, object>(callerAddr, value);

        if (objectStore.ContainsKey(name) && !force)
        {
            if (debug) return "Exists";
            return (object)ExcelError.ExcelErrorName;
        }
        objectStore[name] = value;
        return value;
    }

    [ExcelFunction(Description = "Retrieves an object from the temporary cache", Category = "ExcelDNA Utilities")]
    public static object GetObject(string name, bool debug)
    {
        MaybeVolatile();
        if (string.IsNullOrWhiteSpace(name)) return debug ? "Error: name empty" : (object)ExcelError.ExcelErrorValue;
        if (!objectStore.ContainsKey(name)) return debug ? "Error: not found" : (object)ExcelError.ExcelErrorName;
        object obj = objectStore[name];
        if (obj == null) return debug ? "Error: null" : (object)ExcelError.ExcelErrorValue;
        return obj;
    }

    [ExcelFunction(Description = "Clears all stored objects", Category = "ExcelDNA Utilities")]
    public static string PurgeObjects() { objectStore.Clear(); return "TRUE"; }

    //--------------------------------------------------------------------
    // 9. TrueSplit
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Splits strings by delimiter and returns dynamic array", Category = "ExcelDNA Utilities")]
    public static object[,] TrueSplit(object[] inputStrings, string delimiter)
    {
        int maxCols = 1;
        for (int i = 0; i < inputStrings.Length; i++)
        {
            string sTmp = inputStrings[i] as string;
            if (sTmp != null)
            {
                int cnt = sTmp.Split(new string[] { delimiter }, StringSplitOptions.None).Length;
                if (cnt > maxCols) maxCols = cnt;
            }
        }
        object[,] result = new object[inputStrings.Length, maxCols];
        for (int r = 0; r < inputStrings.Length; r++)
        {
            string s = inputStrings[r] as string;
            if (s != null)
            {
                string[] parts = s.Split(new string[] { delimiter }, StringSplitOptions.None);
                for (int c = 0; c < parts.Length; c++) result[r, c] = parts[c];
            }
            else if (inputStrings[r] is ExcelError)
            {
                result[r, 0] = inputStrings[r];
            }
            else
            {
                result[r, 0] = inputStrings[r] == null ? "" : inputStrings[r].ToString();
            }
        }
        return result;
    }

    //--------------------------------------------------------------------
    // 10. IsMemberOf
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "TRUE if any element/row/col of A exists in B", Category = "ExcelDNA Utilities")]
    public static bool IsMemberOf(object[,] arrayA, object[,] arrayB)
    {
        int aRows = arrayA.GetLength(0), aCols = arrayA.GetLength(1);
        int bRows = arrayB.GetLength(0), bCols = arrayB.GetLength(1);
        bool aSingle = (aRows == 1 && aCols == 1);
        bool bSingle = (bRows == 1 && bCols == 1);
        if (aSingle || bSingle)
        {
            object aVal = arrayA[0, 0];
            if (bSingle) return AreEqual(aVal, arrayB[0, 0]);
            for (int i = 0; i < bRows; i++)
                for (int j = 0; j < bCols; j++) if (AreEqual(aVal, arrayB[i, j])) return true;
            return false;
        }

        bool compareRows = (aCols == bCols);
        bool compareCols = (aRows == bRows);
        if (!compareRows && !compareCols) return false;

        if (compareRows)
        {
            for (int ar = 0; ar < aRows; ar++)
            {
                for (int br = 0; br < bRows; br++)
                {
                    bool match = true;
                    for (int c = 0; c < aCols && match; c++) if (!AreEqual(arrayA[ar, c], arrayB[br, c])) match = false;
                    if (match) return true;
                }
            }
        }
        if (compareCols)
        {
            for (int ac = 0; ac < aCols; ac++)
            {
                for (int bc = 0; bc < bCols; bc++)
                {
                    bool match = true;
                    for (int r = 0; r < aRows && match; r++) if (!AreEqual(arrayA[r, ac], arrayB[r, bc])) match = false;
                    if (match) return true;
                }
            }
        }
        return false;
    }

    //--------------------------------------------------------------------
    // 11. GetThreads & SetThreads
    //--------------------------------------------------------------------
    [ExcelFunction(Name = "GetThreads", Description = "Returns multithreading settings", Category = "ExcelDNA Utilities")]
    public static object GetThreads()
    {
        MaybeVolatile();
        try
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            if (new Version(app.Version) < new Version("12.0")) return "Excel 2007+ required";
            Excel.MultiThreadedCalculation mtc = app.MultiThreadedCalculation;
            int max = 64;
            return new object[,] { { "Current Thread Count", mtc.ThreadCount }, { "Max Available", max }, { "Mode Enabled", mtc.Enabled } };
        }
        catch { return ExcelError.ExcelErrorValue; }
    }

    private static int _lastThreadCount = -2;
    private static bool _lastThreadEnabled;
    private static readonly object _threadLock = new object();

    [ExcelFunction(Name = "SetThreads", Description = "Configures multithreading", Category = "ExcelDNA Utilities", IsMacroType = true)]
    public static object SetThreads(int threadCount, bool enable)
    {
        lock (_threadLock)
        {
            try
            {
                if (_lastThreadCount == threadCount && _lastThreadEnabled == enable) return "Cached";
                ExcelAsyncUtil.QueueAsMacro(delegate
                {
                    Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                    if (new Version(app.Version) < new Version("12.0")) return;
                    Excel.MultiThreadedCalculation mtc = app.MultiThreadedCalculation;
                    int max = 64;
                    int newCount = (threadCount == -1) ? max : (threadCount == 0 ? max / 2 : (threadCount > max ? max : threadCount));
                    if (mtc.ThreadCount != newCount || mtc.Enabled != enable)
                    {
                        mtc.ThreadCount = newCount;
                        mtc.Enabled = enable;
                        _lastThreadCount = threadCount;
                        _lastThreadEnabled = enable;
                        if (enable) app.CalculateFullRebuild();
                    }
                });
                return "Thread settings updated";
            }
            catch { return ExcelError.ExcelErrorValue; }
        }
    }

    //--------------------------------------------------------------------
    // 12. HashArray
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "Returns a stable hash for an array (order‑independent)", Category = "ExcelDNA Utilities")]
    public static object HashArray(object[,] inputArray, object hashLengthObj)
    {
        int hashLen = 8;
        if (hashLengthObj is double) hashLen = (int)(double)hashLengthObj;
        else if (hashLengthObj is int) hashLen = (int)hashLengthObj;
        else if (hashLengthObj is string)
        {
            int parsed; if (int.TryParse((string)hashLengthObj, out parsed)) hashLen = parsed;
        }
        if (hashLen < 4) hashLen = 4; if (hashLen > 32) hashLen = 32;

        List<string> elems = new List<string>();
        int rows = inputArray.GetLength(0), cols = inputArray.GetLength(1);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                object el = inputArray[r, c];
                if (el == null || el is ExcelEmpty) continue;
                if (el is ExcelError) elems.Add("ERROR:" + el.ToString());
                else if (el is double) elems.Add(((double)el).ToString("G17"));
                else elems.Add(el.ToString());
            }
        elems.Sort();
        string combined = string.Join("|", elems.ToArray());
        return GenerateHash(combined, hashLen);
    }

    //--------------------------------------------------------------------
    // 13. isLocalIP
    //--------------------------------------------------------------------
    [ExcelFunction(Description = "TRUE if IP is local/private", Category = "ExcelDNA Utilities")]
    public static object isLocalIP(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return ExcelError.ExcelErrorNA;
        try
        {
            string ipOnly = input;
            if (ipOnly.StartsWith("[") && ipOnly.IndexOf(']') > 0)
            {
                int end = ipOnly.IndexOf(']');
                ipOnly = ipOnly.Substring(1, end - 1);
            }
            int colon = ipOnly.LastIndexOf(':');
            if (colon > -1 && ipOnly.IndexOf(':') == colon) ipOnly = ipOnly.Substring(0, colon);
            System.Net.IPAddress ip;
            if (!System.Net.IPAddress.TryParse(ipOnly, out ip)) return ExcelError.ExcelErrorNA;

            byte[] b = ip.GetAddressBytes();
            if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            {
                if (b[0] == 10) return true;
                if (b[0] == 172 && b[1] >= 16 && b[1] <= 31) return true;
                if (b[0] == 192 && b[1] == 168) return true;
                if (b[0] == 127) return true;
                if (b[0] == 169 && b[1] == 254) return true;
                return false;
            }
            if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
            {
                if (System.Net.IPAddress.IsLoopback(ip)) return true;
                if (ip.IsIPv6LinkLocal || ip.IsIPv6SiteLocal) return true;
                if ((ip.GetAddressBytes()[0] & 0xFE) == 0xFC) return true; // fc00::/7
                return false;
            }
            return false;
        }
        catch { return ExcelError.ExcelErrorNA; }
    }

    //--------------------------------------------------------------------
    // 14. ARRAYSUBTRACT
    //--------------------------------------------------------------------
    [ExcelFunction(Name = "ARRAYSUBTRACT", Description = "Array subtraction (preserves shape)", Category = "ExcelDNA Utilities")]
    public static object[,] ArraySubtract(object[,] arrayA, object[,] arrayB)
    {
        HashSet<string> remove = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int br = arrayB.GetLength(0), bc = arrayB.GetLength(1);
        for (int i = 0; i < br; i++)
            for (int j = 0; j < bc; j++)
            {
                object v = arrayB[i, j];
                if (v != null && !(v is ExcelEmpty) && !(v is ExcelError)) remove.Add(v.ToString());
            }
        int ar = arrayA.GetLength(0), ac = arrayA.GetLength(1);
        bool isRow = (ar == 1 && ac > 1);
        List<object> kept = new List<object>();
        for (int i = 0; i < ar; i++)
            for (int j = 0; j < ac; j++)
            {
                object v = arrayA[i, j];
                if (v == null || v is ExcelEmpty || v is ExcelError) continue;
                if (!remove.Contains(v.ToString())) kept.Add(v);
            }
        if (isRow)
        {
            object[,] res = new object[1, kept.Count];
            for (int i = 0; i < kept.Count; i++) res[0, i] = kept[i];
            return res;
        }
        object[,] resCol = new object[kept.Count, 1];
        for (int i = 0; i < kept.Count; i++) resCol[i, 0] = kept[i];
        return resCol;
    }

    //--------------------------------------------------------------------
    // 15. EXTRACTSUBSTR
    //--------------------------------------------------------------------
    [ExcelFunction(Name = "EXTRACTSUBSTR", Description = "Extracts substring between start and end markers", Category = "ExcelDNA Utilities")]
    public static object ExtractSubstr(
    [ExcelArgument(Description = "String to extract from")] string inputString,
    [ExcelArgument(Description = "Text that precedes the substring to extract")] string startMarker,
    [ExcelArgument(Description = "Text that marks the end of substring (not included in result)")] object endMarkerObj)
    {
        try
        {
            // Validate required parameters
            if (string.IsNullOrEmpty(inputString) || string.IsNullOrEmpty(startMarker))
                return ExcelError.ExcelErrorNA;

            string endMarker = (endMarkerObj is ExcelMissing || endMarkerObj is ExcelEmpty) ? null : endMarkerObj.ToString();

            // Find start position
            int startPos = inputString.IndexOf(startMarker, StringComparison.Ordinal);
            if (startPos == -1)
                return ExcelError.ExcelErrorNA;

            // Calculate where to start extracting (after the start marker)
            int extractStart = startPos + startMarker.Length;

            // Case 1: No end marker - return everything after start marker
            if (string.IsNullOrEmpty(endMarker))
            {
                if (extractStart >= inputString.Length)
                    return string.Empty;
                return inputString.Substring(extractStart);
            }

            // Case 2: With end marker - find end position
            int endPos = inputString.IndexOf(endMarker, extractStart, StringComparison.Ordinal);
            if (endPos == -1)
                return ExcelError.ExcelErrorNA;

            // Extract substring between markers
            return inputString.Substring(extractStart, endPos - extractStart);
        }
        catch
        {
            return ExcelError.ExcelErrorNA;
        }
    }

    //--------------------------------------------------------------------
    // 16. STRING_COMMON
    //--------------------------------------------------------------------
    [ExcelFunction(Name = "STRING_COMMON", Description = "Returns maximal common substrings with a minimum length", Category = "ExcelDNA Utilities")]
    public static object[,] StringCommon(
    [ExcelArgument(Description = "First string")] string s1,
    [ExcelArgument(Description = "Second string")] string s2,
    [ExcelArgument(Description = "Minimum substring length")] int minLength)
    {
        if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2) || minLength < 1)
        {
            return new object[0, 0];
        }

        List<SubstringMatch> matches = GetCommonRunsFromLcs(s1, s2);
        List<string> results = new List<string>();

        foreach (var match in matches)
        {
            if (match.Length >= minLength) results.Add(match.Value);
        }

        return BuildColumnArray(results);
    }

    //--------------------------------------------------------------------
    // 17. STRING_DIFF
    //--------------------------------------------------------------------
    [ExcelFunction(Name = "STRING_DIFF", Description = "Returns maximal differing substrings with a minimum length", Category = "ExcelDNA Utilities")]
    public static object[,] StringDiff(
    [ExcelArgument(Description = "First string")] string s1,
    [ExcelArgument(Description = "Second string")] string s2,
    [ExcelArgument(Description = "Minimum substring length")] int minLength)
    {
        if (string.IsNullOrEmpty(s1) || minLength < 1)
        {
            return new object[0, 0];
        }

        if (string.IsNullOrEmpty(s2))
        {
            return (s1.Length >= minLength) ? BuildColumnArray(new List<string> { s1 }) : new object[0, 0];
        }

        List<SubstringMatch> selected = GetCommonRunsFromLcs(s1, s2);

        List<string> diffs = new List<string>();
        int current = 0;
        foreach (var match in selected)
        {
            if (match.Start1 > current)
            {
                int length = match.Start1 - current;
                if (length >= minLength)
                {
                    diffs.Add(s1.Substring(current, length));
                }
            }
            current = match.Start1 + match.Length;
        }

        if (current < s1.Length)
        {
            int length = s1.Length - current;
            if (length >= minLength)
            {
                diffs.Add(s1.Substring(current, length));
            }
        }

        return BuildColumnArray(diffs);
    }

    //--------------------------------------------------------------------
    // Utility helpers
    //--------------------------------------------------------------------
    private struct SubstringMatch
    {
        public int Start1;
        public int Start2;
        public int Length;
        public string Value;
    }

    private static List<SubstringMatch> GetCommonRunsFromLcs(string s1, string s2)
    {
        int len1 = s1.Length;
        int len2 = s2.Length;
        int[,] dp = new int[len1 + 1, len2 + 1];

        for (int i = 1; i <= len1; i++)
        {
            for (int j = 1; j <= len2; j++)
            {
                if (s1[i - 1] == s2[j - 1])
                {
                    dp[i, j] = dp[i - 1, j - 1] + 1;
                }
                else
                {
                    dp[i, j] = dp[i - 1, j] >= dp[i, j - 1] ? dp[i - 1, j] : dp[i, j - 1];
                }
            }
        }

        List<Tuple<int, int>> matches = new List<Tuple<int, int>>();
        int x = len1;
        int y = len2;
        while (x > 0 && y > 0)
        {
            if (s1[x - 1] == s2[y - 1])
            {
                matches.Add(new Tuple<int, int>(x - 1, y - 1));
                x--;
                y--;
            }
            else if (dp[x - 1, y] >= dp[x, y - 1])
            {
                x--;
            }
            else
            {
                y--;
            }
        }

        matches.Reverse();

        List<SubstringMatch> runs = new List<SubstringMatch>();
        if (matches.Count == 0) return runs;

        int runStart1 = matches[0].Item1;
        int runStart2 = matches[0].Item2;
        int runLength = 1;

        for (int i = 1; i < matches.Count; i++)
        {
            int prev1 = matches[i - 1].Item1;
            int prev2 = matches[i - 1].Item2;
            int curr1 = matches[i].Item1;
            int curr2 = matches[i].Item2;

            if (curr1 == prev1 + 1 && curr2 == prev2 + 1)
            {
                runLength++;
            }
            else
            {
                runs.Add(new SubstringMatch
                {
                    Start1 = runStart1,
                    Start2 = runStart2,
                    Length = runLength,
                    Value = s1.Substring(runStart1, runLength)
                });
                runStart1 = curr1;
                runStart2 = curr2;
                runLength = 1;
            }
        }

        runs.Add(new SubstringMatch
        {
            Start1 = runStart1,
            Start2 = runStart2,
            Length = runLength,
            Value = s1.Substring(runStart1, runLength)
        });

        return runs;
    }

    private static object[,] BuildColumnArray(List<string> items)
    {
        if (items == null || items.Count == 0) return new object[0, 0];
        object[,] result = new object[items.Count, 1];
        for (int i = 0; i < items.Count; i++) result[i, 0] = items[i];
        return result;
    }

    private static bool AreEqual(object a, object b)
    {
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;
        if (a is ExcelEmpty || b is ExcelEmpty) return false;
        if (a is ExcelError || b is ExcelError) return false;
        return a.ToString() == b.ToString();
    }

    private static string GenerateHash(string txt, int len)
    {
        using (var sha = System.Security.Cryptography.SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(txt ?? string.Empty));
            string b64 = Convert.ToBase64String(hash).Replace("+", "0").Replace("/", "1").Replace("=", "2");
            if (len < 4) len = 4; if (len > 32) len = 32;
            return b64.Substring(0, len);
        }
    }
}
