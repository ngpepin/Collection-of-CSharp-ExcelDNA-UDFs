/*
 *
 * ExcelDNA User-Defined Functions 
 *
 * This collection provides powerful worksheet functions that extend Excel's native capabilities. All functions are thread-safe and designed for high performance 
 * in large spreadsheets. Stateful functions (like INJECTVALUE) maintain state between calculations and as such violate Excel's "no side effects" rule 
 * (intentionally of course, because by doing so they allow state machines to be created in spreadsheet models!)
 *
 * VEXCELDNA, SETTARGETVERSION, GETTARGETVERSION, RECALCALL, GETITERATIONSTATUS, SETITERATION, ISVISIBLE, DESCRIBE, INJECTVALUE, FINDPOS,
 * PUTOBJECT, GETOBJECT, PURGEOBJECTS, TRUESPLIT, ISMEMBEROF, GETTHREADS, SETTHREADS, HASHARRAY, ISLOCALIP, ARRAYSUBTRACT,
 * EXTRACTSUBSTR, STRING_COMMON, STRING_DIFF, TEXT_BEFORE, TEXT_AFTER, REGEX_ISMATCH, REGEX_EXTRACT, REGEX_REPLACE, ARRAY_UNIQUE,
 * ARRAY_DISTINCT_COUNT, NUM_CLAMP, VECTOR_DOT, VECTOR_NORM, VECTOR_NORMALIZE, VECTOR_COSINE_SIMILARITY,
 * VECTOR_EUCLIDEAN_DISTANCE, VECTOR_MANHATTAN_DISTANCE, VECTOR_SOFTMAX, VECTOR_SIGMOID, VECTOR_RELU,
 * MATRIX_STANDARDIZE_COLUMNS, MATRIX_MINMAX_SCALE_COLUMNS, MATRIX_PAIRWISE_DISTANCE, MATRIX_COVARIANCE, MATRIX_ONE_HOT, MATRIX_CONFUSION,
 * VECTOR_LOG_SOFTMAX, VECTOR_TOP_K, MATRIX_LINEAR_PREDICT, MATRIX_CORRELATION, MATRIX_KMEANS_ASSIGN
 *
 * New in version 3.9.0:
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
 *    - Returns: Dynamic array of differing substrings from s1 and s2 (empty if none meet minLength)
 *
 * 24. TEXT_BEFORE(text, delimiter, [instance])
 *    - Returns text before the requested delimiter occurrence
 *
 * 25. TEXT_AFTER(text, delimiter, [instance])
 *    - Returns text after the requested delimiter occurrence
 *
 * 26. REGEX_ISMATCH(text, pattern, [ignoreCase])
 *    - Tests whether text matches a regular expression
 *
 * 27. REGEX_EXTRACT(text, pattern, [group])
 *    - Returns a regular-expression match or capture group
 *
 * 28. REGEX_REPLACE(text, pattern, replacement, [ignoreCase])
 *    - Replaces regular-expression matches
 *
 * 29. ARRAY_UNIQUE(inputArray, [ignoreCase])
 *    - Returns unique nonblank values while preserving first-seen order
 *
 * 30. ARRAY_DISTINCT_COUNT(inputArray, [ignoreCase])
 *    - Counts unique nonblank values
 *
 * 31. NUM_CLAMP(value, minimum, maximum)
 *    - Restricts a number to an inclusive range
 *
 * 32. VECTOR_DOT(vectorA, vectorB)
 *    - Computes the dot product of two equally sized numeric vectors
 *
 * 33. VECTOR_NORM(vector, [p])
 *    - Computes an L-p vector norm; p defaults to 2
 *
 * 34. VECTOR_NORMALIZE(vector, [p])
 *    - Returns a spill vector normalized to unit L-p norm
 *
 * 35. VECTOR_COSINE_SIMILARITY(vectorA, vectorB)
 *    - Computes cosine similarity between two vectors
 *
 * 36. VECTOR_EUCLIDEAN_DISTANCE(vectorA, vectorB)
 *    - Computes Euclidean distance between two vectors
 *
 * 37. VECTOR_MANHATTAN_DISTANCE(vectorA, vectorB)
 *    - Computes Manhattan distance between two vectors
 *
 * 38. VECTOR_SOFTMAX(vector)
 *    - Returns a numerically stable softmax probability vector
 *
 * 39. VECTOR_SIGMOID(vector)
 *    - Applies the logistic sigmoid element-wise and spills the result
 *
 * 40. VECTOR_RELU(vector)
 *    - Applies rectified linear activation element-wise
 *
 * 41. MATRIX_STANDARDIZE_COLUMNS(matrix, [sample])
 *    - Z-score standardizes each feature column
 *
 * 42. MATRIX_MINMAX_SCALE_COLUMNS(matrix, [targetMin], [targetMax])
 *    - Min-max scales each feature column into a requested range
 *
 * 43. MATRIX_PAIRWISE_DISTANCE(matrix, [metric])
 *    - Builds a row-to-row Euclidean, Manhattan, or cosine distance matrix
 *
 * 44. MATRIX_COVARIANCE(matrix, [sample])
 *    - Returns the feature covariance matrix for row observations
 *
 * 45. MATRIX_ONE_HOT(labels, [classLabels])
 *    - One-hot encodes a label vector into a dynamic matrix
 *
 * 46. MATRIX_CONFUSION(actual, predicted, [classLabels])
 *    - Builds a multiclass confusion matrix from actual and predicted labels
 *
 * 47. VECTOR_LOG_SOFTMAX(vector)
 *    - Returns numerically stable log-softmax scores
 *
 * 48. VECTOR_TOP_K(vector, k, [largest])
 *    - Returns the 1-based indices and values of the top or bottom k elements
 *
 * 49. MATRIX_LINEAR_PREDICT(matrix, weights, [bias])
 *    - Applies a dense linear layer to row observations
 *
 * 50. MATRIX_CORRELATION(matrix)
 *    - Returns the feature correlation matrix
 *
 * 51. MATRIX_KMEANS_ASSIGN(matrix, centroids, [metric])
 *    - Assigns observations to their nearest centroid
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
    private const string VERSION_MINOR = "9";
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

            Excel.Range rng = App.get_Range(address, Type.Missing);
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
    // AreEqual helper (moved above IsMemberOf for visibility)
    //--------------------------------------------------------------------
    private static bool AreEqual(object a, object b)
    {
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;
        if (a is ExcelEmpty || b is ExcelEmpty) return false;
        if (a is ExcelError || b is ExcelError) return false;
        return a.ToString() == b.ToString();
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

        List<SubstringMatch> matches = GetCommonSubstringsByLongestMatch(s1, s2);
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

        List<SubstringMatch> selected = GetCommonSubstringsByLongestMatch(s1, s2);

        List<string> diffs = new List<string>();
        diffs.AddRange(CollectDiffs(s1, selected.OrderBy(m => m.Start1), minLength, match => match.Start1));
        diffs.AddRange(CollectDiffs(s2, selected.OrderBy(m => m.Start2), minLength, match => match.Start2));

        return BuildColumnArray(diffs);
    }

    //--------------------------------------------------------------------
    // 24. TEXT_BEFORE
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns the text before the requested occurrence of a delimiter.
    /// </summary>
    [ExcelFunction(Name = "TEXT_BEFORE", Description = "Returns text before the requested delimiter occurrence", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object TextBefore(
        [ExcelArgument(Description = "Text to search")] string text,
        [ExcelArgument(Description = "Delimiter to locate")] string delimiter,
        [ExcelArgument(Description = "1-based delimiter occurrence; defaults to 1")] object instanceObj)
    {
        if (text == null || string.IsNullOrEmpty(delimiter)) return ExcelError.ExcelErrorValue;
        int instance;
        if (!TryGetOptionalPositiveInt(instanceObj, 1, out instance)) return ExcelError.ExcelErrorValue;
        int index = FindDelimiterOccurrence(text, delimiter, instance);
        if (index < 0) return ExcelError.ExcelErrorNA;
        return text.Substring(0, index);
    }

    //--------------------------------------------------------------------
    // 25. TEXT_AFTER
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns the text after the requested occurrence of a delimiter.
    /// </summary>
    [ExcelFunction(Name = "TEXT_AFTER", Description = "Returns text after the requested delimiter occurrence", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object TextAfter(
        [ExcelArgument(Description = "Text to search")] string text,
        [ExcelArgument(Description = "Delimiter to locate")] string delimiter,
        [ExcelArgument(Description = "1-based delimiter occurrence; defaults to 1")] object instanceObj)
    {
        if (text == null || string.IsNullOrEmpty(delimiter)) return ExcelError.ExcelErrorValue;
        int instance;
        if (!TryGetOptionalPositiveInt(instanceObj, 1, out instance)) return ExcelError.ExcelErrorValue;
        int index = FindDelimiterOccurrence(text, delimiter, instance);
        if (index < 0) return ExcelError.ExcelErrorNA;
        return text.Substring(index + delimiter.Length);
    }

    //--------------------------------------------------------------------
    // 26. REGEX_ISMATCH
    //--------------------------------------------------------------------
    /// <summary>
    /// Tests whether text contains a match for a regular-expression pattern.
    /// </summary>
    [ExcelFunction(Name = "REGEX_ISMATCH", Description = "TRUE when text matches a regular expression", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object RegexIsMatch(
        [ExcelArgument(Description = "Text to test")] string text,
        [ExcelArgument(Description = "Regular-expression pattern")] string pattern,
        [ExcelArgument(Description = "TRUE for case-insensitive matching; defaults to FALSE")] object ignoreCaseObj)
    {
        if (text == null || string.IsNullOrEmpty(pattern)) return ExcelError.ExcelErrorValue;
        try
        {
            System.Text.RegularExpressions.RegexOptions options = GetOptionalBool(ignoreCaseObj, false)
                ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
                : System.Text.RegularExpressions.RegexOptions.None;
            var regex = new System.Text.RegularExpressions.Regex(pattern, options, TimeSpan.FromSeconds(1));
            return regex.IsMatch(text);
        }
        catch (ArgumentException) { return ExcelError.ExcelErrorValue; }
        catch (System.Text.RegularExpressions.RegexMatchTimeoutException) { return ExcelError.ExcelErrorValue; }
    }

    //--------------------------------------------------------------------
    // 27. REGEX_EXTRACT
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns the first regular-expression match or a requested capture group.
    /// </summary>
    [ExcelFunction(Name = "REGEX_EXTRACT", Description = "Returns the first regex match or capture group", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object RegexExtract(
        [ExcelArgument(Description = "Text to search")] string text,
        [ExcelArgument(Description = "Regular-expression pattern")] string pattern,
        [ExcelArgument(Description = "Optional numeric group index or named group; defaults to 0")] object groupObj)
    {
        if (text == null || string.IsNullOrEmpty(pattern)) return ExcelError.ExcelErrorValue;
        try
        {
            var regex = new System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.None, TimeSpan.FromSeconds(1));
            System.Text.RegularExpressions.Match match = regex.Match(text);
            if (!match.Success) return ExcelError.ExcelErrorNA;
            if (groupObj == null || groupObj is ExcelMissing || groupObj is ExcelEmpty) return match.Value;

            int groupIndex;
            if (TryGetInt(groupObj, out groupIndex))
            {
                if (groupIndex < 0 || groupIndex >= match.Groups.Count) return ExcelError.ExcelErrorNA;
                return match.Groups[groupIndex].Success ? (object)match.Groups[groupIndex].Value : ExcelError.ExcelErrorNA;
            }

            string groupName = groupObj.ToString();
            if (string.IsNullOrWhiteSpace(groupName)) return match.Value;
            System.Text.RegularExpressions.Group group = match.Groups[groupName];
            return group.Success ? (object)group.Value : ExcelError.ExcelErrorNA;
        }
        catch (ArgumentException) { return ExcelError.ExcelErrorValue; }
        catch (System.Text.RegularExpressions.RegexMatchTimeoutException) { return ExcelError.ExcelErrorValue; }
    }

    //--------------------------------------------------------------------
    // 28. REGEX_REPLACE
    //--------------------------------------------------------------------
    /// <summary>
    /// Replaces all regular-expression matches in text.
    /// </summary>
    [ExcelFunction(Name = "REGEX_REPLACE", Description = "Replaces all regular-expression matches", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object RegexReplace(
        [ExcelArgument(Description = "Text to transform")] string text,
        [ExcelArgument(Description = "Regular-expression pattern")] string pattern,
        [ExcelArgument(Description = "Replacement text")] string replacement,
        [ExcelArgument(Description = "TRUE for case-insensitive matching; defaults to FALSE")] object ignoreCaseObj)
    {
        if (text == null || string.IsNullOrEmpty(pattern) || replacement == null) return ExcelError.ExcelErrorValue;
        try
        {
            System.Text.RegularExpressions.RegexOptions options = GetOptionalBool(ignoreCaseObj, false)
                ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
                : System.Text.RegularExpressions.RegexOptions.None;
            var regex = new System.Text.RegularExpressions.Regex(pattern, options, TimeSpan.FromSeconds(1));
            return regex.Replace(text, replacement);
        }
        catch (ArgumentException) { return ExcelError.ExcelErrorValue; }
        catch (System.Text.RegularExpressions.RegexMatchTimeoutException) { return ExcelError.ExcelErrorValue; }
    }

    //--------------------------------------------------------------------
    // 29. ARRAY_UNIQUE
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns unique nonblank values in first-seen order as a vertical spill array.
    /// </summary>
    [ExcelFunction(Name = "ARRAY_UNIQUE", Description = "Returns unique nonblank values in first-seen order", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object[,] ArrayUnique(
        [ExcelArgument(Description = "Range or array to inspect")] object[,] inputArray,
        [ExcelArgument(Description = "TRUE to compare text without case; defaults to FALSE")] object ignoreCaseObj)
    {
        if (inputArray == null) return new object[0, 0];
        return BuildObjectColumnArray(GetUniqueValues(inputArray, GetOptionalBool(ignoreCaseObj, false)));
    }

    //--------------------------------------------------------------------
    // 30. ARRAY_DISTINCT_COUNT
    //--------------------------------------------------------------------
    /// <summary>
    /// Counts unique nonblank values in a range or array.
    /// </summary>
    [ExcelFunction(Name = "ARRAY_DISTINCT_COUNT", Description = "Counts unique nonblank values", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object ArrayDistinctCount(
        [ExcelArgument(Description = "Range or array to inspect")] object[,] inputArray,
        [ExcelArgument(Description = "TRUE to compare text without case; defaults to FALSE")] object ignoreCaseObj)
    {
        if (inputArray == null) return 0;
        return GetUniqueValues(inputArray, GetOptionalBool(ignoreCaseObj, false)).Count;
    }

    //--------------------------------------------------------------------
    // 31. NUM_CLAMP
    //--------------------------------------------------------------------
    /// <summary>
    /// Restricts a number to an inclusive minimum and maximum.
    /// </summary>
    [ExcelFunction(Name = "NUM_CLAMP", Description = "Restricts a number to an inclusive range", Category = "ExcelDNA Utilities", IsThreadSafe = true)]
    public static object NumClamp(
        [ExcelArgument(Description = "Number to restrict")] object valueObj,
        [ExcelArgument(Description = "Inclusive minimum")] object minimumObj,
        [ExcelArgument(Description = "Inclusive maximum")] object maximumObj)
    {
        double value, minimum, maximum;
        if (!TryGetDouble(valueObj, out value) || !TryGetDouble(minimumObj, out minimum) || !TryGetDouble(maximumObj, out maximum)) return ExcelError.ExcelErrorValue;
        if (minimum > maximum) return ExcelError.ExcelErrorValue;
        if (value < minimum) return minimum;
        if (value > maximum) return maximum;
        return value;
    }

    //--------------------------------------------------------------------
    // 32. VECTOR_DOT
    //--------------------------------------------------------------------
    /// <summary>
    /// Computes the dot product of two equally sized numeric vectors.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_DOT", Description = "Dot product of two numeric vectors", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorDot(
        [ExcelArgument(Description = "First row or column vector")] object[,] vectorA,
        [ExcelArgument(Description = "Second row or column vector")] object[,] vectorB)
    {
        double[] a, b;
        int rowsA, colsA, rowsB, colsB;
        if (!TryGetNumericVector(vectorA, out a, out rowsA, out colsA) || !TryGetNumericVector(vectorB, out b, out rowsB, out colsB) || a.Length != b.Length)
            return ExcelError.ExcelErrorValue;
        double sum = 0.0;
        for (int i = 0; i < a.Length; i++) sum += a[i] * b[i];
        return sum;
    }

    //--------------------------------------------------------------------
    // 33. VECTOR_NORM
    //--------------------------------------------------------------------
    /// <summary>
    /// Computes the L-p norm of a numeric vector. The default p value is 2.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_NORM", Description = "L-p norm of a numeric vector", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorNorm(
        [ExcelArgument(Description = "Row or column vector")] object[,] vector,
        [ExcelArgument(Description = "Optional norm order p; defaults to 2")] object pObj)
    {
        double[] values;
        int rows, cols;
        double p;
        if (!TryGetNumericVector(vector, out values, out rows, out cols) || !TryGetOptionalDouble(pObj, 2.0, out p) || p <= 0.0)
            return ExcelError.ExcelErrorValue;
        return ComputeVectorNorm(values, p);
    }

    //--------------------------------------------------------------------
    // 34. VECTOR_NORMALIZE
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns a row or column spill vector normalized to unit L-p norm.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_NORMALIZE", Description = "Normalizes a vector to unit L-p norm", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorNormalize(
        [ExcelArgument(Description = "Row or column vector")] object[,] vector,
        [ExcelArgument(Description = "Optional norm order p; defaults to 2")] object pObj)
    {
        double[] values;
        int rows, cols;
        double p;
        if (!TryGetNumericVector(vector, out values, out rows, out cols) || !TryGetOptionalDouble(pObj, 2.0, out p) || p <= 0.0)
            return ExcelError.ExcelErrorValue;
        double norm = ComputeVectorNorm(values, p);
        if (norm == 0.0) return ExcelError.ExcelErrorDiv0;
        double[] result = new double[values.Length];
        for (int i = 0; i < values.Length; i++) result[i] = values[i] / norm;
        return BuildNumericVector(result, rows, cols);
    }

    //--------------------------------------------------------------------
    // 35. VECTOR_COSINE_SIMILARITY
    //--------------------------------------------------------------------
    /// <summary>
    /// Computes cosine similarity between two equally sized numeric vectors.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_COSINE_SIMILARITY", Description = "Cosine similarity between two vectors", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorCosineSimilarity(
        [ExcelArgument(Description = "First row or column vector")] object[,] vectorA,
        [ExcelArgument(Description = "Second row or column vector")] object[,] vectorB)
    {
        double[] a, b;
        int rowsA, colsA, rowsB, colsB;
        if (!TryGetNumericVector(vectorA, out a, out rowsA, out colsA) || !TryGetNumericVector(vectorB, out b, out rowsB, out colsB) || a.Length != b.Length)
            return ExcelError.ExcelErrorValue;
        double dot = 0.0, normA = 0.0, normB = 0.0;
        for (int i = 0; i < a.Length; i++)
        {
            dot += a[i] * b[i];
            normA += a[i] * a[i];
            normB += b[i] * b[i];
        }
        if (normA == 0.0 || normB == 0.0) return ExcelError.ExcelErrorDiv0;
        return dot / Math.Sqrt(normA * normB);
    }

    //--------------------------------------------------------------------
    // 36. VECTOR_EUCLIDEAN_DISTANCE
    //--------------------------------------------------------------------
    /// <summary>
    /// Computes Euclidean distance between two equally sized numeric vectors.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_EUCLIDEAN_DISTANCE", Description = "Euclidean distance between two vectors", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorEuclideanDistance(
        [ExcelArgument(Description = "First row or column vector")] object[,] vectorA,
        [ExcelArgument(Description = "Second row or column vector")] object[,] vectorB)
    {
        double[] a, b;
        int rowsA, colsA, rowsB, colsB;
        if (!TryGetNumericVector(vectorA, out a, out rowsA, out colsA) || !TryGetNumericVector(vectorB, out b, out rowsB, out colsB) || a.Length != b.Length)
            return ExcelError.ExcelErrorValue;
        double sum = 0.0;
        for (int i = 0; i < a.Length; i++)
        {
            double d = a[i] - b[i];
            sum += d * d;
        }
        return Math.Sqrt(sum);
    }

    //--------------------------------------------------------------------
    // 37. VECTOR_MANHATTAN_DISTANCE
    //--------------------------------------------------------------------
    /// <summary>
    /// Computes Manhattan distance between two equally sized numeric vectors.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_MANHATTAN_DISTANCE", Description = "Manhattan distance between two vectors", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorManhattanDistance(
        [ExcelArgument(Description = "First row or column vector")] object[,] vectorA,
        [ExcelArgument(Description = "Second row or column vector")] object[,] vectorB)
    {
        double[] a, b;
        int rowsA, colsA, rowsB, colsB;
        if (!TryGetNumericVector(vectorA, out a, out rowsA, out colsA) || !TryGetNumericVector(vectorB, out b, out rowsB, out colsB) || a.Length != b.Length)
            return ExcelError.ExcelErrorValue;
        double sum = 0.0;
        for (int i = 0; i < a.Length; i++) sum += Math.Abs(a[i] - b[i]);
        return sum;
    }

    //--------------------------------------------------------------------
    // 38. VECTOR_SOFTMAX
    //--------------------------------------------------------------------
    /// <summary>
    /// Applies a numerically stable softmax transformation and preserves vector orientation.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_SOFTMAX", Description = "Stable softmax probability vector", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorSoftmax([ExcelArgument(Description = "Row or column vector of logits")] object[,] vector)
    {
        double[] values;
        int rows, cols;
        if (!TryGetNumericVector(vector, out values, out rows, out cols)) return ExcelError.ExcelErrorValue;
        double max = values.Max();
        double[] result = new double[values.Length];
        double sum = 0.0;
        for (int i = 0; i < values.Length; i++)
        {
            result[i] = Math.Exp(values[i] - max);
            sum += result[i];
        }
        if (sum == 0.0 || double.IsInfinity(sum) || double.IsNaN(sum)) return ExcelError.ExcelErrorNum;
        for (int i = 0; i < result.Length; i++) result[i] /= sum;
        return BuildNumericVector(result, rows, cols);
    }

    //--------------------------------------------------------------------
    // 39. VECTOR_SIGMOID
    //--------------------------------------------------------------------
    /// <summary>
    /// Applies the logistic sigmoid function element-wise and preserves vector orientation.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_SIGMOID", Description = "Element-wise logistic sigmoid activation", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorSigmoid([ExcelArgument(Description = "Row or column vector")] object[,] vector)
    {
        double[] values;
        int rows, cols;
        if (!TryGetNumericVector(vector, out values, out rows, out cols)) return ExcelError.ExcelErrorValue;
        double[] result = new double[values.Length];
        for (int i = 0; i < values.Length; i++)
        {
            double x = values[i];
            result[i] = x >= 0.0 ? 1.0 / (1.0 + Math.Exp(-x)) : Math.Exp(x) / (1.0 + Math.Exp(x));
        }
        return BuildNumericVector(result, rows, cols);
    }

    //--------------------------------------------------------------------
    // 40. VECTOR_RELU
    //--------------------------------------------------------------------
    /// <summary>
    /// Applies rectified linear activation element-wise and preserves vector orientation.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_RELU", Description = "Element-wise rectified linear activation", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorRelu([ExcelArgument(Description = "Row or column vector")] object[,] vector)
    {
        double[] values;
        int rows, cols;
        if (!TryGetNumericVector(vector, out values, out rows, out cols)) return ExcelError.ExcelErrorValue;
        double[] result = new double[values.Length];
        for (int i = 0; i < values.Length; i++) result[i] = values[i] > 0.0 ? values[i] : 0.0;
        return BuildNumericVector(result, rows, cols);
    }

    //--------------------------------------------------------------------
    // 41. MATRIX_STANDARDIZE_COLUMNS
    //--------------------------------------------------------------------
    /// <summary>
    /// Z-score standardizes each matrix column, treating rows as observations.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_STANDARDIZE_COLUMNS", Description = "Z-score standardizes feature columns", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixStandardizeColumns(
        [ExcelArgument(Description = "Numeric matrix with rows as observations")] object[,] matrix,
        [ExcelArgument(Description = "TRUE for sample standard deviation; defaults to FALSE")] object sampleObj)
    {
        double[,] values;
        if (!TryGetNumericMatrix(matrix, out values)) return ExcelError.ExcelErrorValue;
        bool sample = GetOptionalBool(sampleObj, false);
        int rows = values.GetLength(0), cols = values.GetLength(1);
        if (sample && rows < 2) return ExcelError.ExcelErrorValue;
        double[,] result = new double[rows, cols];
        for (int c = 0; c < cols; c++)
        {
            double mean = 0.0;
            for (int r = 0; r < rows; r++) mean += values[r, c];
            mean /= rows;
            double variance = 0.0;
            for (int r = 0; r < rows; r++)
            {
                double d = values[r, c] - mean;
                variance += d * d;
            }
            variance /= sample ? rows - 1 : rows;
            double sd = Math.Sqrt(variance);
            for (int r = 0; r < rows; r++) result[r, c] = sd == 0.0 ? 0.0 : (values[r, c] - mean) / sd;
        }
        return BuildNumericMatrix(result);
    }

    //--------------------------------------------------------------------
    // 42. MATRIX_MINMAX_SCALE_COLUMNS
    //--------------------------------------------------------------------
    /// <summary>
    /// Min-max scales each matrix column into a requested target range.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_MINMAX_SCALE_COLUMNS", Description = "Min-max scales feature columns", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixMinMaxScaleColumns(
        [ExcelArgument(Description = "Numeric matrix with rows as observations")] object[,] matrix,
        [ExcelArgument(Description = "Optional target minimum; defaults to 0")] object targetMinObj,
        [ExcelArgument(Description = "Optional target maximum; defaults to 1")] object targetMaxObj)
    {
        double[,] values;
        double targetMin, targetMax;
        if (!TryGetNumericMatrix(matrix, out values) || !TryGetOptionalDouble(targetMinObj, 0.0, out targetMin) || !TryGetOptionalDouble(targetMaxObj, 1.0, out targetMax) || targetMin >= targetMax)
            return ExcelError.ExcelErrorValue;
        int rows = values.GetLength(0), cols = values.GetLength(1);
        double[,] result = new double[rows, cols];
        for (int c = 0; c < cols; c++)
        {
            double min = values[0, c], max = values[0, c];
            for (int r = 1; r < rows; r++)
            {
                if (values[r, c] < min) min = values[r, c];
                if (values[r, c] > max) max = values[r, c];
            }
            double range = max - min;
            for (int r = 0; r < rows; r++)
                result[r, c] = range == 0.0 ? targetMin : targetMin + ((values[r, c] - min) / range) * (targetMax - targetMin);
        }
        return BuildNumericMatrix(result);
    }

    //--------------------------------------------------------------------
    // 43. MATRIX_PAIRWISE_DISTANCE
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns a square row-to-row distance matrix using Euclidean, Manhattan, or cosine distance.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_PAIRWISE_DISTANCE", Description = "Pairwise row distance matrix", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixPairwiseDistance(
        [ExcelArgument(Description = "Numeric matrix with rows as observations")] object[,] matrix,
        [ExcelArgument(Description = "Optional metric: euclidean, manhattan, or cosine")] object metricObj)
    {
        double[,] values;
        string metric;
        if (!TryGetNumericMatrix(matrix, out values) || !TryGetDistanceMetric(metricObj, out metric)) return ExcelError.ExcelErrorValue;
        int rows = values.GetLength(0), cols = values.GetLength(1);
        double[,] result = new double[rows, rows];
        double[] norms = null;
        if (metric == "cosine")
        {
            norms = new double[rows];
            for (int r = 0; r < rows; r++)
            {
                double sum = 0.0;
                for (int c = 0; c < cols; c++) sum += values[r, c] * values[r, c];
                norms[r] = Math.Sqrt(sum);
                if (norms[r] == 0.0) return ExcelError.ExcelErrorDiv0;
            }
        }
        for (int i = 0; i < rows; i++)
        {
            for (int j = i; j < rows; j++)
            {
                double distance;
                if (metric == "manhattan")
                {
                    distance = 0.0;
                    for (int c = 0; c < cols; c++) distance += Math.Abs(values[i, c] - values[j, c]);
                }
                else if (metric == "cosine")
                {
                    double dot = 0.0;
                    for (int c = 0; c < cols; c++) dot += values[i, c] * values[j, c];
                    distance = 1.0 - dot / (norms[i] * norms[j]);
                    if (distance < 0.0 && distance > -1e-12) distance = 0.0;
                }
                else
                {
                    double sum = 0.0;
                    for (int c = 0; c < cols; c++)
                    {
                        double d = values[i, c] - values[j, c];
                        sum += d * d;
                    }
                    distance = Math.Sqrt(sum);
                }
                result[i, j] = distance;
                result[j, i] = distance;
            }
        }
        return BuildNumericMatrix(result);
    }

    //--------------------------------------------------------------------
    // 44. MATRIX_COVARIANCE
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns the covariance matrix of feature columns, treating rows as observations.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_COVARIANCE", Description = "Feature covariance matrix", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixCovariance(
        [ExcelArgument(Description = "Numeric matrix with rows as observations")] object[,] matrix,
        [ExcelArgument(Description = "TRUE for sample covariance; defaults to TRUE")] object sampleObj)
    {
        double[,] values;
        if (!TryGetNumericMatrix(matrix, out values)) return ExcelError.ExcelErrorValue;
        bool sample = GetOptionalBool(sampleObj, true);
        int rows = values.GetLength(0), cols = values.GetLength(1);
        if (sample && rows < 2) return ExcelError.ExcelErrorValue;
        double[] means = new double[cols];
        for (int c = 0; c < cols; c++)
        {
            for (int r = 0; r < rows; r++) means[c] += values[r, c];
            means[c] /= rows;
        }
        double denominator = sample ? rows - 1 : rows;
        double[,] result = new double[cols, cols];
        for (int i = 0; i < cols; i++)
        {
            for (int j = i; j < cols; j++)
            {
                double sum = 0.0;
                for (int r = 0; r < rows; r++) sum += (values[r, i] - means[i]) * (values[r, j] - means[j]);
                double covariance = sum / denominator;
                result[i, j] = covariance;
                result[j, i] = covariance;
            }
        }
        return BuildNumericMatrix(result);
    }

    //--------------------------------------------------------------------
    // 45. MATRIX_ONE_HOT
    //--------------------------------------------------------------------
    /// <summary>
    /// One-hot encodes a row or column label vector. Class order is explicit or first-seen.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_ONE_HOT", Description = "One-hot encodes a label vector", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixOneHot(
        [ExcelArgument(Description = "Row or column vector of labels")] object[,] labels,
        [ExcelArgument(Description = "Optional row or column vector defining class order")] object classLabelsObj)
    {
        List<object> labelValues;
        if (!TryGetLabelVector(labels, out labelValues)) return ExcelError.ExcelErrorValue;
        List<object> classes;
        if (!TryGetClassLabels(classLabelsObj, labelValues, out classes)) return ExcelError.ExcelErrorValue;
        Dictionary<string, int> index = BuildLabelIndex(classes);
        object[,] result = new object[labelValues.Count, classes.Count];
        for (int r = 0; r < labelValues.Count; r++)
        {
            int classIndex;
            if (!index.TryGetValue(BuildValueKey(labelValues[r]), out classIndex)) return ExcelError.ExcelErrorNA;
            for (int c = 0; c < classes.Count; c++) result[r, c] = c == classIndex ? 1.0 : 0.0;
        }
        return result;
    }

    //--------------------------------------------------------------------
    // 46. MATRIX_CONFUSION
    //--------------------------------------------------------------------
    /// <summary>
    /// Builds a multiclass confusion matrix with actual classes as rows and predicted classes as columns.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_CONFUSION", Description = "Multiclass confusion matrix", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixConfusion(
        [ExcelArgument(Description = "Actual label vector")] object[,] actual,
        [ExcelArgument(Description = "Predicted label vector")] object[,] predicted,
        [ExcelArgument(Description = "Optional row or column vector defining class order")] object classLabelsObj)
    {
        List<object> actualValues, predictedValues;
        if (!TryGetLabelVector(actual, out actualValues) || !TryGetLabelVector(predicted, out predictedValues) || actualValues.Count != predictedValues.Count)
            return ExcelError.ExcelErrorValue;
        List<object> inferred = new List<object>(actualValues);
        inferred.AddRange(predictedValues);
        List<object> classes;
        if (!TryGetClassLabels(classLabelsObj, inferred, out classes)) return ExcelError.ExcelErrorValue;
        Dictionary<string, int> index = BuildLabelIndex(classes);
        object[,] result = new object[classes.Count, classes.Count];
        for (int r = 0; r < classes.Count; r++)
            for (int c = 0; c < classes.Count; c++) result[r, c] = 0.0;
        for (int i = 0; i < actualValues.Count; i++)
        {
            int actualIndex, predictedIndex;
            if (!index.TryGetValue(BuildValueKey(actualValues[i]), out actualIndex) || !index.TryGetValue(BuildValueKey(predictedValues[i]), out predictedIndex))
                return ExcelError.ExcelErrorNA;
            result[actualIndex, predictedIndex] = (double)result[actualIndex, predictedIndex] + 1.0;
        }
        return result;
    }


    //--------------------------------------------------------------------
    // 47. VECTOR_LOG_SOFTMAX
    //--------------------------------------------------------------------
    /// <summary>
    /// Applies a numerically stable log-softmax transformation and preserves vector orientation.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_LOG_SOFTMAX", Description = "Stable log-softmax vector", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorLogSoftmax([ExcelArgument(Description = "Row or column vector of logits")] object[,] vector)
    {
        double[] values;
        int rows, cols;
        if (!TryGetNumericVector(vector, out values, out rows, out cols)) return ExcelError.ExcelErrorValue;
        double max = values.Max();
        double sum = 0.0;
        for (int i = 0; i < values.Length; i++) sum += Math.Exp(values[i] - max);
        if (sum <= 0.0 || double.IsInfinity(sum) || double.IsNaN(sum)) return ExcelError.ExcelErrorNum;
        double logDenominator = max + Math.Log(sum);
        double[] result = new double[values.Length];
        for (int i = 0; i < values.Length; i++) result[i] = values[i] - logDenominator;
        return BuildNumericVector(result, rows, cols);
    }

    //--------------------------------------------------------------------
    // 48. VECTOR_TOP_K
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns the 1-based source indices and values of the largest or smallest k vector elements.
    /// </summary>
    [ExcelFunction(Name = "VECTOR_TOP_K", Description = "Top or bottom k vector indices and values", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object VectorTopK(
        [ExcelArgument(Description = "Row or column vector")] object[,] vector,
        [ExcelArgument(Description = "Number of elements to return")] object kObj,
        [ExcelArgument(Description = "TRUE for largest values; FALSE for smallest; defaults to TRUE")] object largestObj)
    {
        double[] values;
        int rows, cols, k;
        if (!TryGetNumericVector(vector, out values, out rows, out cols) || !TryGetInt(kObj, out k) || k < 1 || k > values.Length)
            return ExcelError.ExcelErrorValue;
        bool largest = GetOptionalBool(largestObj, true);
        List<IndexedValue> ranked = new List<IndexedValue>(values.Length);
        for (int i = 0; i < values.Length; i++) ranked.Add(new IndexedValue { Index = i, Value = values[i] });
        ranked.Sort(delegate(IndexedValue x, IndexedValue y)
        {
            int comparison = x.Value.CompareTo(y.Value);
            if (largest) comparison = -comparison;
            return comparison != 0 ? comparison : x.Index.CompareTo(y.Index);
        });
        object[,] result = new object[k, 2];
        for (int i = 0; i < k; i++)
        {
            result[i, 0] = ranked[i].Index + 1.0;
            result[i, 1] = ranked[i].Value;
        }
        return result;
    }

    //--------------------------------------------------------------------
    // 49. MATRIX_LINEAR_PREDICT
    //--------------------------------------------------------------------
    /// <summary>
    /// Applies a dense linear layer: observations multiplied by weights plus an optional bias.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_LINEAR_PREDICT", Description = "Dense linear predictions for row observations", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixLinearPredict(
        [ExcelArgument(Description = "Numeric matrix with rows as observations and columns as features")] object[,] matrix,
        [ExcelArgument(Description = "Weight matrix with features as rows and outputs as columns")] object[,] weights,
        [ExcelArgument(Description = "Optional scalar or output-length bias vector")] object biasObj)
    {
        double[,] inputValues, weightValues;
        if (!TryGetNumericMatrix(matrix, out inputValues) || !TryGetNumericMatrix(weights, out weightValues))
            return ExcelError.ExcelErrorValue;
        int observations = inputValues.GetLength(0);
        int features = inputValues.GetLength(1);
        int weightFeatures = weightValues.GetLength(0);
        int outputs = weightValues.GetLength(1);
        if (features != weightFeatures) return ExcelError.ExcelErrorValue;
        double[] bias;
        if (!TryGetBiasVector(biasObj, outputs, out bias)) return ExcelError.ExcelErrorValue;
        double[,] result = new double[observations, outputs];
        for (int r = 0; r < observations; r++)
            for (int o = 0; o < outputs; o++)
            {
                double sum = bias[o];
                for (int f = 0; f < features; f++) sum += inputValues[r, f] * weightValues[f, o];
                result[r, o] = sum;
            }
        return BuildNumericMatrix(result);
    }

    //--------------------------------------------------------------------
    // 50. MATRIX_CORRELATION
    //--------------------------------------------------------------------
    /// <summary>
    /// Returns the Pearson correlation matrix of feature columns, treating rows as observations.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_CORRELATION", Description = "Feature correlation matrix", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixCorrelation([ExcelArgument(Description = "Numeric matrix with rows as observations")] object[,] matrix)
    {
        double[,] values;
        if (!TryGetNumericMatrix(matrix, out values)) return ExcelError.ExcelErrorValue;
        int rows = values.GetLength(0), cols = values.GetLength(1);
        if (rows < 2) return ExcelError.ExcelErrorValue;
        double[] means = new double[cols];
        double[] sumSquares = new double[cols];
        for (int c = 0; c < cols; c++)
        {
            for (int r = 0; r < rows; r++) means[c] += values[r, c];
            means[c] /= rows;
            for (int r = 0; r < rows; r++)
            {
                double centered = values[r, c] - means[c];
                sumSquares[c] += centered * centered;
            }
            if (sumSquares[c] == 0.0) return ExcelError.ExcelErrorDiv0;
        }
        double[,] result = new double[cols, cols];
        for (int i = 0; i < cols; i++)
        {
            result[i, i] = 1.0;
            for (int j = i + 1; j < cols; j++)
            {
                double crossProduct = 0.0;
                for (int r = 0; r < rows; r++) crossProduct += (values[r, i] - means[i]) * (values[r, j] - means[j]);
                double correlation = crossProduct / Math.Sqrt(sumSquares[i] * sumSquares[j]);
                result[i, j] = correlation;
                result[j, i] = correlation;
            }
        }
        return BuildNumericMatrix(result);
    }

    //--------------------------------------------------------------------
    // 51. MATRIX_KMEANS_ASSIGN
    //--------------------------------------------------------------------
    /// <summary>
    /// Assigns each observation row to the nearest centroid and returns the 1-based centroid index and distance.
    /// </summary>
    [ExcelFunction(Name = "MATRIX_KMEANS_ASSIGN", Description = "Nearest-centroid assignments for observations", Category = "ExcelDNA ML & AI", IsThreadSafe = true)]
    public static object MatrixKMeansAssign(
        [ExcelArgument(Description = "Numeric matrix with rows as observations")] object[,] matrix,
        [ExcelArgument(Description = "Numeric matrix with rows as centroids")] object[,] centroids,
        [ExcelArgument(Description = "Optional metric: euclidean, manhattan, or cosine")] object metricObj)
    {
        double[,] values, centerValues;
        string metric;
        if (!TryGetNumericMatrix(matrix, out values) || !TryGetNumericMatrix(centroids, out centerValues) || !TryGetDistanceMetric(metricObj, out metric))
            return ExcelError.ExcelErrorValue;
        int rows = values.GetLength(0), features = values.GetLength(1);
        int centerCount = centerValues.GetLength(0), centerFeatures = centerValues.GetLength(1);
        if (features != centerFeatures) return ExcelError.ExcelErrorValue;
        double[] rowNorms = null, centerNorms = null;
        if (metric == "cosine")
        {
            rowNorms = new double[rows];
            centerNorms = new double[centerCount];
            for (int r = 0; r < rows; r++)
            {
                for (int f = 0; f < features; f++) rowNorms[r] += values[r, f] * values[r, f];
                rowNorms[r] = Math.Sqrt(rowNorms[r]);
                if (rowNorms[r] == 0.0) return ExcelError.ExcelErrorDiv0;
            }
            for (int c = 0; c < centerCount; c++)
            {
                for (int f = 0; f < features; f++) centerNorms[c] += centerValues[c, f] * centerValues[c, f];
                centerNorms[c] = Math.Sqrt(centerNorms[c]);
                if (centerNorms[c] == 0.0) return ExcelError.ExcelErrorDiv0;
            }
        }
        object[,] result = new object[rows, 2];
        for (int r = 0; r < rows; r++)
        {
            int bestCenter = 0;
            double bestDistance = double.PositiveInfinity;
            for (int c = 0; c < centerCount; c++)
            {
                double distance;
                if (metric == "manhattan")
                {
                    distance = 0.0;
                    for (int f = 0; f < features; f++) distance += Math.Abs(values[r, f] - centerValues[c, f]);
                }
                else if (metric == "cosine")
                {
                    double dot = 0.0;
                    for (int f = 0; f < features; f++) dot += values[r, f] * centerValues[c, f];
                    distance = 1.0 - dot / (rowNorms[r] * centerNorms[c]);
                    if (distance < 0.0 && distance > -1e-12) distance = 0.0;
                }
                else
                {
                    double sumSquares = 0.0;
                    for (int f = 0; f < features; f++)
                    {
                        double difference = values[r, f] - centerValues[c, f];
                        sumSquares += difference * difference;
                    }
                    distance = Math.Sqrt(sumSquares);
                }
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestCenter = c;
                }
            }
            result[r, 0] = bestCenter + 1.0;
            result[r, 1] = bestDistance;
        }
        return result;
    }

    private struct SubstringMatch
    {
        public int Start1;
        public int Start2;
        public int Length;
        public string Value;
    }

    private static List<SubstringMatch> GetCommonSubstringsByLongestMatch(string s1, string s2)
    {
        List<SubstringMatch> matches = new List<SubstringMatch>();
        AddLongestMatchRuns(s1, s2, 0, 0, matches);
        return matches;
    }

    private static void AddLongestMatchRuns(string s1, string s2, int offset1, int offset2, List<SubstringMatch> matches)
    {
        if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2)) return;

        int len1 = s1.Length;
        int len2 = s2.Length;
        int[,] dp = new int[len1 + 1, len2 + 1];
        int maxLen = 0;
        int end1 = 0;
        int end2 = 0;

        for (int i = 1; i <= len1; i++)
        {
            for (int j = 1; j <= len2; j++)
            {
                if (s1[i - 1] == s2[j - 1])
                {
                    int val = dp[i - 1, j - 1] + 1;
                    dp[i, j] = val;
                    if (val > maxLen)
                    {
                        maxLen = val;
                        end1 = i;
                        end2 = j;
                    }
                }
                else
                {
                    dp[i, j] = 0;
                }
            }
        }

        if (maxLen == 0) return;

        int start1 = end1 - maxLen;
        int start2 = end2 - maxLen;

        if (start1 > 0 && start2 > 0)
        {
            AddLongestMatchRuns(
                s1.Substring(0, start1),
                s2.Substring(0, start2),
                offset1,
                offset2,
                matches);
        }

        matches.Add(new SubstringMatch
        {
            Start1 = offset1 + start1,
            Start2 = offset2 + start2,
            Length = maxLen,
            Value = s1.Substring(start1, maxLen)
        });

        int nextStart1 = start1 + maxLen;
        int nextStart2 = start2 + maxLen;
        if (nextStart1 < len1 && nextStart2 < len2)
        {
            AddLongestMatchRuns(
                s1.Substring(nextStart1),
                s2.Substring(nextStart2),
                offset1 + nextStart1,
                offset2 + nextStart2,
                matches);
        }
    }

    private static List<string> CollectDiffs(string source, IEnumerable<SubstringMatch> matches, int minLength, Func<SubstringMatch, int> startSelector)
    {
        List<string> diffs = new List<string>();
        int current = 0;
        foreach (var match in matches)
        {
            int matchStart = startSelector(match);
            if (matchStart > current)
            {
                int length = matchStart - current;
                if (length >= minLength)
                {
                    diffs.Add(source.Substring(current, length));
                }
            }
            int matchEnd = matchStart + match.Length;
            if (matchEnd > current) current = matchEnd;
        }

        if (current < source.Length)
        {
            int length = source.Length - current;
            if (length >= minLength)
            {
                diffs.Add(source.Substring(current, length));
            }
        }

        return diffs;
    }

    private static object[,] BuildColumnArray(List<string> items)
    {
        if (items == null || items.Count == 0) return new object[0, 0];
        object[,] result = new object[items.Count, 1];
        for (int i = 0; i < items.Count; i++) result[i, 0] = items[i];
        return result;
    }


    private static int FindDelimiterOccurrence(string text, string delimiter, int instance)
    {
        int searchStart = 0;
        for (int occurrence = 1; occurrence <= instance; occurrence++)
        {
            int index = text.IndexOf(delimiter, searchStart, StringComparison.Ordinal);
            if (index < 0) return -1;
            if (occurrence == instance) return index;
            searchStart = index + delimiter.Length;
        }
        return -1;
    }

    private static bool TryGetOptionalPositiveInt(object arg, int defaultValue, out int value)
    {
        if (arg == null || arg is ExcelMissing || arg is ExcelEmpty)
        {
            value = defaultValue;
            return true;
        }
        if (!TryGetInt(arg, out value) || value < 1)
        {
            value = 0;
            return false;
        }
        return true;
    }

    private static bool GetOptionalBool(object arg, bool defaultValue)
    {
        if (arg == null || arg is ExcelMissing || arg is ExcelEmpty) return defaultValue;
        if (arg is bool) return (bool)arg;
        if (arg is double) return (double)arg != 0.0;
        bool parsed;
        return bool.TryParse(arg.ToString(), out parsed) ? parsed : defaultValue;
    }

    private static List<object> GetUniqueValues(object[,] inputArray, bool ignoreCase)
    {
        List<object> values = new List<object>();
        HashSet<string> seen = new HashSet<string>(ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        int rows = inputArray.GetLength(0), cols = inputArray.GetLength(1);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                object value = inputArray[r, c];
                if (value == null || value is ExcelEmpty || (value is string && ((string)value).Length == 0)) continue;
                if (seen.Add(BuildValueKey(value))) values.Add(value);
            }
        return values;
    }

    private static string BuildValueKey(object value)
    {
        if (value is ExcelError) return "ERROR:" + value.ToString();
        if (value is double) return "NUMBER:" + ((double)value).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        if (value is float) return "NUMBER:" + ((float)value).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        if (value is decimal) return "NUMBER:" + ((decimal)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (value is bool) return "BOOL:" + ((bool)value ? "1" : "0");
        if (value is DateTime) return "DATE:" + ((DateTime)value).ToOADate().ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        return "TEXT:" + value.ToString();
    }

    private static object[,] BuildObjectColumnArray(List<object> items)
    {
        if (items == null || items.Count == 0) return new object[0, 0];
        object[,] result = new object[items.Count, 1];
        for (int i = 0; i < items.Count; i++) result[i, 0] = items[i];
        return result;
    }

    private static bool TryGetNumericVector(object[,] input, out double[] values, out int rows, out int cols)
    {
        values = null;
        rows = 0;
        cols = 0;
        if (input == null) return false;
        rows = input.GetLength(0);
        cols = input.GetLength(1);
        if (rows < 1 || cols < 1 || (rows != 1 && cols != 1)) return false;
        values = new double[rows * cols];
        int index = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                double value;
                if (!TryGetDouble(input[r, c], out value))
                {
                    values = null;
                    return false;
                }
                values[index++] = value;
            }
        return true;
    }

    private static bool TryGetNumericMatrix(object[,] input, out double[,] values)
    {
        values = null;
        if (input == null) return false;
        int rows = input.GetLength(0), cols = input.GetLength(1);
        if (rows < 1 || cols < 1) return false;
        values = new double[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (!TryGetDouble(input[r, c], out values[r, c]))
                {
                    values = null;
                    return false;
                }
        return true;
    }

    private static object[,] BuildNumericVector(double[] values, int rows, int cols)
    {
        object[,] result = new object[rows, cols];
        int index = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++) result[r, c] = values[index++];
        return result;
    }

    private static object[,] BuildNumericMatrix(double[,] values)
    {
        int rows = values.GetLength(0), cols = values.GetLength(1);
        object[,] result = new object[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++) result[r, c] = values[r, c];
        return result;
    }

    private static double ComputeVectorNorm(double[] values, double p)
    {
        if (double.IsPositiveInfinity(p))
        {
            double max = 0.0;
            for (int i = 0; i < values.Length; i++) max = Math.Max(max, Math.Abs(values[i]));
            return max;
        }
        double sum = 0.0;
        for (int i = 0; i < values.Length; i++) sum += Math.Pow(Math.Abs(values[i]), p);
        return Math.Pow(sum, 1.0 / p);
    }

    private static bool TryGetOptionalDouble(object arg, double defaultValue, out double value)
    {
        if (arg == null || arg is ExcelMissing || arg is ExcelEmpty)
        {
            value = defaultValue;
            return true;
        }
        return TryGetDouble(arg, out value);
    }

    private static bool TryGetDistanceMetric(object arg, out string metric)
    {
        metric = "euclidean";
        if (arg == null || arg is ExcelMissing || arg is ExcelEmpty) return true;
        metric = arg.ToString().Trim().ToLowerInvariant();
        return metric == "euclidean" || metric == "manhattan" || metric == "cosine";
    }

    private static bool IsBlankValue(object value)
    {
        return value == null || value is ExcelMissing || value is ExcelEmpty || (value is string && string.IsNullOrWhiteSpace((string)value));
    }

    private static bool TryGetLabelVector(object[,] input, out List<object> labels)
    {
        labels = new List<object>();
        if (input == null) return false;
        int rows = input.GetLength(0), cols = input.GetLength(1);
        if (rows < 1 || cols < 1 || (rows != 1 && cols != 1)) return false;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                object value = input[r, c];
                if (IsBlankValue(value) || value is ExcelError) return false;
                labels.Add(value);
            }
        return labels.Count > 0;
    }

    private static bool TryGetClassLabels(object classLabelsObj, IEnumerable<object> inferredValues, out List<object> classes)
    {
        classes = new List<object>();
        HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
        object[,] classArray = classLabelsObj as object[,];
        if (classArray != null)
        {
            List<object> explicitLabels;
            if (!TryGetLabelVector(classArray, out explicitLabels)) return false;
            foreach (object value in explicitLabels)
            {
                string key = BuildValueKey(value);
                if (!seen.Add(key)) return false;
                classes.Add(value);
            }
            return classes.Count > 0;
        }
        if (!(classLabelsObj == null || classLabelsObj is ExcelMissing || classLabelsObj is ExcelEmpty))
        {
            if (IsBlankValue(classLabelsObj) || classLabelsObj is ExcelError) return false;
            classes.Add(classLabelsObj);
            return true;
        }
        foreach (object value in inferredValues)
        {
            string key = BuildValueKey(value);
            if (seen.Add(key)) classes.Add(value);
        }
        return classes.Count > 0;
    }

    private static Dictionary<string, int> BuildLabelIndex(List<object> classes)
    {
        Dictionary<string, int> index = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < classes.Count; i++) index[BuildValueKey(classes[i])] = i;
        return index;
    }

    private struct IndexedValue
    {
        public int Index;
        public double Value;
    }

    private static bool TryGetBiasVector(object biasObj, int outputs, out double[] bias)
    {
        bias = new double[outputs];
        if (biasObj == null || biasObj is ExcelMissing || biasObj is ExcelEmpty) return true;
        double scalar;
        if (TryGetDouble(biasObj, out scalar))
        {
            for (int i = 0; i < outputs; i++) bias[i] = scalar;
            return true;
        }
        object[,] biasArray = biasObj as object[,];
        double[] values;
        int rows, cols;
        if (biasArray == null || !TryGetNumericVector(biasArray, out values, out rows, out cols) || values.Length != outputs)
            return false;
        for (int i = 0; i < outputs; i++) bias[i] = values[i];
        return true;
    }

    private static bool TryGetDouble(object arg, out double value)
    {
        if (arg is double)
        {
            value = (double)arg;
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
        if (arg is int) { value = (int)arg; return true; }
        if (arg is long) { value = (long)arg; return true; }
        if (arg is decimal) { value = (double)(decimal)arg; return true; }
        if (arg is string)
            return double.TryParse((string)arg, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out value);
        value = 0.0;
        return false;
    }


    // Helper to robustly parse Excel argument to int (handles double, string, etc.)
    private static bool TryGetInt(object arg, out int value)
    {
        if (arg is int)
        {
            value = (int)arg;
            return true;
        }
        if (arg is double)
        {
            double d = (double)arg;
            if (d % 1 == 0 && d >= int.MinValue && d <= int.MaxValue)
            {
                value = (int)d;
                return true;
            }
        }
        if (arg is string)
        {
            int parsed;
            if (int.TryParse((string)arg, out parsed))
            {
                value = parsed;
                return true;
            }
        }
        value = 0;
        return false;
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
