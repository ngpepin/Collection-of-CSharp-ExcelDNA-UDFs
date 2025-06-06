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
 * Notes:
 * - Functions marked as volatile recalculate when any cell changes
 * - Stateful functions (like INJECTVALUE) maintain state between calculations
 * - Temporary storage persists until PURGEOBJECTS() is called or workbook closes
 * - Thread management requires Excel 2007 or later
 *
 */

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

public class C
{
    // Version components
    private const string VERSION_MAJOR = "3";
    private const string VERSION_MINOR = "1";
    private const string VERSION_PATCH = "1";

    // Current version string
    private const string CurrentVersion =
        VERSION_MAJOR + "." + VERSION_MINOR + "." + VERSION_PATCH;

    private static string _targetVersion = CurrentVersion;

    public static string Version
    {
        get { return CurrentVersion; }
    }

    public static string TargetVersion
    {
        get { return _targetVersion; }
        set
        {
            // Validate version format (simple check for x.y.z pattern)
            if (System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d+\.\d+\.\d+$"))
            {
                _targetVersion = value;
            }
        }
    }
    // Usage Examples in Other UDFs:
    //
    // [ExcelFunction(Description = "Example UDF with version-aware behavior")]
    // public static object VersionAwareFunction()
    // {
    //     // Compare versions using standard string comparison
    //     if (string.Compare(TargetVersion, "2.0.0") < 0)
    //     {
    //         // Legacy behavior for pre-2.0 versions
    //         return LegacyImplementation();
    //     }
    //     else
    //     {
    //         // Current behavior
    //         return CurrentImplementation();
    //     }
    // }

    private static Dictionary<string, object> objectStore = new Dictionary<string, object>();
    private static Dictionary<string, object> injectedCells = new Dictionary<string, object>();
    private static Dictionary<string, Tuple<object, object>> invocationCache =
        new Dictionary<string, Tuple<object, object>>();
    private static Dictionary<string, Tuple<object, object>> visibilityCache =
        new Dictionary<string, Tuple<object, object>>();
    private static Excel.Application _excelApp;
    private static Excel.Application _app;
    const int defCachingTime = 10; // seconds

    // This is a helper method to manage caching time for the IsVisible UDF.
    public static void AttachEvents()
    {
        _excelApp = (Excel.Application)ExcelDnaUtil.Application;
        _excelApp.WorkbookBeforeClose += WorkbookBeforeClose;
    }

    // This is a helper method to handle cleanup actions when the workbook is closed.
    private static void WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
    {
        // Perform cleanup actions here
        Cleanup();
    }

    // This is a helper method to detach events when the add-in is unloaded.
    public static void DetachEvents()
    {
        if (_excelApp != null)
        {
            _excelApp.WorkbookBeforeClose -= WorkbookBeforeClose;
        }
    }

    // This is a helper method to get the Excel application instance.
    public static Excel.Application App
    {
        get
        {
            if (_app == null)
                _app = (Excel.Application)ExcelDnaUtil.Application;
            return _app;
        }
    }

    // Call this helper method to clean up when the add-in is unloaded or when the Excel application is closed.
    // It releases the COM object and sets it to null. This is important to avoid memory leaks and ensure proper cleanup.
    public static void Cleanup()
    {
        if (_app != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
            _app = null;
        }
    }

    //
    // UDFS START HERE
    //
    // vExcelDNA UDF
    // ------------------------------------------------------------------------------------
    //
    // Returns the current version of the UDF collection
    [ExcelFunction(
        Name = "vExcelDNA",
        Description = "Returns the version of the Excel-DNA UDF collection",
        Category = "ExcelDNA Utilities",
        IsVolatile = false
    )]
    public static string GetExcelDnaVersion()
    {
        return CurrentVersion;
    }

    //
    // SetTargetVersion UDF
    // ------------------------------------------------------------------------------------
    //
    // Sets the compatibility target version for backward compatibility
    [ExcelFunction(
        Name = "SetTargetVersion",
        Description = "Sets the target version for backward compatibility",
        Category = "ExcelDNA Utilities",
        IsMacroType = true
    )]
    public static string SetTargetVersion(
        [ExcelArgument(Description = "Target version in x.y.z format")] string version
    )
    {
        string previousVersion = TargetVersion;
        TargetVersion = version;
        return "Target version changed from " + previousVersion + " to " + TargetVersion;
    }

    //
    // GetTargetVersion UDF
    // ------------------------------------------------------------------------------------
    //
    // Gets the current compatibility target version
    [ExcelFunction(
        Name = "GetTargetVersion",
        Description = "Gets the current target version for backward compatibility",
        Category = "ExcelDNA Utilities",
        IsVolatile = false
    )]
    public static string GetTargetVersion()
    {
        return TargetVersion;
    }

    // RecalcAll UDF
    // ------------------------------------------------------------------------------------
    //
    // This function triggers a full recalculation of the workbook.
    // It uses the Excel application object to perform the recalculation.
    // The function is marked as a UDF (User Defined Function) and can be called from Excel.

    [ExcelFunction(Description = "Triggers full recalculation of the workbook")]
    public static object RecalcAll()
    {
        try
        {
            // Verify we have a valid Excel application instance
            var excelApp = (Excel.Application)ExcelDnaUtil.Application;
            if (excelApp == null)
                return ExcelError.ExcelErrorValue;

            // Queue as macro to ensure proper execution context
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    excelApp.CalculateFull();
                }
                catch (Exception)
                {
                    // Silently handle any calculation errors
                }
            });

            return "TRUE";
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception)
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    //
    // GetIterationStatus UDF
    // ------------------------------------------------------------------------------------
    //
    // This function retrieves the status of Excel's iterative calculations mode.
    // It returns a string indicating whether the mode is ON or OFF,
    // along with the maximum number of iterations and maximum change between iterations.
    // The function uses the Excel application object to access the properties
    // related to iterative calculations.
    // It also handles exceptions and returns an error message if an exception occurs.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.

    [ExcelFunction(
        Description = "Returns the status of Excel's iterative calculations mode.",
        IsVolatile = true
    )]
    public static string GetIterationStatus()
    {
        bool isIterationOn = false;
        int maxIterations = 9999;
        double maxChange = 0.9999;
        string infoMessage = "";
        string stateMsg = "";

        try
        {
            isIterationOn = App.Iteration;
            maxIterations = App.MaxIterations;
            maxChange = App.MaxChange;

            if (isIterationOn)
                stateMsg = "ON";
            else
                stateMsg = "OFF";

            string maxChange_str = maxChange.ToString();
            infoMessage =
                "Status: "
                + stateMsg
                + "  Max Iterations: "
                + maxIterations.ToString()
                + "  Max Change: "
                + maxChange.ToString();
        }
        catch (Exception ex)
        {
            return ex.Message;
        }

        return infoMessage;
    }

    //
    // SetIteration UDF
    // ------------------------------------------------------------------------------------
    //
    // This function enables or disables Excel's iterative calculations mode
    // and sets the maximum number of iterations and maximum change between iterations.
    // It takes three parameters: IterationOn (boolean), maxIterations (integer),
    // and maxChange (double).
    // The function checks if the parameters are valid and sets the corresponding properties
    // in the Excel application.
    // It also handles exceptions and returns an error message if an exception occurs.
    // The function returns a confirmation message indicating the status of the iterative calculations mode.

    [ExcelFunction(
        Description = "Enables or disables Excel iterative calculations mode and sets parameters."
    )]
    public static string SetIteration(
        [ExcelArgument(Description = "The state of the iterative calculation option.")]
            bool IterationOn = false,
        [ExcelArgument(Description = "The maximum number of iterations.")] int maxIterations = 100,
        [ExcelArgument(Description = "The maximum change between iterations.")]
            double maxChange = 0.001
    )
    {
        double myChange = 0.001;
        int myIterations = 100;

        try
        {
            App.Iteration = IterationOn;

            myIterations = maxIterations;
            if (myIterations <= 0)
                myIterations = 100;
            App.MaxIterations = myIterations;

            myChange = maxChange;
            if ((myChange <= 0.0) || (myChange >= 1.0))
                myChange = 0.001;
            App.MaxChange = myChange;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }

        string confirmationMsg = GetIterationStatus();
        return confirmationMsg;
    }

    [ExcelFunction(
        Description = "Returns TRUE if the cell is visible",
        IsVolatile = true,
        IsMacroType = true
    )]


    //
    // IsVisible UDF
    // ------------------------------------------------------------------------------------
    //
    // This function checks if a cell is visible in the Excel worksheet.
    // It takes one parameter: the caching time in seconds (default is 10 seconds).
    // The function uses a dictionary to cache the visibility status of cells.
    // It checks if the cell is hidden in either the row or column.
    // If the cell is visible, it returns "TRUE", otherwise it returns "FALSE".
    // The function also handles exceptions and returns an error message if an exception occurs.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.

    public static object IsVisible(
        [ExcelArgument(Description = "The caching time in seconds (default is 10 seconds).")]
            int cachingTime = defCachingTime
    )
    {
        Excel.Range range = null;
        DateTime lastChecked = DateTime.MinValue;
        bool isVis = false;
        bool RetVis = false;

        try
        {
            var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (caller == null)
                return ExcelDna.Integration.ExcelError.ExcelErrorRef;

            string address = (string)XlCall.Excel(XlCall.xlfReftext, caller, true);

            if (visibilityCache.ContainsKey(address))
            {
                Tuple<object, object> tuple = visibilityCache[address];
                lastChecked = (DateTime)tuple.Item1;
                isVis = (bool)tuple.Item2;

                if ((DateTime.Now - lastChecked).TotalSeconds < cachingTime)
                {
                    // Return cached result if last check was less than 5 seconds ago
                    if (isVis)
                        return "TRUE";
                    else
                        return "FALSE";
                }
            }

            // If not cached or cache is outdated, check visibility (involves COM call)
            range = App.Range[address];
            if (range == null)
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;

            RetVis = !((bool)range.EntireRow.Hidden || (bool)range.EntireColumn.Hidden);

            // Update the cache
            visibilityCache[address] = new Tuple<object, object>(DateTime.Now, RetVis);

            if (RetVis)
                return "TRUE";
            else
                return "FALSE";
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    //
    // Describe UDF
    // ------------------------------------------------------------------------------------
    //
    // This function describes the value passed to the function or contained within a referenced cell.
    // It takes one parameter: the cell for which the value should be described.
    // The function checks the type of the value and returns a string description of it.
    // The function handles different types of values, including double, string, boolean,
    // ExcelError, object array, ExcelMissing, ExcelEmpty, and ExcelReference.
    // The function also handles cases where the value is null or of an unexpected type.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.

    [ExcelFunction(
        Description = "Describes the value passed to the function or contained within a referenced cell.",
        IsMacroType = true
    )]
    public static string Describe( // [ExcelArgument(AllowReference = false)] object arg)
        [ExcelArgument(Description = "The cell for which the value should be described.")] object arg
    )
    {
        if (arg is double)
            return "Double: " + (double)arg;
        else if (arg is string)
            return "String: " + (string)arg;
        else if (arg is bool)
            return "Boolean: " + (bool)arg;
        else if (arg is ExcelError)
            return "ExcelError: " + arg.ToString();
        else if (arg is object[,])
            // The object array returned here may contain a mixture of different types,
            // reflecting the different cell contents.
            return string.Format(
                "Array[{0},{1}]",
                ((object[,])arg).GetLength(0),
                ((object[,])arg).GetLength(1)
            );
        else if (arg is ExcelMissing)
            return "Missing";
        else if (arg is ExcelEmpty)
            return "Empty";
        else if (arg is ExcelReference)
            return "Reference: " + XlCall.Excel(XlCall.xlfReftext, arg, true);
        else
            return "!?Unheard Of";
    }

    //
    // InjectValue UDF
    // ------------------------------------------------------------------------------------
    //
    // This function injects a specified value into a cell if it has not already been injected
    // during the current calculation session.
    // It takes two parameters: the cell reference where the value will be injected
    // and the value to inject into the cell.
    // The function checks if the cell reference and value are valid and returns an error if not.
    // It also checks if the cell reference is a valid ExcelReference and retrieves the cell's address.
    // The function uses a dictionary to keep track of injected cells and their values.
    // If the cell has already been injected with the same value, it returns the existing value.
    // If the cell has not been injected or has a different value, it injects the new value
    // and updates the dictionary.
    // The function handles exceptions and returns an error message if an exception occurs.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.

    [ExcelFunction(
        Description = "Injects a specified value into a cell if not already done during this calculation session"
    )]
    public static object InjectValue(
        [ExcelArgument(
            Description = "The cell reference where the value will be injected",
            AllowReference = true
        )]
            object potentialRef,
        [ExcelArgument(Description = "The value to inject into the cell")] object value
    )
    {
        string cellKey = null;

        if (potentialRef == null || value == null)
            return ExcelDna.Integration.ExcelError.ExcelErrorValue;

        try
        {
            ExcelReference cellRef = potentialRef as ExcelReference;
            if (cellRef == null)
                return "Error: The first argument must be a cell reference.";

            try
            {
                string cellReference = (string)
                    XlCall.Excel(XlCall.xlfAddress, 1 + cellRef.RowFirst, 1 + cellRef.ColumnFirst);
                long sheetId = (long)cellRef.SheetId;
                cellKey = sheetId + "!" + cellReference;
            }
            catch // (Exception ex)
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
                // return "Error: An exception occurred - " + ex.Message;
            }

            // Default return value
            object[,] save_values = new object[1, 1]
            {
                { value }
            };

            if (injectedCells.ContainsKey(cellKey))
            {
                if (injectedCells[cellKey].Equals(value))
                    // return cellKey;
                    return save_values; // No change, so return the value
            }

            // Inject the value
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                object[,] values = new object[1, 1]
                {
                    { value }
                };
                if (cellRef.GetValue() != values)
                {
                    cellRef.SetValue(values);

                    // Record the injection
                    injectedCells[cellKey] = value;
                }
                save_values = values;
            });

            return save_values;
        }
        catch (System.Exception ex)
        {
            return "Failed to inject value due to: " + ex.Message;
        }
    }

    //
    // FindPos UDF
    // ------------------------------------------------------------------------------------
    //
    // This function finds the position of a substring within a text string.
    // It takes three parameters: the text to search, the substring to find,
    // and the occurrence instance to find.
    // The function returns the position of the substring in the text string.
    // If the substring is not found, it returns an error.
    // The function handles case-insensitive searches by converting both the text and substring to lowercase.
    // It also handles the case where the instance is -1, which returns the last occurrence of the substring.
    // The function uses a list to store the indices of all occurrences of the substring.
    // The function checks if the text and substring are null or empty and returns an error if they are.
    // The function also checks if the instance is out of range and returns an error if it is.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.
    // The function uses the IndexOf method to find the position of the substring within the text string.
    // The function also handles exceptions and returns an error message if an exception occurs.
    // The function is marked as a UDF (User Defined Function) and can be called from Excel.
    // The function is also marked as a macro type, meaning it can be called from a macro.

    [ExcelFunction(Description = "Returns the position of a substring within a text string.")]
    public static object FindPos(
        [ExcelArgument(Description = "The text to search")] string text,
        [ExcelArgument(Description = "The substring to find")] string substring,
        [ExcelArgument(Description = "The occurrence instance to find")] int instance
    )
    {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(substring))
            return ExcelDna.Integration.ExcelError.ExcelErrorValue;

        // Normalize the string and substring to lowercase for case-insensitive comparison.
        string lowerText = text.ToLower();
        string lowerSubstring = substring.ToLower();
        var indices = new List<int>();

        // Find all occurrences of the substring.
        int pos = lowerText.IndexOf(lowerSubstring);
        while (pos != -1)
        {
            indices.Add(pos + 1); // Convert 0-based index to 1-based for Excel compatibility
            pos = lowerText.IndexOf(lowerSubstring, pos + 1);
        }

        // Return the appropriate position based on instance value.
        if (instance == -1)
        {
            if (indices.Count == 0)
                return ExcelDna.Integration.ExcelError.ExcelErrorValue; // Return error if no instances found
            return indices.Last(); // Return the last instance
        }
        else if (instance > 0 && instance <= indices.Count)
        {
            return indices[instance - 1]; // Return the nth instance
        }
        else
        {
            return ExcelDna.Integration.ExcelError.ExcelErrorValue; // Return error if instance is out of range
        }
    }

    //
    // PutObject UDF
    // ------------------------------------------------------------------------------------
    //
    // This function stores an object in a temporary storage dictionary.
    // It takes a name, the object to store, and optional parameters for overwriting
    // existing objects and displaying debug information.
    // The function checks if the name is valid and if the object already exists in the cache.
    // If the object already exists and force is set to false, it returns an error message.
    // If the object is successfully stored, it returns the stored object.
    // The function also handles exceptions and returns an error message if an exception occurs.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.
    // The function uses a dictionary to store the objects, allowing for efficient retrieval.
    // The function also uses a cache to avoid redundant writes to the same key with the same value.
    // The function also checks if the caller reference matches the one in the cache to avoid conflicts.
    // The function uses a tuple to store the caller reference and the value in the cache.
    // The function also handles cases where the name is empty or null,
    // and provides an option to display debug information.

    [ExcelFunction(Description = "Enters an object into temporary storage.", IsVolatile = true)]
    public static object PutObject(
        [ExcelArgument(Description = "The name of the object to store.")] string name,
        [ExcelArgument(Description = "The object to store.")] object value,
        [ExcelArgument(Description = "Overwrite the object if it already exists.")]
            bool force = true,
        [ExcelArgument(Description = "Display debug information.")] bool debug = false
    )
    {
        string CacheKey = null;
        string callerReference = null;
        object caller = XlCall.Excel(XlCall.xlfCaller);

        try
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                if (debug)
                    return "Error: Name parameter cannot be empty.";
                else
                    return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }

            try
            {
                //
                // Check if this exact call has been made before and ignore if it has
                //
                ExcelReference caller_er = caller as ExcelReference;
                callerReference = (string)
                    XlCall.Excel(
                        XlCall.xlfAddress,
                        1 + caller_er.RowFirst,
                        1 + caller_er.ColumnFirst
                    );
                CacheKey = callerReference + ":" + name;

                if (invocationCache.ContainsKey(CacheKey))
                {
                    Tuple<object, object> tuple = invocationCache[CacheKey];
                    object inv_caller = tuple.Item1;
                    string inv_callerReference = (string)inv_caller;
                    object inv_value = tuple.Item2;

                    if (inv_callerReference != callerReference)
                    {
                        if (debug)
                            return "Caller does not match the caller in the keystore.";
                        else
                            return ExcelDna.Integration.ExcelError.ExcelErrorName;
                    }
                    else
                    {
                        if (Equals(inv_value, value))
                        {
                            // If cell and value are the same, skip writing
                            if (debug)
                                return "This is a redundant write to the same key of the same value, ignoring.";
                            else
                                return value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "Error: An exception occurred - " + ex.Message;
            }
            // Update or add to invocation cache so we know no to keep re-writing the same value
            invocationCache[CacheKey] = new Tuple<object, object>(callerReference, value);

            //
            // Store the object into the object store
            //
            if (objectStore.ContainsKey(name) && !force)
            {
                if (debug)
                    return "Error: Object with this name already exists. Set 'force' to TRUE to overwrite.";
                else
                    return ExcelDna.Integration.ExcelError.ExcelErrorName;
            }
            objectStore[name] = value; // This will add or overwrite if force is TRUE

            return value; // Indicate successful storage of object
        }
        catch (Exception ex)
        {
            return "Error: An exception occurred - " + ex.Message;
        }
    }

    //
    // GetObject UDF
    // ------------------------------------------------------------------------------------
    //
    // This function retrieves an object from temporary storage based on its name.
    // It returns the object if found, or an error message if not found.
    // The function also handles cases where the name is empty or null,
    // and provides an option to display debug information.
    // The function is marked as volatile, meaning it will recalculate whenever any cell in the workbook changes.
    // The function also handles exceptions and returns an error message if an exception occurs.
    // The function uses a dictionary to store the objects, allowing for efficient retrieval.

    [ExcelFunction(Description = "Retrieves an object from temporary storage.", IsVolatile = true)]
    public static object GetObject(
        [ExcelArgument(Description = "The name of the object to retrieve.")] string name,
        [ExcelArgument(Description = "Display debug information.")] bool debug = false
    )
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                if (debug)
                    return "Error: Name parameter cannot be empty.";
                else
                    return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }

            if (!objectStore.ContainsKey(name))
            {
                if (debug)
                    return "Error: No object found with this name.";
                else
                    return ExcelDna.Integration.ExcelError.ExcelErrorName;
            }

            return objectStore[name] ?? "Error: Stored object is null.";
        }
        catch (Exception ex)
        {
            return "Error: An exception occurred -" + ex.Message;
        }
    }

    //
    // PurgeObjects UDF
    // ------------------------------------------------------------------------------------
    //

    [ExcelFunction(Description = "Purges all stored objects.")]
    public static string PurgeObjects()
    {
        objectStore.Clear(); // Clears the dictionary holding the objects
        return "TRUE";
    }

    //
    // TrueSplit UDF
    // ------------------------------------------------------------------------------------
    //
    // This function splits strings by a specified delimiter and returns a dynamic array of substrings.
    // It handles both single strings and arrays of strings.
    // It also handles Excel-specific types like ExcelEmpty and ExcelError.
    // The function returns a 2D array where each row corresponds to an input string,
    // and each column corresponds to a substring obtained by splitting the input string.
    // If the input is null or empty, it returns an empty string.
    // If the input is an error, it returns the error in the first column and empty strings in the rest.
    // The function also handles cases where the number of substrings varies across input strings,
    // filling remaining columns with empty strings as needed.


    [ExcelFunction(
        Description = "Splits strings by delimiter and returns a dynamic array of substrings"
    )]
    public static object[,] TrueSplit(
        [ExcelArgument(Description = "Range or array of strings to split")] object[] inputStrings,
        [ExcelArgument(Description = "Delimiter to split by")] string delimiter
    )
    {
        try
        {
            // Determine the maximum number of columns needed
            int maxColumns = 1;
            foreach (var item in inputStrings)
            {
                string str = item as string;
                if (str != null)
                {
                    int count = str.Split(
                        new string[] { delimiter },
                        StringSplitOptions.None
                    ).Length;
                    maxColumns = Math.Max(maxColumns, count);
                }
                else if (item == null || item is ExcelDna.Integration.ExcelEmpty)
                {
                    maxColumns = Math.Max(maxColumns, 1); // Handle null or empty as single column
                }
                else if (item is ExcelDna.Integration.ExcelError)
                {
                    maxColumns = Math.Max(maxColumns, 1); // Handle errors as single column
                }
            }

            // Create the output array
            object[,] result = new object[inputStrings.Length, maxColumns];

            for (int i = 0; i < inputStrings.Length; i++)
            {
                string str = inputStrings[i] as string;
                if (str != null)
                {
                    string[] parts = str.Split(new string[] { delimiter }, StringSplitOptions.None);
                    for (int j = 0; j < parts.Length; j++)
                    {
                        result[i, j] = parts[j];
                    }

                    // Fill remaining columns with empty strings if needed
                    for (int j = parts.Length; j < maxColumns; j++)
                    {
                        result[i, j] = "";
                    }
                }
                else if (
                    inputStrings[i] == null
                    || inputStrings[i] is ExcelDna.Integration.ExcelEmpty
                )
                {
                    result[i, 0] = ""; // First column gets empty string
                    for (int j = 1; j < maxColumns; j++)
                    {
                        result[i, j] = "";
                    }
                }
                else if (inputStrings[i] is ExcelDna.Integration.ExcelError)
                {
                    result[i, 0] = inputStrings[i]; // First column gets the error
                    for (int j = 1; j < maxColumns; j++)
                    {
                        result[i, j] = "";
                    }
                }
                else
                {
                    // Handle other unexpected types (shouldn't normally happen)
                    result[i, 0] = inputStrings[i].ToString();
                    for (int j = 1; j < maxColumns; j++)
                    {
                        result[i, j] = "";
                    }
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,]
            {
                { ExcelError.ExcelErrorValue }
            };
        }
    }

    //
    // IsMemberOf UDF
    // ------------------------------------------------------------------------------------
    //
    // This function checks if any element/row/column from the first array exists in the second array.
    // It returns TRUE if a match is found, otherwise FALSE.
    // It handles both single-element arrays (1x1) and multi-dimensional arrays.
    // It also handles Excel-specific types like ExcelEmpty and ExcelError.

    [ExcelFunction(
        Description = "Checks if any element/row/column from first array exists in second array"
    )]
    public static bool IsMemberOf(
        [ExcelArgument(Description = "First array (will be searched in second array)")]
            object[,] arrayA,
        [ExcelArgument(Description = "Second array (will be searched against)")] object[,] arrayB
    )
    {
        try
        {
            // Get dimensions of both arrays
            int aRows = arrayA.GetLength(0);
            int aCols = arrayA.GetLength(1);
            int bRows = arrayB.GetLength(0);
            int bCols = arrayB.GetLength(1);

            // Handle single-element arrays (1x1)
            bool aIsSingle = (aRows == 1 && aCols == 1);
            bool bIsSingle = (bRows == 1 && bCols == 1);

            // Special case: if either array is single-element, compare as values
            if (aIsSingle || bIsSingle)
            {
                object aValue = arrayA[0, 0];
                if (bIsSingle)
                {
                    return AreEqual(aValue, arrayB[0, 0]);
                }
                else
                {
                    // Search single value in arrayB
                    for (int i = 0; i < bRows; i++)
                    {
                        for (int j = 0; j < bCols; j++)
                        {
                            if (AreEqual(aValue, arrayB[i, j]))
                            {
                                return true;
                            }
                        }
                    }
                    return false;
                }
            }

            // Determine if we're comparing rows or columns
            bool compareRows = (aCols == bCols); // If column counts match, compare rows
            bool compareCols = (aRows == bRows); // If row counts match, compare columns

            if (!compareRows && !compareCols)
            {
                // Arrays have incompatible dimensions for comparison
                return false;
            }

            if (compareRows)
            {
                // Compare rows of A against rows of B
                for (int aRow = 0; aRow < aRows; aRow++)
                {
                    for (int bRow = 0; bRow < bRows; bRow++)
                    {
                        bool match = true;
                        for (int col = 0; col < aCols; col++)
                        {
                            if (!AreEqual(arrayA[aRow, col], arrayB[bRow, col]))
                            {
                                match = false;
                                break;
                            }
                        }
                        if (match)
                            return true;
                    }
                }
            }

            if (compareCols)
            {
                // Compare columns of A against columns of B
                for (int aCol = 0; aCol < aCols; aCol++)
                {
                    for (int bCol = 0; bCol < bCols; bCol++)
                    {
                        bool match = true;
                        for (int row = 0; row < aRows; row++)
                        {
                            if (!AreEqual(arrayA[row, aCol], arrayB[row, bCol]))
                            {
                                match = false;
                                break;
                            }
                        }
                        if (match)
                            return true;
                    }
                }
            }

            return false;
        }
        catch (Exception)
        {
            return false;
        }
    }

    [ExcelFunction(
     Name = "GetThreads",
     Description = "Gets Excel's current multithreading calculation thread count",
     Category = "ExcelDNA Utilities",
     IsVolatile = true
    )]
    public static object GetThreads()
    {
        try
        {
            // Verify Excel version supports multithreading (Excel 2007+)
            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (new Version(app.Version) < new Version("12.0"))
                return "Excel 2007+ required";

            // Access multithreading properties
            var mtc = app.MultiThreadedCalculation;
            // int maxThreads = Environment.ProcessorCount;
            int maxThreads = 64;

            return new object[,]
            {
            { "Current Thread Count", mtc.ThreadCount },
            { "Max Available", maxThreads },
            { "Mode Enabled", mtc.Enabled }
            };
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return ExcelError.ExcelErrorNA; // #N/A if Excel not available
        }
        catch (Exception)
        {
            return ExcelError.ExcelErrorValue; // #VALUE! for other errors
        }
    }

    //
    // SetThreads UDF
    // ------------------------------------------------------------------------------------
    // This function sets Excel's multithreading calculation settings.
    // It takes two parameters: threadCount (number of threads to use) and enable (boolean to enable/disable multithreading).
    // The function validates the threadCount and enables/disables multithreading accordingly.
    // It also handles exceptions and returns an error message if an exception occurs.
    // The function uses a lock to ensure thread safety when accessing shared resources.
    // The function caches the last thread count and enabled state to avoid redundant calls.
    // The function uses ExcelAsyncUtil to queue the operation as a macro.
    // The function also checks if the settings are the same as the last call to avoid unnecessary updates.
    // The function returns a confirmation message indicating the status of the multithreading settings.

    private static int _lastThreadCount = -2; // Initialize with invalid value
    private static bool _lastThreadEnabled;
    private static readonly object _threadLock = new object();

    [ExcelFunction(
        Name = "SetThreads",
        Description = "Configures Excel's multithreading settings",
        Category = "ExcelDNA Utilities",
        IsMacroType = true
    )]
    public static object SetThreads(
        [ExcelArgument(Description = "Number of threads (0=Auto, -1=Max)")]
    int threadCount,
        [ExcelArgument(Description = "Enable multithreading")]
    bool enable = true)
    {
        lock (_threadLock)
        {
            try
            {
                // Check if settings are the same as last call
                if (_lastThreadCount == threadCount && _lastThreadEnabled == enable)
                {
                    return "Using cached thread settings";
                }

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;

                    // Validate Excel version
                    if (new Version(app.Version) < new Version("12.0"))
                    {
                        return;
                    }

                    var mtc = app.MultiThreadedCalculation;
                    // int maxThreads = Environment.ProcessorCount;
                    int maxThreads = 64;
                    int newCount;

                    // Determine thread count
                    if (threadCount == -1)
                    {
                        newCount = maxThreads; // Use all processors
                    }
                    else if (threadCount == 0)
                    {
                        newCount = maxThreads / 2; // Automatic (half of max)
                    }
                    else
                    {
                        newCount = Math.Min(threadCount, maxThreads);
                    }

                    // Only update if settings actually changed
                    if (mtc.ThreadCount != newCount || mtc.Enabled != enable)
                    {
                        mtc.ThreadCount = newCount;
                        mtc.Enabled = enable;
                        _lastThreadCount = threadCount;
                        _lastThreadEnabled = enable;

                        if (enable)
                        {
                            app.CalculateFullRebuild();
                        }
                    }
                });

                return "Thread settings updated";
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return ExcelError.ExcelErrorNA;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }
    }

    // Helper method to compare Excel values including handling of different types
    private static bool AreEqual(object a, object b)
    {
        if (a == null && b == null)
            return true;
        if (a == null || b == null)
            return false;

        // Handle Excel-specific types
        if (a is ExcelDna.Integration.ExcelEmpty && b is ExcelDna.Integration.ExcelEmpty)
            return true;
        if (a is ExcelDna.Integration.ExcelEmpty || b is ExcelDna.Integration.ExcelEmpty)
            return false;
        if (a is ExcelDna.Integration.ExcelError || b is ExcelDna.Integration.ExcelError)
            return false;

        // Convert both to strings for comparison
        return a.ToString() == b.ToString();
    }

    // HashArray UDF
    // ------------------------------------------------------------------------------------
    //
    // This function computes a consistent hash value for an array of values.
    // It takes an array of objects and an optional hash length parameter.
    // The function converts all elements to strings, sorts them, and combines them into a single string.
    // It then generates a hash of the combined string using SHA256 and returns the hash value.
    // The hash length can be specified (4-32 characters), and defaults to 8 if not provided.
    // The function handles different types of input, including strings, numbers, booleans, and errors.
    [ExcelFunction(
        Description = "Computes a consistent hash value for an array (order-independent)",
        IsVolatile = false
    )]
    public static object HashArray(
        [ExcelArgument(Description = "Input array (can be horizontal or vertical)")] object[,] inputArray,
        [ExcelArgument(Description = "Optional: Length of hash (4-32 characters, default=8)")] object hashLengthObj
    )
    {
        try
        {
            // Default hash length
            int hashLength = 8;

            // Process optional hash length parameter
            if (hashLengthObj != null && !(hashLengthObj is ExcelMissing) && !(hashLengthObj is ExcelEmpty))
            {
                if (hashLengthObj is double)
                {
                    hashLength = (int)(double)hashLengthObj;
                }
                else if (hashLengthObj is int)
                {
                    hashLength = (int)hashLengthObj;
                }
                else if (hashLengthObj is string)
                {
                    if (!int.TryParse((string)hashLengthObj, out hashLength))
                    {
                        hashLength = 8;
                    }
                }

                // Validate hash length
                if (hashLength < 4) hashLength = 4;
                if (hashLength > 32) hashLength = 32;
            }

            // Convert all elements to strings and collect them in a list
            var elements = new List<string>();
            int rows = inputArray.GetLength(0);
            int cols = inputArray.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    object element = inputArray[i, j];
                    string elementStr = string.Empty;

                    if (element is double)
                    {
                        elementStr = ((double)element).ToString("G17");
                    }
                    else if (element is string)
                    {
                        elementStr = (string)element;
                    }
                    else if (element is bool)
                    {
                        elementStr = ((bool)element) ? "TRUE" : "FALSE";
                    }
                    else if (element is ExcelError)
                    {
                        elementStr = "ERROR:" + element.ToString();
                    }
                    else if (element == null || element is ExcelEmpty)
                    {
                        elementStr = string.Empty;
                    }
                    else
                    {
                        elementStr = element.ToString();
                    }

                    if (!string.IsNullOrEmpty(elementStr))
                    {
                        elements.Add(elementStr);
                    }
                }
            }

            // If no elements, return hash of empty string
            if (elements.Count == 0)
            {
                return GenerateHash(string.Empty, hashLength);
            }

            // Sort elements to make hash order-independent
            elements.Sort();

            // Combine all elements with a delimiter
            string combined = string.Join("|", elements);

            // Generate hash of the requested length
            return GenerateHash(combined, hashLength);
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    // Helper method to generate a hash of a string
    private static string GenerateHash(string input, int length)
    {
        // Use SHA256 to generate a hash of arbitrary length
        using (System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create())
        {
            byte[] inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            byte[] hashBytes = sha256.ComputeHash(inputBytes);

            // Convert to base64 and take the requested length
            string base64Hash = Convert.ToBase64String(hashBytes)
                .Replace("+", "0")  // Remove characters that might cause issues
                .Replace("/", "1")
                .Replace("=", "2");

            // Take the requested length (minimum 4, maximum 32)
            return base64Hash.Substring(0, Math.Min(Math.Max(length, 4), 32));
        }
    }

    // isLocalIP UDF
    // ------------------------------------------------------------------------------------
    //
    // This function checks if an IP address is local/private or routable.
    // It takes an IP address (optionally with port) as input and returns TRUE if local/private,
    // FALSE if routable, and #N/A if the input is invalid.
    [ExcelFunction(Description = "Returns TRUE if an IP address is local/private, FALSE if routable, #N/A if invalid.")]
    public static object isLocalIP([ExcelArgument(Description = "IP address (optionally with port)")] string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return ExcelError.ExcelErrorNA;

        try
        {
            // Remove port if present
            int colonIndex = input.LastIndexOf(':');
            string ipOnly = colonIndex > -1 && input.IndexOf(':') == colonIndex
                ? input.Substring(0, colonIndex)
                : input;

            // Support for IPv6 with port syntax like [::1]:1234
            if (ipOnly.StartsWith("[") && ipOnly.Contains("]"))
            {
                int end = ipOnly.IndexOf("]");
                ipOnly = ipOnly.Substring(1, end - 1);
            }

            System.Net.IPAddress ip;
            if (!System.Net.IPAddress.TryParse(ipOnly, out ip))
                return ExcelError.ExcelErrorNA;

            byte[] bytes = ip.GetAddressBytes();

            // IPv4 checks
            if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            {
                if (bytes[0] == 10)
                    return true; // 10.0.0.0/8
                if (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31)
                    return true; // 172.16.0.0/12
                if (bytes[0] == 192 && bytes[1] == 168)
                    return true; // 192.168.0.0/16
                if (bytes[0] == 127)
                    return true; // Loopback 127.0.0.0/8
                if (bytes[0] == 169 && bytes[1] == 254)
                    return true; // Link-local 169.254.0.0/16

                return false;
            }

            // IPv6 checks
            if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
            {
                if (System.Net.IPAddress.IsLoopback(ip))
                    return true;

                if (ip.IsIPv6LinkLocal || ip.IsIPv6SiteLocal)
                    return true;

                // Unique local address fc00::/7
                byte first = ip.GetAddressBytes()[0];
                if ((first & 0xFE) == 0xFC)
                    return true;

                return false;
            }

            return false;
        }
        catch
        {
            return ExcelError.ExcelErrorNA;
        }
    }
}