using System;

namespace ExcelDna.Integration
{
    [AttributeUsage(AttributeTargets.Method)]
    public sealed class ExcelFunctionAttribute : Attribute
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
        public bool IsMacroType { get; set; }
        public bool IsThreadSafe { get; set; }
    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public sealed class ExcelArgumentAttribute : Attribute
    {
        public string Description { get; set; }
        public bool AllowReference { get; set; }
    }

    public enum ExcelError
    {
        ExcelErrorNull,
        ExcelErrorDiv0,
        ExcelErrorValue,
        ExcelErrorRef,
        ExcelErrorName,
        ExcelErrorNum,
        ExcelErrorNA,
        ExcelErrorGettingData
    }

    public sealed class ExcelMissing { }
    public sealed class ExcelEmpty { }

    public sealed class ExcelReference
    {
        public int RowFirst { get; set; }
        public int ColumnFirst { get; set; }
        public int SheetId { get; set; }
        public void SetValue(object value) { }
    }

    public static class XlCall
    {
        public const int xlfVolatile = 1;
        public const int xlfCaller = 2;
        public const int xlfReftext = 3;
        public const int xlfAddress = 4;
        public static object Excel(int function, params object[] args) { return null; }
    }

    public static class ExcelDnaUtil
    {
        public static object Application { get; set; }
    }

    public static class ExcelAsyncUtil
    {
        public static void QueueAsMacro(Action action)
        {
            if (action != null) action();
        }
    }
}

namespace Microsoft.Office.Interop.Excel
{
    public delegate void WorkbookBeforeCloseEventHandler(Workbook workbook, ref bool cancel);

    public sealed class Workbook { }

    public sealed class Range
    {
        public Range EntireRow { get { return this; } }
        public Range EntireColumn { get { return this; } }
        public object Hidden { get; set; }
    }

    public sealed class MultiThreadedCalculation
    {
        public int ThreadCount { get; set; }
        public bool Enabled { get; set; }
    }

    public sealed class RangeIndexer
    {
        public Range this[string address] { get { return new Range(); } }
    }

    public sealed class Application
    {
        public event WorkbookBeforeCloseEventHandler WorkbookBeforeClose;
        public bool Iteration { get; set; }
        public int MaxIterations { get; set; }
        public double MaxChange { get; set; }
        public string Version { get; set; }
        public MultiThreadedCalculation MultiThreadedCalculation { get; set; }
        public RangeIndexer Range { get; private set; }

        public Application()
        {
            Version = "16.0";
            MultiThreadedCalculation = new MultiThreadedCalculation();
            Range = new RangeIndexer();
        }

        public Range get_Range(object cell1, object cell2) { return new Range(); }
        public void CalculateFull() { }
        public void CalculateFullRebuild() { }

        public void RaiseWorkbookBeforeClose(Workbook workbook, ref bool cancel)
        {
            if (WorkbookBeforeClose != null) WorkbookBeforeClose(workbook, ref cancel);
        }
    }
}
