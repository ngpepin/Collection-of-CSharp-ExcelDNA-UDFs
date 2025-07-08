using System;
using System.Timers;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

public class C 
{
    static Timer _timer;
    
    [ExcelDna.Integration.ExcelCommand()] 
    public static void startAutoUpdate()
    {
        if (_timer != null)
        {
            _timer.Stop();
            _timer.Dispose();
            _timer = null;
            statusMsg("Auto-refresh stopped.");
            return;
        }

        // Immediate refresh
        PerformFullRefresh();

        // Scheduled every 10 minutes
        _timer = new Timer(10 * 60 * 1000); // 10 minutes
        _timer.Elapsed += (s, e) => PerformFullRefresh();
        _timer.AutoReset = true;
        _timer.Start();

        statusMsg("Auto-refresh started: workbook will refresh every 10 minutes.");
    }
    
    [ExcelDna.Integration.ExcelCommand()]     
    public static void stopAutoUpdate()
    {
        if (_timer != null)
        {
            _timer.Stop();
            _timer.Dispose();
            _timer = null;
            statusMsg("Auto-refresh stopped.");
            return;
        }
    }

    static void PerformFullRefresh()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;

                // Refresh all data connections and PivotTables
                wb.RefreshAll();

                // Optionally update timestamp in A1 of the active sheet
                var ws = (Excel.Worksheet)app.ActiveSheet;
                ws.Range["A1"].Value2 = "Last full refresh: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Refresh error: " + ex.Message);
            }
        });
    }

    static void statusMsg(string message)
    {
        var app = (Excel.Application)ExcelDnaUtil.Application;
        app.StatusBar = message;
    }
}