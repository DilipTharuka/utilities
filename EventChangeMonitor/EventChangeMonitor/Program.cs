using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Diagnostics;
using System.Data;
using Microsoft.Win32;
using System.Windows.Forms;

namespace EventChangeMonitor
{
    class Program 
    {
        private static Dictionary<string, Activity> activityList = new Dictionary<string, Activity>();
        private static List<String> packageList = new List<string>();
        private static string lastWindowName = null;
        private static AutomationElement element = null;
        private static Microsoft.Office.Interop.Excel.Application excel;
        private static Microsoft.Office.Interop.Excel.Workbook excelworkBook;
        private static Microsoft.Office.Interop.Excel.Worksheet excelSheet;
        private static Microsoft.Office.Interop.Excel.Range excelCellrange;
       

        static void Main(string[] args)
        {
            TimeSpan startTime = new TimeSpan(19,28,0);
            TimeSpan interval = new TimeSpan(0,5,0);
            TimeSpan endTime = startTime.Add(interval);
            bool isGenerate = false;
            TimeSpan currentTime;

            /**** set console configurations ****/
            //Console.WindowHeight = 50;
            //Console.WindowWidth = 150;
            //Console.BufferHeight = 999;
            //Console.BufferWidth = 200;

            //Console.WriteLine("Process Name".PadRight(30) + "Duration".PadRight(20) + "Main Window Title");
            definePackageList();

            /**** start focus listner ****/
            Automation.AddAutomationFocusChangedEventHandler(OnFocusChangedHandler);

            /**** start screen lock listner ****/
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SysEventsCheck);

            /**** wait for enter key ****/
            //while (true)
            //{
            //    ConsoleKeyInfo c = Console.ReadKey();
            //    if (c.Key == ConsoleKey.Enter)
            //        generateExcel();
            //}

            while (true)
            {
                currentTime = DateTime.Now.TimeOfDay;
                if (currentTime > startTime && currentTime < endTime)
                {
                    if (isGenerate == false)
                    {
                        generateExcel();
                        isGenerate = true;
                    }
                }
                else
                    isGenerate = false;
            }

        }

        private static void OnFocusChangedHandler(object src, AutomationFocusChangedEventArgs args)
        {
            try
            {
                DateTime startTime = DateTime.Now;
                element = src as AutomationElement;
                if (element != null)
                {
                    /*** get the main window title of focus element ***/
                    int processId = element.Current.ProcessId;
                    Process process = Process.GetProcessById(processId);
                    
                    string mainWindowTitle = null;
                    if (packageList.Contains(process.ProcessName))
                        mainWindowTitle = process.ProcessName;
                    else
                        mainWindowTitle = process.MainWindowTitle;

                    /*** get the activity object corresponding to main window title ***/
                    Activity currentActivity = null;
                    if (activityList.ContainsKey(mainWindowTitle))
                    {
                        currentActivity = activityList[mainWindowTitle];
                        activityList.Remove(mainWindowTitle);
                    }
                    else
                    {
                        currentActivity = new Activity();
                        currentActivity.processName = process.ProcessName;
                    }        

                    if (lastWindowName != null)
                    {

                        if (lastWindowName == mainWindowTitle)
                        {
                            currentActivity.endTime = startTime;
                            currentActivity.generateDuration();
                        }

                        else
                        {
                            activityList[lastWindowName].endTime = startTime;
                            activityList[lastWindowName].generateDuration();
                        }
                        
                    }
                    currentActivity.startTime = startTime;
                    activityList.Add(mainWindowTitle, currentActivity);
                    lastWindowName = mainWindowTitle;
                }
                //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                //foreach (KeyValuePair<string, Activity> pair in activityList)
                //{
                //    Console.WriteLine(pair.Value.processName.PadRight(30) + pair.Value.duration.ToString().PadRight(20) + pair.Key);
                //}
            }
            catch (Exception e)
            {
                return;
            }
        }


        private static void definePackageList()
        {
            packageList.Add("eclipse");
            packageList.Add("chrome");
            packageList.Add("WDExpress");
            packageList.Add("Brackets");
            packageList.Add("notepad++");
            packageList.Add("netbeans64");
            packageList.Add("explorer");
            packageList.Add("firefox");
            packageList.Add("iexplore");
            packageList.Add("taskmgr");
            packageList.Add("OUTLOOK");
            packageList.Add("");

        }

        public static void generateExcel()
        {
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excelworkBook = excel.Workbooks.Add(Type.Missing);
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = "Test work sheet";
            excelSheet.Cells[1,1] = "Process Name";
            excelSheet.Cells[1,2] = "Duration";
            excelSheet.Cells[1,3] = "Main Window Title";
            int i = 2;
            foreach (KeyValuePair<string, Activity> pair in activityList)
	        {
		        excelSheet.Cells[i,1] = pair.Value.processName;  
                excelSheet.Cells[i,2] = pair.Value.duration.ToString("g");
                excelSheet.Cells[i,3] = pair.Key;
                i++;
	        }
            excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[activityList.Count+1,3]];
            excelCellrange.NumberFormat = "hh:mm:ss.000";
            excelCellrange.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, 3]];
            FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);
            String filePath = "C:\\Users\\" + Environment.UserName + "\\Desktop\\ActivityList-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx";
            excelworkBook.SaveAs(filePath); 
            excelworkBook.Close();
            excel.Quit();
            //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
            //Console.WriteLine("Export to Excel");
            //System.Diagnostics.Process.Start(filePath);
        }


        public static void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }


        private static void SysEventsCheck(object sender, SessionSwitchEventArgs e)
        {
            switch (e.Reason)
            {
                case SessionSwitchReason.SessionLock:
                    //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    //Console.WriteLine("Lock Encountered");
                    OnFocusChangedHandler(element, null);
                    lastWindowName = null;
                    break;
                case SessionSwitchReason.ConsoleDisconnect:
                    //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    //Console.WriteLine("Lock Encountered");
                    OnFocusChangedHandler(element, null);
                    lastWindowName = null;
                    break;
                //case SessionSwitchReason.SessionUnlock: Console.WriteLine("UnLock Encountered"); break;
            }
        }
    }
}
