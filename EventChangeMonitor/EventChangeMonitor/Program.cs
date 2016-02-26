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
using System.Threading;
using System.IO;
using System.Reflection;
using System.DirectoryServices.AccountManagement; 

namespace ActivityMonitor
{
    class Program 
    {
        private static Dictionary<string, Activity> activityList = new Dictionary<string, Activity>();
        private static Dictionary<string, Activity> packagedActivityList = new Dictionary<string, Activity>();
        private static Dictionary<string, List<String>> buckets = new Dictionary<string, List<String>>();
        private static Dictionary<string, TimeSpan> bucketedTime = new Dictionary<string, TimeSpan>();
        private static List<String> packageList = new List<string>();
        private static string lastWindowName = null;
        private static AutomationElement element = null;
        private static bool isGenerate = false;
        private static bool isReset = false;
        private static TimeSpan currentTime;
        private static Microsoft.Office.Interop.Excel.Application excel;
        private static Microsoft.Office.Interop.Excel.Workbook excelworkBook;
        private static Microsoft.Office.Interop.Excel.Worksheet excelSheetAll;
        private static Microsoft.Office.Interop.Excel.Worksheet excelSheetPackaged;
        private static Microsoft.Office.Interop.Excel.Worksheet excelSheetBucketed;
        private static Microsoft.Office.Interop.Excel.Range excelCellrange;
        private static Microsoft.Office.Interop.Excel.Range chartRange;
        private static Microsoft.Office.Interop.Excel.ChartObjects chartObjects;
        private static Microsoft.Office.Interop.Excel.ChartObject chartObject;
        private static Microsoft.Office.Interop.Excel.Chart chart;
        
        static void Main(string[] args)
        {

            DBConnector.getInstance().createNewBucket("Development");
            Console.WriteLine("Done !");
            Console.ReadKey();

            //RegisterInStartup();
            //TimeSpan startTime = new TimeSpan(9, 0, 0);
            //TimeSpan endTime = new TimeSpan(17,0, 0);
            //TimeSpan interval = new TimeSpan(0, 1, 0);

            ///**** set console configurations ****/
            //Console.WindowHeight = 50;
            //Console.WindowWidth = 150;
            //Console.BufferHeight = 9999;
            //Console.BufferWidth = 300;

            //Console.WriteLine("Process Name".PadRight(30) + "Duration".PadRight(20) + "Main Window Title");
            //definePackageList();
            //defineBuckets();

            ///**** start screen lock listner ****/
            //SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SysEventsCheck);

            //Thread t = new Thread(() => botSchedular(startTime, endTime, interval));
            //t.Start();

            /////**** wait for enter key ****/
            ////while (true)
            ////{
            ////    ConsoleKeyInfo c = Console.ReadKey();
            ////    if (c.Key == ConsoleKey.Enter)
            ////        generateExcel();
            ////}

            //if (Process.GetProcessesByName(Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location)).Count() > 1)
            //    Process.GetCurrentProcess().Kill();


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
                        //activityList.Remove(mainWindowTitle);
                    }
                    else
                    {
                        currentActivity = new Activity();
                        currentActivity.processName = process.ProcessName;
                        activityList.Add(mainWindowTitle, currentActivity);
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
                    //activityList.Add(mainWindowTitle, currentActivity);
                    lastWindowName = mainWindowTitle;
                }
                Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                foreach (KeyValuePair<string, Activity> pair in activityList)
                {
                    Console.WriteLine(pair.Value.processName.PadRight(30) + pair.Value.duration.ToString().PadRight(20) + pair.Key);
                }
            }
            catch (Exception e)
            {
                return;
            }
        }


        private static void definePackageList()
        {
            packageList.Add("eclipse");
            //packageList.Add("chrome");
            packageList.Add("WDExpress");
            packageList.Add("Brackets");
            packageList.Add("notepad++");
            packageList.Add("netbeans64");
            packageList.Add("explorer");
            //packageList.Add("firefox");
            packageList.Add("iexplore");
            packageList.Add("taskmgr");
            packageList.Add("OUTLOOK");
            packageList.Add("");
        }

        public static void generateExcel()
        {
            /******************** create a workbook *************************/
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excelworkBook = excel.Workbooks.Add(Type.Missing);

            /********************* create new sheet (Activity List) ***************************/
            excelSheetAll = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheetAll.Name = "Activity List";

            int row = 1;
            int tb1_start_x = row;
            int tb1_start_y = 1;
            excelSheetAll.Cells[row, 1] = "Process Name";
            excelSheetAll.Cells[row, 2] = "Duration";
            excelSheetAll.Cells[row, 3] = "Main Window Title";
            row++;
            foreach (KeyValuePair<string, Activity> pair in activityList)
            {
                excelSheetAll.Cells[row, 1] = pair.Value.processName;
                excelSheetAll.Cells[row, 2] = pair.Value.duration.ToString("g");
                excelSheetAll.Cells[row, 3] = pair.Key;
                row++;
            }

            int tb1_end_x = row-1;
            int tb1_end_y = 3;

            excelCellrange = excelSheetAll.Range[excelSheetAll.Cells[tb1_start_x, tb1_start_y], excelSheetAll.Cells[tb1_end_x, tb1_end_y]];
            excelCellrange.NumberFormat = "hh:mm:ss.000";
            excelCellrange.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            excelCellrange = excelSheetAll.Range[excelSheetAll.Cells[tb1_start_x, tb1_start_y], excelSheetAll.Cells[tb1_start_x, tb1_end_y]];
            FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);

            /*************************** create new sheet (Packaged Activity List) ****************************/
            excelSheetPackaged = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add();
            excelSheetPackaged.Name = "Packaged Activity List";
            row = 1;
            int tb2_start_x = row;
            int tb2_start_y = 1;
            excelSheetPackaged.Cells[row, 1] = "Process Name";
            excelSheetPackaged.Cells[row, 2] = "Duration";

            row++;
            foreach (KeyValuePair<string, Activity> pair in packagedActivityList)
            {
                excelSheetPackaged.Cells[row, 1] = pair.Key;
                excelSheetPackaged.Cells[row, 2] = pair.Value.duration.ToString("g");
                row++;
            }

            int tb2_end_x = row - 1;
            int tb2_end_y = 2;

            excelCellrange = excelSheetPackaged.Range[excelSheetPackaged.Cells[tb2_start_x, tb2_start_y], excelSheetPackaged.Cells[tb2_end_x, tb2_end_y]];

            excelCellrange.NumberFormat = "hh:mm:ss.000";
            excelCellrange.EntireColumn.AutoFit();
            border= excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            excelCellrange = excelSheetPackaged.Range[excelSheetPackaged.Cells[tb2_start_x, tb2_start_y], excelSheetPackaged.Cells[tb2_start_x, tb2_end_y]];
            FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);

            chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)excelSheetPackaged.ChartObjects(Type.Missing);
            chartObject = (Microsoft.Office.Interop.Excel.ChartObject)chartObjects.Add(220, 0, 400, 300);
            chart = chartObject.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Packaged Activity List";

            chartRange = excelSheetPackaged.get_Range("A" + tb2_start_x, "B" + tb2_end_x);
            chart.SetSourceData(chartRange, System.Reflection.Missing.Value);
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;

            chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)excelSheetPackaged.ChartObjects(Type.Missing);
            chartObject = (Microsoft.Office.Interop.Excel.ChartObject)chartObjects.Add(220, 320, 400, 300);
            chart = chartObject.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Packaged Activity List";

            chartRange = excelSheetPackaged.get_Range("A" + tb2_start_x, "B" + tb2_end_x);
            chart.SetSourceData(chartRange, System.Reflection.Missing.Value);
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            /************************* create new sheet (Bucketed Activity List) ******************************/
            excelSheetBucketed = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add();
            excelSheetBucketed.Name = "Bucketed Activity List";

            row = 1;
            int tb3_start_x = row;
            int tb3_start_y = 1;
            excelSheetBucketed.Cells[row, 1] = "Bucket Name";
            excelSheetBucketed.Cells[row, 2] = "Duration";

            row++;
            foreach (KeyValuePair<string, TimeSpan> pair in bucketedTime)
            {
                excelSheetBucketed.Cells[row, 1] = pair.Key;
                excelSheetBucketed.Cells[row, 2] = pair.Value.ToString("g");
                row++;
            }

            int tb3_end_x = row - 1;
            int tb3_end_y = 2;

            excelCellrange = excelSheetBucketed.Range[excelSheetBucketed.Cells[tb3_start_x, tb3_start_y], excelSheetBucketed.Cells[tb3_end_x, tb3_end_y]];

            excelCellrange.NumberFormat = "hh:mm:ss.000";
            excelCellrange.EntireColumn.AutoFit();
            border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            excelCellrange = excelSheetBucketed.Range[excelSheetBucketed.Cells[tb3_start_x, tb3_start_y], excelSheetBucketed.Cells[tb3_start_x, tb3_end_y]];
            FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);

            chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)excelSheetBucketed.ChartObjects(Type.Missing);
            chartObject = (Microsoft.Office.Interop.Excel.ChartObject)chartObjects.Add(220, 0, 400, 300);
            chart = chartObject.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Buckted Activity List";
            
            chartRange = excelSheetBucketed.get_Range("A" + tb3_start_x, "B" + tb3_end_x);
            chart.SetSourceData(chartRange, System.Reflection.Missing.Value);
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;

            chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)excelSheetBucketed.ChartObjects(Type.Missing);
            chartObject = (Microsoft.Office.Interop.Excel.ChartObject)chartObjects.Add(220, 320, 400, 300);
            chart = chartObject.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Buckted Activity List"; 

            chartRange = excelSheetBucketed.get_Range("A" + tb3_start_x, "B" + tb3_end_x);
            chart.SetSourceData(chartRange, System.Reflection.Missing.Value);
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            /*************** save excel *******************/

            //UserPrincipal.Current.DisplayName
            String filePath = "C:\\Users\\" + Environment.UserName + "\\Desktop\\ActivityList-" + Environment.UserName + "-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx";
            excelworkBook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //excelworkBook.SaveAs(filePath);
            excelworkBook.Close();
            excel.Quit();
            Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("Export to Excel");
            System.Diagnostics.Process.Start(filePath);
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
                    Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    Console.WriteLine("Lock Encountered");
                    OnFocusChangedHandler(element, null);
                    lastWindowName = null;
                    break;
                case SessionSwitchReason.ConsoleDisconnect:
                    Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    Console.WriteLine("Lock Encountered");
                    OnFocusChangedHandler(element, null);
                    lastWindowName = null;
                    break;
                //case SessionSwitchReason.SessionUnlock: Console.WriteLine("UnLock Encountered"); break;
            }
        }

        private static void botSchedular(TimeSpan startTime, TimeSpan endTime, TimeSpan interval)
        {
            while (true)
            {
                currentTime = DateTime.Now.TimeOfDay;
                if (currentTime > endTime && currentTime < endTime + interval)
                {
                    if (isGenerate == false)
                    {
                        Automation.RemoveAutomationFocusChangedEventHandler(OnFocusChangedHandler);
                        Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                        Console.WriteLine("stop listner");
                        Thread.Sleep(100);
                        generatePackageList();
                        calculateBucketTime();
                        generateExcel();
                        isGenerate = true;
                    }
                }
                else
                    isGenerate = false;

                if (currentTime > startTime && currentTime < endTime)
                {
                    if (isReset == false)
                    {
                        resetAll();
                        isReset = true;
                        Automation.AddAutomationFocusChangedEventHandler(OnFocusChangedHandler);
                        //Automation.AddStructureChangedEventHandler(eventdetect);
                        Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                        Console.WriteLine("start listner");
                    }
                }
                else
                {
                    isReset = false;
                }

                Thread.Sleep(1000);
            }
        }

        private static void RegisterInStartup()
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (registryKey.GetValue("CentroidBot") == null)
            {
                registryKey.SetValue("CentroidBot", Application.ExecutablePath);
            }
            else
            {
                if (registryKey.GetValue("CentroidBot").ToString() != Application.ExecutablePath)
                {
                    registryKey.SetValue("CentroidBot", Application.ExecutablePath);
                }
            }
        }

        private static void resetAll()
        {
            activityList.Clear();
            lastWindowName = null;
        }

        private static void generatePackageList()
        {
            foreach (KeyValuePair<string, Activity> pair in activityList)
            {
                if (packagedActivityList.ContainsKey(pair.Value.processName))
                {
                    packagedActivityList[pair.Value.processName].duration = packagedActivityList[pair.Value.processName].duration + activityList[pair.Key].duration;
                }
                else
                    packagedActivityList.Add(pair.Value.processName, (Activity)pair.Value.Clone());
            }

            //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
            //foreach (KeyValuePair<string, Activity> pair in packagedActivityList)
            //{
            //    Console.WriteLine(pair.Key.PadRight(30) + pair.Value.duration.ToString().PadRight(20));
            //}

        }

        private static void defineBuckets()
        {
            List<string> researchList = new List<string>();
            researchList.Add("chrome");
            researchList.Add("firefox");
            researchList.Add("iexplore");

            List<string> developmentList = new List<string>();
            developmentList.Add("notepad++");
            developmentList.Add("Brackets");
            developmentList.Add("netbeans64");
            developmentList.Add("eclipse");
            developmentList.Add("cmd");
            developmentList.Add("sh");
            developmentList.Add("MySQLWorkbench");
            developmentList.Add("java");
            developmentList.Add("mintty");
            developmentList.Add("idea");
            developmentList.Add("TortoiseGitProc");
            developmentList.Add("dwm");
            developmentList.Add("javaw");
            developmentList.Add("SourceTree");
            developmentList.Add("WDExpress");


            List<string> communicationList = new List<string>();
            communicationList.Add("OUTLOOK");
            communicationList.Add("lync");

            List<string> documentationList = new List<string>();
            documentationList.Add("EXCEL");
            documentationList.Add("WINWORD");
            documentationList.Add("POWERPNT");
            documentationList.Add("notepad");
            documentationList.Add("AcroRd32");
            documentationList.Add("StikyNot");

            List<string> otherList = new List<string>();
            otherList.Add("taskmgr");
            otherList.Add("regedit");
            otherList.Add("dllhost");
            otherList.Add("SnippingTool");
            otherList.Add("WinRAR");
            otherList.Add("7zG");
            otherList.Add("explorer");

            buckets.Add("Research", researchList);
            buckets.Add("Development", developmentList);
            buckets.Add("Communication", communicationList);
            buckets.Add("Documentation", documentationList);
            buckets.Add("Other", otherList);
        }

        private static void calculateBucketTime()
        {
            bucketedTime.Add("Research",new TimeSpan(0,0,0));
            bucketedTime.Add("Development", new TimeSpan(0, 0, 0));
            bucketedTime.Add("Communication", new TimeSpan(0, 0, 0));
            bucketedTime.Add("Documentation", new TimeSpan(0, 0, 0));
            bucketedTime.Add("Other", new TimeSpan(0, 0, 0));

            foreach (KeyValuePair<string, Activity> packagedActivity in packagedActivityList)
            {
                foreach (KeyValuePair<string, List<string>> bucket in buckets)
                {
                    if (bucket.Value.Contains(packagedActivity.Key))
                    {
                        bucketedTime[bucket.Key] = bucketedTime[bucket.Key] + packagedActivity.Value.duration;
                        break;
                    }
                }
            }

            //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
            //foreach (KeyValuePair<string, TimeSpan> pair in bucketedTime)
            //{
            //    Console.WriteLine(pair.Key.PadRight(30) + pair.Value.ToString());

            //}

        }


    }
}
