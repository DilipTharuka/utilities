﻿using System;
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
using System.Configuration.Install;

namespace ActivityMonitor
{
    class Program
    {
        private static Dictionary<string, Activity> activityList = new Dictionary<string, Activity>();
        private static List<String> packages = new List<string>();
        private static Dictionary<string, TimeSpan> packagedList = new Dictionary<string, TimeSpan>();
        public static Dictionary<string, List<string>> buckets = new Dictionary<string, List<string>>();
        private static Dictionary<string, TimeSpan> bucketedList = new Dictionary<string, TimeSpan>();
        private static string dir;
        private static Thread io;

        private static string lastWindowName = null;
        private static AutomationElement element = null;
        private static bool isGenerate = false;
        private static bool isReset = false;
        private static TimeSpan currentTime;

        private static TimeSpan startTime ;
        private static TimeSpan endTime ;
        private static TimeSpan interval ;

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


        private static void initActivityMonitor()
        {
            /**************** define backup dir *********************/
            dir = "C:\\Users\\" + Environment.UserName + "\\AppData\\Local\\AM\\";
            Directory.CreateDirectory(dir);

            startTime = new TimeSpan(9, 0, 0);
            endTime = new TimeSpan(17, 0, 0);
            interval = new TimeSpan(0, 1, 0);

            /**** set console configurations ****/
            //Console.WindowHeight = 50;
            //Console.WindowWidth = 150;
            //Console.BufferHeight = 9999;
            //Console.BufferWidth = 300;

            Console.WriteLine("Process Name".PadRight(30) + "Duration".PadRight(20) + "Main Window Title");

        }

        static void Main(string[] args)
        {

            initActivityMonitor();

            /**************** Kill If Process Exists ****************/
            if (Process.GetProcessesByName(Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location)).Count() > 1)
                Process.GetCurrentProcess().Kill();

            /***************** Set UI Property **********************/
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            /****************** Add to StartUp **********************/
            addToStartup();

            /*************** start screen lock listner **************/
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEventListner);

            Thread schedular = new Thread(() => botSchedular(startTime, endTime, interval));
            schedular.Start();

            ///**** wait for enter key ****/
            //while (true)
            //{
            //    ConsoleKeyInfo c = Console.ReadKey();
            //    if (c.Key == ConsoleKey.Enter)
            //        generateExcel();
            //}
     
        }

        private static void userEventListner(object src, AutomationFocusChangedEventArgs args)
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
                    if (packages.Contains(process.ProcessName))
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
                return ;
            }
        }

        public static void generateExcel()
        {

            packages = DBConnector.getInstance().getPackages();
            buckets = DBConnector.getInstance().getBuckets();
            generatePackageList();
            checkProcess();
            generateBucketList();

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

            int tb1_end_x = row - 1;
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
            excelSheetPackaged.Cells[row, 1] = "Package Name";
            excelSheetPackaged.Cells[row, 2] = "Duration";

            row++;
            foreach (KeyValuePair<string, TimeSpan> pair in packagedList)
            {
                excelSheetPackaged.Cells[row, 1] = pair.Key;
                excelSheetPackaged.Cells[row, 2] = pair.Value.ToString("g");
                row++;
            }

            int tb2_end_x = row - 1;
            int tb2_end_y = 2;

            excelCellrange = excelSheetPackaged.Range[excelSheetPackaged.Cells[tb2_start_x, tb2_start_y], excelSheetPackaged.Cells[tb2_end_x, tb2_end_y]];

            excelCellrange.NumberFormat = "hh:mm:ss.000";
            excelCellrange.EntireColumn.AutoFit();
            border = excelCellrange.Borders;
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
            foreach (KeyValuePair<string, TimeSpan> pair in bucketedList)
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
            Thread.Sleep(1000);
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

        private static void SystemEventListner(object sender, SessionSwitchEventArgs e)
        {
            switch (e.Reason)
            {
                case SessionSwitchReason.SessionLock:
                    Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    Console.WriteLine("Lock Encountered");
                    userEventListner(element, null);
                    lastWindowName = null;
                    break;
                case SessionSwitchReason.ConsoleDisconnect:
                    Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    Console.WriteLine("Lock Encountered");
                    userEventListner(element, null);
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
                if (currentTime > startTime && currentTime < endTime)
                {
                    if (isReset == false)
                    {
                        resetMonitor();
                        Deserialize();
                        //if (!checkExcelIsgenerated())
                        //    generateExcel();
                        Automation.AddAutomationFocusChangedEventHandler(userEventListner);
                        Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                        Console.WriteLine("start listner");
                        io = new Thread(() => Serialize());
                        io.Start();
                        isReset = true;
                    }
                }
                else
                {
                    isReset = false;
                }

                if (currentTime > endTime && currentTime < endTime + interval)
                {
                    if (isGenerate == false)
                    {
                        Automation.RemoveAutomationFocusChangedEventHandler(userEventListner);
                        Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                        Console.WriteLine("stop listner");
                        Thread.Sleep(100);
                        generateExcel();
                        isGenerate = true;
                        io.Abort();
                    }
                }
                else
                    isGenerate = false;

                Thread.Sleep(1000);
            }
        }

        private static void addToStartup()
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

        private static void resetMonitor()
        {
            activityList.Clear();
            packagedList.Clear();
            bucketedList.Clear();
            lastWindowName = null;
        }

        private static void generatePackageList()
        {
            foreach (KeyValuePair<string, Activity> pair in activityList)
            {
                if (packagedList.ContainsKey(pair.Value.processName))
                {
                    packagedList[pair.Value.processName] = packagedList[pair.Value.processName] + activityList[pair.Key].duration;
                }
                else
                    packagedList.Add(pair.Value.processName, pair.Value.duration);
            }

            //Console.WriteLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
            //foreach (KeyValuePair<string, Activity> pair in packagedActivityList)
            //{
            //    Console.WriteLine(pair.Key.PadRight(30) + pair.Value.duration.ToString().PadRight(20));
            //}

        }

        private static void generateBucketList()
        {
            foreach (KeyValuePair<string, List<string>> pair in buckets)
            {
                bucketedList.Add(pair.Key, new TimeSpan(0, 0, 0));             
            }
            foreach (KeyValuePair<string, TimeSpan> packagedActivity in packagedList)
            {
                foreach (KeyValuePair<string, List<string>> bucket in buckets)
                {
                    if (bucket.Value.Contains(packagedActivity.Key))
                    {
                        bucketedList[bucket.Key] = bucketedList[bucket.Key] + packagedActivity.Value;
                        break;
                    }
                }
            }
        }


        public static void checkProcess()
        {
            bool isFound;
            foreach (KeyValuePair<string, TimeSpan> packagedActivity in packagedList)
            {
                isFound = false;
                foreach (KeyValuePair<string, List<string>> bucket in buckets)
                {
                    if (bucket.Value.Contains(packagedActivity.Key))
                    {
                        isFound = true;
                        break;
                    }
                }
                if (isFound == false)
                {
                    Console.WriteLine(packagedActivity.Key);
                    FormUser formUser = new FormUser(packagedActivity.Key);
                    formUser.TopMost = true;
                    Application.Run(formUser);                   
                }
            }
        }


        public static void Serialize()
        {
            while (true)
            {
                var f_fileStream = new FileStream(dir + DateTime.Now.ToString("yyyy-MM-dd") + ".amos", FileMode.Create, FileAccess.Write);
                var f_binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                f_binaryFormatter.Serialize(f_fileStream, activityList);
                f_fileStream.Close();
                Thread.Sleep(10000);
            }
        }

        public static void Deserialize()
        {
            if (File.Exists(dir + DateTime.Now.ToString("yyyy-MM-dd") + ".amos"))
            {
                var f_fileStream = File.OpenRead(dir + DateTime.Now.ToString("yyyy-MM-dd") + ".amos");
                var f_binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                activityList = (Dictionary<string, Activity>)f_binaryFormatter.Deserialize(f_fileStream);
                lastWindowName = null;
                f_fileStream.Close();
            }
        }


        //public static bool checkExcelIsgenerated()
        //{
        //    bool isGenerated = false;
        //    DateTime today = DateTime.Now;
        //    DateTime yesterday = today.AddDays(-1);
        //    if (File.Exists(dir + yesterday.ToString("yyyy-MM-dd")))
        //        isGenerated = true;

        //    return isGenerated;
        //}
    }
}
