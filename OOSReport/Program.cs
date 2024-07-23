using OOSReport.SQL_Classes;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Pdf;
using OOSReport.Analytics;
using System.IO;
using OOSReport.Classes;
using System.Threading;
using DevExpress.ReportServer.ServiceModel.DataContracts;
using DevExpress.DataAccess.DataFederation;
using DevExpress.XtraReports;
using System.Configuration;
using DevExpress.Printing.Utils.DocumentStoring;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DevExpress.XtraRichEdit.Import.Html;
using DevExpress.DataAccess.Sql;
using DevExpress.CodeParser;
using DevExpress.PivotGrid.OLAP.Mdx;
using DevExpress.UnitConversion;
using static DevExpress.XtraPrinting.Native.ExportOptionsPropertiesNames;
using CER.Logger;

namespace OOSReport
{
    public class Program
    {
       

        #region Declarations
        private static string emailFrom = "ReportingServices@Spar.co.za";
        private static string mailingList = ConfigurationManager.AppSettings["MailingList"]; 
        public static DataTable StoreIds { get; set; }
        public static DataTable SupplierIds { get; set; }
        public static DataTable StoreIdXRows { get; set; }
        public static DataTable RemainingStoreIds { get; set; }
        public static DataTable RemainingStoreIdXRows { get; set; }
        public static DataTable RemainingStores { get; set; }
        public static DataTable subscribedTable { get; set; }
        public static DataTable CheckDB { get; set; }
        public static int NumStores = 0;
        public static int NumSuppliers = 0;
        public static DateTime GenDate = DateTime.Now;
        public static string ReportDate = GenDate.ToString("dd/MM/yyyy");     
     
        #endregion
        static void Main(string[] args)
        {
            CatmanErrorLogger.AppName = "OOS Report";
            CatmanErrorLogger.Recipients = mailingList.Contains(',') ? mailingList.Split(',').Select(p => p.Trim()).ToArray() : new string[] { mailingList.Trim() };
            var oosReportTask = Task.Run(StartOOSReport);
            var supplierReportTask = Task.Run(StartSupplierReport);
            //supplierReportTask
            Task.WaitAll(oosReportTask);
        }

        public static void StartOOSReport()
        {
            

            //This method gets the stores that need reports generated for, sends the store IDs to the CreateReport method for more processing.
            //It differenciates between the 2 tables it gets data from by putting the different datas in 2 different data tables. 
            //It then gives feedback to the console indicating the number of stores queried and whether a report was successfully generated or not.   
            string logFilePath = ConfigurationManager.AppSettings["LogFilePath"]; //gets the file location to send the log files to.
            File.WriteAllText(Path.Combine(logFilePath, "Logs.txt"), "");

            DataTable CheckDB = SqlFunctionsCoscat4.GetData($@"SELECT DISTINCT StoreId FROM [dbo].[//]");
            DateTime CurrentTime = DateTime.Now;
            var watch = new Stopwatch();
            if (CheckDB.Rows.Count == 0) //If this returns true then it will not use the // to generate reports
            {
                try
                {
                    watch.Start();

                    DataTable StoreID_dt = SqlFunctionsCoscat4.GetData
                        ($@"SELECT DISTINCT psss.StoreID                                                                        
                            FROM //");

                    DataTable StoreID_dtXRows = SqlFunctionsCoscat4.GetData
                        ($@"SELECT DISTINCT psss.StoreId, psss.NumRows
                            FROM //");

                    Console.WriteLine("Starting Store reports...");
                    StoreIds = StoreID_dt;
                    StoreIdXRows = StoreID_dtXRows;
                    NumStores = StoreID_dt.Rows.Count + StoreID_dtXRows.Rows.Count;
                    Console.WriteLine("Number of stores to query: " + NumStores.ToString());
                    Analytic.NumberOfStores = NumStores;
                    Console.WriteLine("Querying Database...");
                    RemainingStoreIds = CreateReport(StoreID_dt); //Whichever storeIDs are not used after the first run of the program will
                                                                  //go in these tables to be used again on another run.
                    RemainingStoreIdXRows = CreateReportXRows(StoreID_dtXRows);
                    DataRow[] storeIds = RemainingStoreIds.Select();
                    foreach (var storeid in storeIds)
                    {
                        int storeId = 0;
                        storeId = (int)storeid.ItemArray[0];
                        SqlFunctionsCoscat4.ExecuteNonReturnSQL($@"INSERT INTO // ([StoreId]) VALUES ({storeId})");
                    }
                    Console.WriteLine($"Elapsed Time: {watch.Elapsed}");
                    string FileEntry = Analytic.AnalyticsStatsReport();
                    if (FileEntry.Equals("ALERT: No reports generated. Please check for errors."))
                    {
                        Logs.CreateFile(FileEntry);
                        string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                        logPath = Path.Combine(logPath, "Logs.txt");
                        string EmailDetails = Analytic.AnalyticsStatsEmail();
                        eMail.SendeMail(mailingList, emailFrom, "---*OOS ALERT*---", "---*Alert*---" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                    }
                    else
                    {
                        watch.Stop();
                        Logs.CreateFile(FileEntry);
                        string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                        logPath = Path.Combine(logPath, "Logs.txt");
                        string EmailDetails = Analytic.AnalyticsStatsEmail();
                        eMail.SendeMail(mailingList, emailFrom, "OOS Report Logs", "Please see log file\n" + EmailDetails + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                    }
                }
                catch (IndexOutOfRangeException e)
                {
                    Logs.GeneralError = e.ToString();
                    string FileEntry = Logs.LogGeneralErrorReport();
                    Logs.CreateFile(FileEntry);
                    Logs.LogsCount++;
                    string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                    logPath = Path.Combine(logFilePath, "Logs.txt");
                    eMail.SendeMail(mailingList, emailFrom, "OOS Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                }
            }
            else
            {
                watch.Start();
                try
                { //The storeIDs that were not used in the previous run will now be used here and this will repeat each time the program is run until the table becomes empty.
                  //The OOSRemaingStores table is automatically truncated before each day to make sure the correct data is used.
                    RemainingStores = SqlFunctionsCoscat4.GetData($@"//");
                    CurrentTime = DateTime.Now;

                    if (RemainingStores.Rows.Count > 0 /*&& CurrentTime.TimeOfDay.Hours < 17*/)
                    {
                        Console.WriteLine("Starting Program...");
                        StoreIds = RemainingStores;
                        NumStores = RemainingStores.Rows.Count;
                        Console.WriteLine("Number of stores to query: " + NumStores.ToString());
                        Analytic.NumberOfStores = NumStores;
                        Console.WriteLine("Querying Database...");
                        RemainingStoreIds = CreateRemainingReports(RemainingStores);
                    }
                    Console.WriteLine($"Elapsed Time: {watch.Elapsed}");
                    string FileEntry = Analytic.AnalyticsStatsReport();
                    watch.Stop();
                    Logs.CreateFile(FileEntry);
                    string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                    logPath = Path.Combine(logPath, "Logs.txt");
                    string EmailDetails = Analytic.AnalyticsStatsEmail();
                    eMail.SendeMail(mailingList, emailFrom, "OOS Report Logs", "Please see log file\n" + EmailDetails + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                }
                catch (IndexOutOfRangeException e)
                {
                    Logs.GeneralError = e.ToString();
                    string FileEntry = Logs.LogGeneralErrorReport();
                    Logs.CreateFile(FileEntry);
                    Logs.LogsCount++;
                    string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                    logPath = Path.Combine(logPath, "Logs.txt");
                    eMail.SendeMail(mailingList, emailFrom, "OOS Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                }
            }
        }

        #region Start Supplier Report

        public static void StartSupplierReport()
        {
            string logFilePath = ConfigurationManager.AppSettings["SupplierLogFilePath"]; //gets the file location to send the log files to.
            File.WriteAllText(Path.Combine(logFilePath, "Logs.txt"), "");
            DateTime CurrentTime = DateTime.Now;
            var watch = new Stopwatch();

            try
            {
                watch.Start();

                DataTable SupplierId_dt = SqlFunctionsCoscat4.GetData
                    ($@"SELECT DISTINCT SupplierID
                        FROM //;");
                         
                Console.WriteLine("Starting Supplier Reports...");
                SupplierIds = SupplierId_dt;

                NumSuppliers = SupplierId_dt.Rows.Count;
                Console.WriteLine("Number of suppliers to query: " + NumSuppliers.ToString());
                SupplierAnalytic.NumberOfSuppliers = NumSuppliers;
                Console.WriteLine("Querying Database...");
                RemainingStoreIds = CreateReport(SupplierId_dt); 
                DataRow[] supplierIds = RemainingStoreIds.Select();
                foreach (var supplierid in supplierIds)
                {
                    int supplierId = 0;
                    supplierId = (int)supplierid.ItemArray[0];
                    SqlFunctionsCoscat4.ExecuteNonReturnSQL($@"INSERT INTO // ([StoreId]) VALUES ({supplierId})");
                }
                Console.WriteLine($"Elapsed Time: {watch.Elapsed}");
                string FileEntry = SupplierAnalytic.AnalyticsStatsReport();
                if (FileEntry.Equals("ALERT: No reports generated. Please check for errors."))
                {
                    SupplierLogs.CreateFile(FileEntry);
                    string logPath = ConfigurationManager.AppSettings["SupplierLogFilePath"];
                    logPath = Path.Combine(logPath, "Logs.txt");
                    string EmailDetails = SupplierAnalytic.AnalyticsStatsEmail();
                    eMail.SendeMail(mailingList, emailFrom, "---*OOS ALERT*---", "---*Alert*---" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                    return;
                }
                else
                {
                    watch.Stop();
                    SupplierLogs.CreateFile(FileEntry);
                    string logPath = ConfigurationManager.AppSettings["SupplierLogFilePath"];
                    logPath = Path.Combine(logPath, "Logs.txt");
                    string EmailDetails = SupplierAnalytic.AnalyticsStatsEmail();
                    eMail.SendeMail(mailingList, emailFrom, "OOS Supplier Report Logs", "Please see log file\n" + EmailDetails + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                    return;
                }
            }
            catch (IndexOutOfRangeException e)
            {
                SupplierLogs.GeneralError = e.ToString();
                string FileEntry = SupplierLogs.LogGeneralErrorReport();
                SupplierLogs.CreateFile(FileEntry);
                SupplierLogs.LogsCount++;
                string logPath = ConfigurationManager.AppSettings["SupplierLogFilePath"];
                logPath = Path.Combine(logFilePath, "Logs.txt");
                eMail.SendeMail(mailingList, emailFrom, "OOS Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                return;
            }
        }

        #endregion

        #region Create Report       
        public static DataTable CreateReport(DataTable StoreIDV1)
        {
            //Gets the StoreID
            DataTable LoadedTable = StoreIDV1.Copy();
            int Store_ID = 0;
            int checkID = 0;
            int NumReports = 0;
            int Type = 1;
            Analytic.NumberOfNoData = 0;
            try
            {
                foreach (DataRow Store_Id in StoreIDV1.Rows)
                {
                    try
                    {
                        checkID = (int)Store_Id.ItemArray[0];
                        bool Loaded = false;
                        bool subscribed = false;
                        //Checks if the store is subscribed to the OOS service or not
                        DataTable subscribedTable = SqlFunctionsCoscat4.GetData
                            ($@"DECLARE @storeId AS INT;
                                SET @storeId = {checkID};
                                SELECT mps.Unsubscribed
                                //
                             ");
                      
                        if (subscribedTable.Rows.Count == 0)
                        {
                            continue;
                        }
                        DataRow subscribedRow = subscribedTable.Rows[0];
                        string subscribedResult = subscribedRow[0].ToString();
                        if (subscribedResult.Equals("False"))
                        {
                            subscribed = true;
                        }

                        if (subscribed == true)
                        {
                            //Compiles the relavent data needed to create the report and sends the data to the OOSReport class.
                            Store_ID = (int)Store_Id.ItemArray[0];
                            OOSReport report1 = new OOSReport(Store_ID, ref Loaded, Type, 40);
                            
                            if (Loaded)
                            {
                                //Creates the report file and sends it to the correct folder to be stored.
                                DateTime GenDate = DateTime.Now;
                                string ReportDate = GenDate.ToString("yyyyMMdd");
                                string filename = "Out Of Stock Report_" + Store_ID.ToString() + "_" + ReportDate + ".pdf";
                               
                                string filePath = ConfigurationManager.AppSettings["OutPutPath"];
                                filePath = Path.Combine(filePath, filename);
                                if (!File.Exists(filePath))
                                {
                                    string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;

                                    string defaultFilePath = Path.GetDirectoryName(exePath);
                                    defaultFilePath = Path.Combine(defaultFilePath, filename);
                                    report1.CreateDocument();
                                    report1.ExportToPdf(defaultFilePath);
                                    if (File.Exists(filePath))
                                    {
                                        File.Delete(filePath);
                                    }
                                   
                                    File.Move(defaultFilePath, filePath);
                                    
                                    NumReports++;
                                    Analytic.GetNumReports(NumReports);
                                    DataRow[] selection = LoadedTable.Select($"StoreId={Store_Id[0]}");
                                    LoadedTable.Rows.Remove(selection[0]);

                                }
                            }
                        }
                    }
                    catch(Exception e)
                    {

                        continue;
                    }
                }
                return LoadedTable;
            }
            catch (FileNotFoundException ex)
            {
                Logs.GeneralError = ex.ToString();
                string FileEntry = Logs.LogGeneralErrorReport();
                Logs.LogsCount++;
                string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                logPath = Path.Combine(logPath, "Logs.txt");
                eMail.SendeMail(mailingList, emailFrom, "OOS Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);              
                return LoadedTable;                
            }
        }
        #endregion

        #region CreateReportXRows
        public static DataTable CreateReportXRows(DataTable StoreIDV1)
        {
            //Same method as CreateReport except using // table to check whether the Store is subscribed or not.
            //Also sends different data that determines the number of rows on the report.
            DataTable LoadedTable = StoreIDV1.Copy();
            int Store_ID = 0;
            int checkID = 0;
            int NumReports = 0;
            int NumRows = 0;
            Analytic.NumberOfNoData = 0;
            try
            {
                foreach (DataRow Store_Id in StoreIDV1.Rows)
                {
                    try
                    {
                        checkID = (int)Store_Id.ItemArray[0];
                        bool Loaded = false;
                        bool subscribed = false;
                        int Type = 2;                        
                        DataTable subscribedTable = SqlFunctionsCoscat4.GetData
                            ($@"DECLARE @storeId AS INT;
                                SET @storeId = {checkID};
                                SELECT mps.Unsubscribed
                                //                               
                            ");
                        if (subscribedTable.Rows.Count == 0)
                        {
                            continue;
                        }
                        DataRow subscribedRow = subscribedTable.Rows[0];
                        string subscribedResult = subscribedRow[0].ToString();
                        if (subscribedResult.Equals("False"))
                        {
                            subscribed = true;
                        }

                        if (subscribed == true)
                        {
                            Store_ID = (int)Store_Id.ItemArray[0];
                            NumRows = (int)Store_Id.ItemArray[1];
                            OOSReport report1 = new OOSReport(Store_ID, ref Loaded, Type, NumRows);

                            if (Loaded)
                            {
                                DateTime GenDate = DateTime.Now;
                                string ReportDate = GenDate.ToString("yyyyMMdd");
                                string filename = "Out Of Stock Report_" + Store_ID.ToString() + "_" + ReportDate + ".pdf";                           
                                string filePath = ConfigurationManager.AppSettings["OutPutPath"];
                                filePath = Path.Combine(filePath, filename);
                                if (!File.Exists(filePath))
                                {
                                    string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;

                                    string defaultFilePath = Path.GetDirectoryName(exePath);
                                    defaultFilePath = Path.Combine(defaultFilePath, filename);
                                    report1.CreateDocument();
                                    report1.ExportToPdf(defaultFilePath);
                                    if (File.Exists(filePath))
                                    {
                                        File.Delete(filePath);
                                    }
                                  
                                    File.Move(defaultFilePath, filePath);
                                  
                                    NumReports++;
                                    Analytic.GetNumReports(NumReports);
                                    DataRow[] selection = LoadedTable.Select($"StoreId={Store_Id[0]}");
                                    LoadedTable.Rows.Remove(selection[0]);

                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {

                        continue;
                    }
                }
                return LoadedTable;
            }
            catch (FileNotFoundException ex)
            {
                Logs.GeneralError = ex.ToString();
                string FileEntry = Logs.LogGeneralErrorReport();
                Logs.LogsCount++;
                string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                logPath = Path.Combine(logPath, "Logs.txt");
                eMail.SendeMail(mailingList, emailFrom, "OOS Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                return LoadedTable;
            }
        }
        #endregion

        #region Create Remaining Reports

        //Same as CreateReport method except using the Remaining Stores values to generate the reports that were not generated on a previous run.
        public static DataTable CreateRemainingReports(DataTable StoreIDV1)
        {
            DataTable LoadedTable = StoreIDV1.Copy();
            int Store_ID = 0;
            int NumReports = 0;
            int Type = 1;
            Analytic.NumberOfNoData = 0;
            try
            {
                foreach (DataRow Store_Id in StoreIDV1.Rows)
                {
                    bool Loaded = false;
                  
                    Store_ID = (int)Store_Id.ItemArray[0];
                    OOSReport report1 = new OOSReport(Store_ID, ref Loaded, Type, 40);

                    if (Loaded)
                    {
                        DateTime GenDate = DateTime.Now;
                        string ReportDate = GenDate.ToString("yyyyMMdd");
                        string filename = "Out Of Stock Report_" + Store_ID.ToString() + "_" + ReportDate + ".pdf";        
                        string filePath = ConfigurationManager.AppSettings["OutPutPath"];
                        filePath = Path.Combine(filePath, filename);
                        
                        if (!File.Exists(filePath))
                        {
                            string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
                            string defaultFilePath = Path.GetDirectoryName(exePath);
                            defaultFilePath = Path.Combine(defaultFilePath, filename);
                            report1.CreateDocument();
                            report1.ExportToPdf(defaultFilePath);
                           
                            if (File.Exists(filePath))
                            {
                                File.Delete(filePath);
                            }
                           
                            File.Move(defaultFilePath, filePath);  
                            NumReports++;
                            Analytic.GetNumReports(NumReports);
                            DataRow[] selection = LoadedTable.Select($"StoreId={Store_Id[0]}");
                            LoadedTable.Rows.Remove(selection[0]);
                        }
                    }
                }
                return LoadedTable;
            }
            catch (FileNotFoundException ex)
            {
                Logs.GeneralError = ex.ToString();
                string FileEntry = Logs.LogGeneralErrorReport();
                Logs.LogsCount++;
                string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                logPath = Path.Combine(logPath, "Logs.txt");
                eMail.SendeMail(mailingList, emailFrom, "OOS Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                return LoadedTable;
            }
        }
        #endregion

        #region Create Supplier Report 
       

        public static DataTable CreateSupplierReport(DataTable SupplierID)
        {
            //Gets the StoreID
            DataTable LoadedTable = SupplierID.Copy();
            int Supplier_ID = 0;
            int checkID = 0;
            int NumReports = 0;           
            SupplierAnalytic.NumberOfNoData = 0;
            try
            {
                foreach (DataRow SupplierId in SupplierID.Rows)
                {
                    try
                    {
                        checkID = (int)SupplierId.ItemArray[0];
                        bool Loaded = false;
                        bool subscribed = false;
                        //Checks if the store is subscribed to the OOS service or not
                        DataTable subscribedTable = SqlFunctionsCoscat4.GetData
                            ($@"DECLARE @supplierId AS INT;
                                SET @supplierId = {checkID};
                                SELECT mps.Unsubscribed
                                //
                             ");

                        if (subscribedTable.Rows.Count == 0)
                        {
                            continue;
                        }
                        DataRow subscribedRow = subscribedTable.Rows[0];
                        string subscribedResult = subscribedRow[0].ToString();
                        if (subscribedResult.Equals("False"))
                        {
                            subscribed = true;
                        }

                        if (subscribed == true)
                        {
                            //Compiles the relavent data needed to create the report and sends the data to the OOSReport class.
                            Supplier_ID = (int)SupplierId.ItemArray[0];
                            SupplierOOSReport report1 = new SupplierOOSReport(Supplier_ID, ref Loaded);

                            if (Loaded)
                            {
                                //Creates the report file and sends it to the correct folder to be stored.
                                DateTime GenDate = DateTime.Now;
                                string ReportDate = GenDate.ToString("yyyyMMdd");
                                string filename = "Out Of Stock Report_" + Supplier_ID.ToString() + "_" + ReportDate + ".pdf";

                                string filePath = ConfigurationManager.AppSettings["OutPutPath"];
                                filePath = Path.Combine(filePath, filename);
                                if (!File.Exists(filePath))
                                {
                                    string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;

                                    string defaultFilePath = Path.GetDirectoryName(exePath);
                                    defaultFilePath = Path.Combine(defaultFilePath, filename);
                                    report1.CreateDocument();
                                    report1.ExportToPdf(defaultFilePath);
                                    if (File.Exists(filePath))
                                    {
                                        File.Delete(filePath);
                                    }

                                    File.Move(defaultFilePath, filePath);

                                    NumReports++;
                                    SupplierAnalytic.GetNumReports(NumReports);
                                    DataRow[] selection = LoadedTable.Select($"SupplierId={SupplierId[0]}");
                                    LoadedTable.Rows.Remove(selection[0]);

                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {

                        continue;
                    }
                }
                return LoadedTable;
            }
            catch (FileNotFoundException ex)
            {
                SupplierLogs.GeneralError = ex.ToString();
                string FileEntry = SupplierLogs.LogGeneralErrorReport();
                SupplierLogs.LogsCount++;
                string logPath = ConfigurationManager.AppSettings["LogFilePath"];
                logPath = Path.Combine(logPath, "Logs.txt");
                eMail.SendeMail(mailingList, emailFrom, "OOS Supplier Report Error Logs", "Please see log file" + "\nLog Generated Date: " + ReportDate, logPath, false, false);
                return LoadedTable;
            }
        }
        #endregion
    }
}


