using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.DataAccess.DataFederation;
using OOSReport.Classes;
using OOSReport.SQL_Classes;

namespace OOSReport.Analytics
{
    //Class to log errors and whether a report was successfully generated or not. These logs are stored on the LogFilePath location in App.config.
    public class SupplierLogs
    {
        public static string ReportStatusOk;
        public static int ReportID;
        public static string ReportStatusError;
        public static int LogsCount;
        public static string GeneralError;
        public static int LastID;
        
        public SupplierLogs(string reportStatusOk, int logsCount, string reportStatusError, string generalError, int reportID, int lastID)
        {
            ReportStatusOk = reportStatusOk;
            ReportStatusError = reportStatusError;
            LogsCount = logsCount;
            GeneralError = generalError;
            ReportID = reportID;
            LastID = lastID;
        }
     
        public static void CreateFile(string FileEntry)
        {
            //.. / .. / OOSReport / OOSReport / bin / Debug /
            //"C:/Users/FeldmanJ.SPAR/OneDrive - The SPAR Group Ltd/source/repos/OOSReport/OOSReport/bin/Debug/"
            string docPath = ConfigurationManager.AppSettings["SupplierLogFilePath"];
            using (StreamWriter outPutFile = new StreamWriter(Path.Combine(docPath, "Logs.txt"), true))
            {
                outPutFile.WriteLine(FileEntry);
            }               
        }

        public static string LogReportOk()
        {
            string logReport = "Report status for Supplier ID " + ReportID + ": " + ReportStatusOk;          
            Console.WriteLine(logReport);
            return logReport;
        }

        public static string LogReportError()
        {
            string logReport = "Report status for Supplier ID " + ReportID + ": " + ReportStatusError;
            Console.WriteLine(logReport);
            return logReport;
        }

        public static string LogGeneralErrorReport()
        {
            string CurrentDateTime = DateTime.Now.ToString("hh:mm tt dd/MM/yyyy");
            string report = "The program failed with the error " + GeneralError + "\nAt this date and time: " + CurrentDateTime;
            Console.WriteLine(report);            
            return report;
        }
    }
}
