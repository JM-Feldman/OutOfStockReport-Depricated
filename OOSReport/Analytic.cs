using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.BarCodes;
using DevExpress.Charts.Native;
using DevExpress.Printing.Utils.DocumentStoring;
using DevExpress.ReportServer.ServiceModel.DataContracts;
using OOSReport;
using OOSReport.Classes;
using static DevExpress.DataProcessing.InMemoryDataProcessor.AddSurrogateOperationAlgorithm;
using OOSReport.Classes;
using OOSReport.SQL_Classes;
using System.Data.SqlClient;

namespace OOSReport.Analytics
{
    //Class to gather data about the reports such as how many were generated and to email that data to users on the mailing list.
    public class Analytic
    {
        public static int NumberOfStores;
        public static int NumberOfReports;
        public static int NumberOfNoData;

        public Analytic(int numberOfStores, int numberOfReports, int numberOfNoData)
        {
            NumberOfStores = numberOfStores;
            NumberOfReports = numberOfReports;
            NumberOfNoData = numberOfNoData;
        }

        public static void GetNumStores(int NumStores)
        {          
             NumberOfStores = NumStores;
        }
        
        public static void GetNumReports(int NumReports)
        {            
            NumberOfReports = NumReports;
        }

        public static string AnalyticsStatsReport()
        {
            string Report = "";
            if (NumberOfReports == 0)
            {
                Report = "ALERT: No reports generated. Please check for errors.";              
            }
            else
            {
                DateTime GenDate = DateTime.Now;
                string CurrentDateTime = GenDate.ToString("yyyy-MM-dd HH:mm:ss");
                Report = "Number of stores Queried: " + NumberOfStores.ToString() + "\nNumber of reports Generated: " + NumberOfReports.ToString() + "\nNumber of stores with no data: "
                + NumberOfNoData.ToString();
                string sql = $@"INSERT INTO // VALUES                                                             (
                {NumberOfStores},{NumberOfReports},{NumberOfNoData},'{CurrentDateTime}')";                              
                SqlFunctionsCoscat4.ExecuteNonReturnSQL(sql);               
            }
            return Report;
        }

        public static string AnalyticsStatsEmail()
        {
            string Report = "";
            if(NumberOfReports == 0)
            {
                 Report = "ALERT: No reports generated. Please check for errors.";        
            }
            else
            {
                DateTime GenDate = DateTime.Now;
                string CurrentDateTime = GenDate.ToString("yyyy-MM-dd");
                Report = "Number of stores Queried: " + NumberOfStores.ToString() + "\nNumber of reports Generated: " + NumberOfReports.ToString() + "\nNumber of stores with no data: "
                + NumberOfNoData.ToString();                             
            }                  
            return Report;
        }
    }
}
