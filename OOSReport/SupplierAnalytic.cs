using System;
using OOSReport.SQL_Classes;

namespace OOSReport.Analytics
{
    //Class to gather data about the reports such as how many were generated and to email that data to users on the mailing list.
    public class SupplierAnalytic
    {
        public static int NumberOfSuppliers;
        public static int NumberOfReports;
        public static int NumberOfNoData;

        public SupplierAnalytic(int numberOfStores, int numberOfReports, int numberOfNoData)
        {
            NumberOfSuppliers = numberOfStores;
            NumberOfReports = numberOfReports;
            NumberOfNoData = numberOfNoData;
        }

        public static void GetNumStores(int NumStores)
        {          
             NumberOfSuppliers = NumStores;
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
                Report = "Number of Suppliers Queried: " + NumberOfSuppliers.ToString() + "\nNumber of reports Generated: " + NumberOfReports.ToString() + "\nNumber of Suppliers with no data: "
                + NumberOfNoData.ToString();
                string sql = $@"INSERT INTO // VALUES                                                             (
                {NumberOfSuppliers},{NumberOfReports},{NumberOfNoData},'{CurrentDateTime}')";                              
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
                Report = "Number of Suppliers Queried: " + NumberOfSuppliers.ToString() + "\nNumber of reports Generated: " + NumberOfReports.ToString() + "\nNumber of Suppliers with no data: "
                + NumberOfNoData.ToString();                             
            }                  
            return Report;
        }
    }
}
