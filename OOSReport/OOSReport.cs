using DevExpress.DataAccess.Sql;
using DevExpress.XtraReports.UI;
using OOSReport.SQL_Classes;
using OOSReport.Analytics;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using DevExpress.Printing.Utils.DocumentStoring;
using System.Collections.Generic;
using System.Text;

namespace OOSReport
{
    public partial class OOSReport : DevExpress.XtraReports.UI.XtraReport
    {
        public OOSReport(int Store, ref bool Loaded, int Type, int NumRows)
        {
            InitializeComponent();
            if (Type == 1) //The type of report indicates whether the report has the standard number of rows or not. Type 1 being the standard.
            {
                DateTime GenDate = DateTime.Now;
                this.xrLabel1.Text = "Selection Status:\nTop 40 Products:\n" +
                    "Base Limit for Selling Percentage >= 75%\nBase Sales Gap Day Count is between 3 and 6";
                DataTable extract = new DataTable();
                string sql;
                sql = $@"EXEC [OOSReport] @storeId = {Store}"; //This executes the OOSReport stored proc in coscat4 db.
                                                               //This stored proc generates all the store and product data in the report.
                Logs.ReportID = Store;
                extract = SqlFunctionsCoscat4.GetData(sql);
                if (extract.Rows.Count > 0)
                {
                    Loaded = true;
                    string storeName = extract.Rows[0][0].ToString();
                    this.ReportTitle.Text = "Out of Stock Report\n for\n " + storeName;
                    DateTime CurrentDate = GenDate.AddDays(-1);
                    this.xrLabel2.Text = "This report up to " + CurrentDate.ToString("dd MMM yyyy");
                    this.DataSource = extract;
                    Logs.ReportStatusOk = "Ok"; //Log that the report was generated successfully
                    string FileEntry = Logs.LogReportOk();
                    Logs.LogsCount++;
                    Logs.CreateFile(FileEntry);
                }
                else
                {
                    Loaded = false;
                    Analytic.NumberOfNoData++;
                    Logs.ReportStatusError = "Failed: No Data"; //Log that the report failed to generate because of no data.
                    string FileEntry = Logs.LogReportError();
                    Logs.CreateFile(FileEntry);
                    Logs.LogsCount++;
                    List<int> StoreIds = new List<int>();
                    StoreIds.Add(Store);
                }
            }
            else
            {
                DateTime GenDate = DateTime.Now;
                this.xrLabel1.Text = "Selection Status:\nTop " + NumRows + " Products:\n" + 
                    "Base Limit for Selling Percentage >= 75%\nBase Sales Gap Day Count is between 3 and 6";
                                                                        //This type of report has a custom number of rows on the report
                                                                        //as requested by that store
                DataTable extract = new DataTable();
                string sql;
                sql = $@"EXEC [OOSReportTopX] @storeId = {Store}, @numrows = {NumRows}";
                                                      
                Logs.ReportID = Store;
                extract = SqlFunctionsCoscat4.GetData(sql);
                if (extract.Rows.Count > 0)
                {
                    Loaded = true;
                    string storeName = extract.Rows[0][0].ToString();
                    this.ReportTitle.Text = "Out of Stock Report\n for\n " + storeName;
                    DateTime CurrentDate = GenDate.AddDays(-1);
                    this.xrLabel2.Text = "This report up to " + CurrentDate.ToString("dd MMM yyyy");
                    this.DataSource = extract;
                    Logs.ReportStatusOk = "Ok";
                    string FileEntry = Logs.LogReportOk();
                    Logs.LogsCount++;
                    Logs.CreateFile(FileEntry);
                }
                else
                {
                    Loaded = false;
                    Analytic.NumberOfNoData++;
                    Logs.ReportStatusError = "Failed: No Data";
                    string FileEntry = Logs.LogReportError();
                    Logs.CreateFile(FileEntry);
                    Logs.LogsCount++;
                    List<int> StoreIds = new List<int>();
                    StoreIds.Add(Store);
                }
            }           
        }
    }
}
