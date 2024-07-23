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
    public partial class SupplierOOSReport : DevExpress.XtraReports.UI.XtraReport
    {
        public SupplierOOSReport(int SupplierID, ref bool Loaded)
        {
            InitializeComponent();
            DateTime GenDate = DateTime.Now;
            this.xrLabel1.Text = "Selection Status:\nTop 40 Products:\n" +
                "Base Limit for Selling Percentage >= 75%\nBase Sales Gap Day Count is between 3 and 6";
            DataTable extract = new DataTable();
            string sql;
            sql = $@"EXEC [vOOSReport_SupplierVersion] @SupplierId = {SupplierID}"; //This executes the OOSReport stored proc in coscat4 db.
                                                           //This stored proc generates all the store and product data in the report.
            SupplierLogs.ReportID = SupplierID;
            extract = SqlFunctionsCoscat4.GetData(sql);
            if (extract.Rows.Count > 0)
            {
                Loaded = true;
                string supplierName = extract.Rows[0][0].ToString();
                this.ReportTitle.Text = "Out of Stock Report\n for\n " + supplierName;
                DateTime CurrentDate = GenDate.AddDays(-1);
                this.xrLabel2.Text = "This report up to " + CurrentDate.ToString("dd MMM yyyy");
                this.DataSource = extract;
                SupplierLogs.ReportStatusOk = "Ok"; //Log that the report was generated successfully
                string FileEntry = Logs.LogReportOk();
                SupplierLogs.LogsCount++;
                SupplierLogs.CreateFile(FileEntry);
            }
            else
            {
                Loaded = false;
                SupplierAnalytic.NumberOfNoData++;
                SupplierLogs.ReportStatusError = "Failed: No Data"; //Log that the report failed to generate because of no data.
                string FileEntry = Logs.LogReportError();
                SupplierLogs.CreateFile(FileEntry);
                SupplierLogs.LogsCount++;
                List<int> StoreIds = new List<int>();
                StoreIds.Add(SupplierID);
            }
        }

    }
}
