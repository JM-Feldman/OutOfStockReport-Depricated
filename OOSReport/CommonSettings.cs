using OOSReport.SQL_Classes;
using System;
using System.Collections.Generic;

namespace OOSReport.Classes
{
    /// <summary>
    ///     Summary description for CommonSettings.
    /// </summary>
    public class CommonSettings
    {
        #region GetSetting

        public static string GetSetting(string setting)
        {
            var dt = SqlFunctions.GetData(string.Format("SELECT [Value] FROM AppSettings WHERE [Key] = '{0}'", setting));
            if(dt.Rows.Count == 0)
            {
                throw new Exception(string.Format("The setting '{0}' does not exist.", setting));
            }
            return dt.Rows[0]["Value"].ToString();
        }

        #endregion

        #region SetSetting

        public static void SetSetting(string setting, string value)
        {
            List<KeyValuePair<string, object>> values = new List<KeyValuePair<string, object>>();
            values.Add(new KeyValuePair<string, object>("@setting", setting));
            values.Add(new KeyValuePair<string, object>("@value", value));
            SqlFunctions.ExecuteNonReturnSQLSafe(new SQLQuery("UPDATE AppSettings SET [Value] = @setting WHERE [Key] = @value", values));
        }

        #endregion
    }
}