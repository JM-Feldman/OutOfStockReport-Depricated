using System.Data;
using System.Text;
using System.Xml;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Collections.Generic;
using System;

namespace OOSReport.SQL_Classes
{
    public class SqlFunctions
    {
        #region Declarations 
        private SqlConnection _connection;
        private const int _sqlConnTimeout = 60000; // timeout for this class
        private static SqlConnection staticSqlConnection;
        private static string ConnectionString { get; set; } = "Data Source=coscat4;Initial Catalog=SparCE;User Id=process_3p_DataImport;Password=Th1rdp@rty@9876; Packet Size=4096;";
        #endregion

        #region SetConnectionCredentials
        public static void SetConnectionCredentials(string server, string database, string username, string password)
        {
            ConnectionString = $"Data Source={server};Initial Catalog={database};User Id={username};Password={password}; Packet Size=4096;";
        }
        #endregion

        #region Destructor
        ~SqlFunctions()
        {
            _connection.Close();
            _connection.Dispose();
        }
        #endregion

        #region ConnectionTimeOut
        public static int ConnectionTimeOut
        {
            get
            {
                return _sqlConnTimeout;
            }
        }
        #endregion

        #region Connection
        public SqlConnection Connection
        {
            get
            {
                return _connection;
            }
            set
            {
                _connection.Close();
                _connection = value;
            }
        }
        #endregion

        #region Open Connection
        public void OpenConnection()
        {
            // Open connection
            if (_connection.State != ConnectionState.Open)
            {
                _connection.Open();
            }
        }
        #endregion

        #region Close Connection
        public void CloseConnection()
        {
            // Close connection
            if (_connection.State != ConnectionState.Closed)
            {
                _connection.Close();
            }
        }
        #endregion

        #region OpenStaticConnection
        public static void OpenStaticConnection()
        {
            staticSqlConnection = new SqlConnection();
            staticSqlConnection.ConnectionString = ConnectionString;
            try
            {
                staticSqlConnection.Open();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region CloseStaticConnection
        public static void CloseStaticConnection()
        {
            try
            {
                staticSqlConnection.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region GetDataFromStatic
        public static DataTable GetDataFromStatic(SQLQuery query)
        {
            DataTable table = new DataTable();
            try
            {
                // Create the command
                SqlDataAdapter command = new SqlDataAdapter(query.sql, staticSqlConnection);
                command.SelectCommand.CommandTimeout = _sqlConnTimeout;
                if (query.values != null)
                {
                    foreach (KeyValuePair<string, object> param in query.values)
                    {
                        command.SelectCommand.Parameters.AddWithValue(param.Key, param.Value);
                    }
                }
                try
                {
                    command.Fill(table);
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    // Release Resources
                    command.Dispose();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.Message);
                throw e;
            }
            return table;
        }
        #endregion

        #region GetDataParameterized
        public static DataTable GetDataParameterized(SQLQuery query)
        {
            DataTable table = new DataTable();
            SqlConnection sqlConnection = new SqlConnection();
            SqlConnection sqlConnCos4 = new SqlConnection();

            // open the connection
            sqlConnection.ConnectionString = "Data Source=coscat4;Initial Catalog=SparCE;User Id=process_intouch;Password=1Ntouch7864; Packet Size=4096;"; 
            sqlConnection.Open();

           /* sqlConnCos4.ConnectionString = "Data Source=catsrv-coscat4;Initial Catalog=Daily Data;User Id=_CatmanDev;Password=Develop246;Packet Size=4096;";
            sqlConnCos4.Open();*/

            // Create the command
            SqlDataAdapter command = new SqlDataAdapter(query.sql, sqlConnection);
            command.SelectCommand.CommandTimeout = _sqlConnTimeout;
            if (query.values != null)
            {
                foreach (KeyValuePair<string, object> param in query.values)
                {
                    command.SelectCommand.Parameters.AddWithValue(param.Key, param.Value);
                }
            }
            try
            {
                command.Fill(table);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                // Release Resources
                command.Dispose();
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
            return table;
        }
        #endregion

        #region Execute SQL No Return Value
        public static void ExecuteNonReturnSQLStatic(SQLQuery query)
        {
            SqlDataReader dataReader;  // xml reader to return data as xml
            SqlCommand sqlCommand = new SqlCommand();
            SqlConnection sqlConnection = new SqlConnection();
            sqlConnection.ConnectionString = ConnectionString;
            sqlConnection.Open();
            sqlCommand.Connection = sqlConnection;
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.CommandText = query.sql;
            sqlCommand.CommandTimeout = _sqlConnTimeout;
            if (query.values != null)
            {
                foreach (KeyValuePair<string, object> param in query.values)
                {
                    sqlCommand.Parameters.AddWithValue(param.Key, param.Value);
                }
            }

            try
            {
                dataReader = sqlCommand.ExecuteReader();
                dataReader.Close();
                sqlCommand.Dispose();
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                 throw e;
            }
        }
        #endregion

        #region Execute SQL No Return Value
        public static void ExecuteNonReturnSQLSafe(SQLQuery query)
        {
            SqlConnection sqlConnection = new SqlConnection();
            sqlConnection.ConnectionString = ConnectionString;
            try
            {
                sqlConnection.Open();
            }
            catch (Exception e)
            {
                throw e;
            }
            SqlDataReader dataReader;  // xml reader to return data as xml
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.Connection = sqlConnection;
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.CommandText = query.sql;
            sqlCommand.CommandTimeout = _sqlConnTimeout;
            if (query.values != null)
            {
                foreach (KeyValuePair<string, object> param in query.values)
                {
                    sqlCommand.Parameters.AddWithValue(param.Key, param.Value);
                }
            }
            Debug.WriteLine(sqlCommand.CommandText);
            try
            {
                dataReader = sqlCommand.ExecuteReader();
                dataReader.Close();
                sqlCommand.Dispose();
                sqlConnection.Close();
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                throw e;
            }
        }
        #endregion

        #region GetData
        public static DataTable GetData(string query, string customConnectionString = "")
        {
            DataTable table = new DataTable();
            SqlConnection sqlConnection = new SqlConnection();

            // open the connection
            sqlConnection.ConnectionString = string.IsNullOrEmpty(customConnectionString) ? ConnectionString : customConnectionString;
            sqlConnection.Open();

            // Create the command
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            command.SelectCommand.CommandTimeout = _sqlConnTimeout;
            try
            {
                command.Fill(table);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                // Release Resources
                command.Dispose();
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
            return table;
        }

        public DataTable PGetData(string query)
        {
            DataTable table = new DataTable();

            // Create the command
            SqlDataAdapter command = new SqlDataAdapter(query, _connection);
            command.SelectCommand.CommandTimeout = _sqlConnTimeout;
            try
            {
                try
                {
                    command.Fill(table);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            catch
            {
                try
                {
                    command.Fill(table);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            finally
            {
                // Release Resources
                command.Dispose();
            }
            return table;
        }
        #endregion

        #region ReturnType
        public static int ReturnInteger(string query)
        {
            int value = 0;

            SqlConnection sqlConnection = new SqlConnection();

            // open the connection
            sqlConnection.ConnectionString = ConnectionString;
            sqlConnection.Open();

            SqlCommand command = new SqlCommand(query, sqlConnection);


            try
            {
                try
                {
                    value = Convert.ToInt32(command.ExecuteScalar());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            catch
            {
                try
                {
                    value = Convert.ToInt32(command.ExecuteScalar());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            finally
            {
                // Release Resources
                command.Dispose();
                sqlConnection.Close();
            }

            return value;
        }

        #endregion

    }
    #region SQLQuery
    public class SQLQuery
    {
        public string sql;
        public List<KeyValuePair<string, object>> values;
        public SQLQuery(string sql, List<KeyValuePair<string, object>> values = null)
        {
            Debug.WriteLine(sql);
            this.sql = sql;
            if (values != null) { this.values = values; }
        }
    }
    #endregion
}
