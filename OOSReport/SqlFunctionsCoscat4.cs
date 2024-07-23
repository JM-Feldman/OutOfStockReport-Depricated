using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Xml;
using System.Runtime.InteropServices;
//using DTS;

namespace OOSReport.SQL_Classes
{
	/// <summary>
	/// Created by:
	///		Rhyno Linde, 04 May 2004
	///		
	/// Purpose:
	///		Controlled access to SQL to retrieve data
	/// </summary>
	public class SqlFunctionsCoscat4
	{
		#region Declarations 
		private SqlConnection _connection;
		private const int _sqlConnTimeout = 60000; // timeout for this class
		#endregion

		#region Constructor
		public SqlFunctionsCoscat4()
		{	
			_connection = new SqlConnection(ConnectionString);			
		}
		#endregion

		#region Destructor
		~SqlFunctionsCoscat4()
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

		#region ConnectionString
		public static string ConnectionString
		{
			get 
			{
               
               return "//";
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
			if(_connection.State != ConnectionState.Open)
			{
				_connection.Open();
			}
		}
		#endregion

		#region Close Connection
		public void CloseConnection()
		{
			// Close connection
			if(_connection.State != ConnectionState.Closed)
			{
				_connection.Close();
			}
		}
		#endregion

		#region Make SQL Text
		public static string MakeSQLText (string textValue,bool isNumber)
		{
			return makeSQLTextAsText(textValue,isNumber,true);
		}
		public static string MakeSQLText (string textValue,bool isNumber,bool addInvertedComma)
		{
			return makeSQLTextAsText(textValue,isNumber,addInvertedComma);
		}
		public static string MakeSQLText (object textObject, bool isNumber)
		{
			return makeSQLTextAsText(textObject.ToString(),isNumber,true);
		}
		#endregion

		#region Make SQL Text as Text
		private static string makeSQLTextAsText(string textValue, bool isNumber,bool addInvertedComma)
		{
			// return the text "NULL" if the value is empty and it is a number
            if ((textValue == null || textValue.Trim() == "") & isNumber)
			{
				return "NULL";				
			}
			else
			{
				// remove leading and trailing spaces
				textValue = textValue.Trim();
				// replace single quotes with double quotes
				textValue = textValue.Replace("'", "''").Replace(","," ");
				// return the string with quotes
				if(!isNumber & addInvertedComma)
				{
					textValue = string.Format("'{0}'",textValue);
				}
				return textValue;
			}//if (textValue.Trim() == "")
		}
		#endregion

		#region Get XML
		public static XmlDataDocument GetXML(string sqlQuery)
		{
			// create a cache key
			StringBuilder cacheKey = new StringBuilder();			
			
			XmlDataDocument xmlDocument = new XmlDataDocument();		// xml document to hold return values
			SqlConnection sqlConnection = new SqlConnection();			// SQL Connection to retrieve data
			XmlReader xmlReader;													// xml reader to return data as xml

			// open the connection
			sqlConnection.ConnectionString = ConnectionString;
			sqlConnection.Open();
			
			// set the command to execute on SQL
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection = sqlConnection;
			sqlCommand.CommandType = CommandType.Text;
			sqlCommand.CommandText = sqlQuery;
			sqlCommand.CommandTimeout = _sqlConnTimeout;

			// return the data as XML
			xmlReader = sqlCommand.ExecuteXmlReader();

			// load the data into the xml document
			xmlDocument.Load(xmlReader);

			// close the connection
			sqlConnection.Close();			
				
			// return the retrieved XML
			return xmlDocument;
		}
		#endregion			

		#region Execute SQL No Return Value
		public static void ExecuteNonReturnSQL(string sqlQuery)
		{
			SqlConnection sqlConnection = new SqlConnection();			// SQL Connection to retrieve data
			SqlDataReader dataReader;											// xml reader to return data as xml

			// open the connection
			sqlConnection.ConnectionString = ConnectionString;
			sqlConnection.Open();
		
			// set the command to execute on SQL
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection = sqlConnection;
			sqlCommand.CommandType = CommandType.Text;
			sqlCommand.CommandText = sqlQuery;
			sqlCommand.CommandTimeout = _sqlConnTimeout;

			try
			{
				// return the data
				dataReader = sqlCommand.ExecuteReader();

				// CLose Reader
				dataReader.Close();
			}
			catch (SqlException e)
			{				
				throw e;
			}	
			finally
			{
				// close the connection
				sqlConnection.Close();	
			}

		}

		public void PExecuteNonReturnSQL(string sqlQuery)
		{
			SqlDataReader dataReader;											// xml reader to return data as xml
		
			// set the command to execute on SQL
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection = _connection;
			sqlCommand.CommandType = CommandType.Text;
			sqlCommand.CommandText = sqlQuery;
			sqlCommand.CommandTimeout = _sqlConnTimeout;

			try
			{
				// return the data
				dataReader = sqlCommand.ExecuteReader();

				// CLose Reader
				dataReader.Close();
			}
			catch (SqlException e)
			{				
				throw e;
			}		
			
		}
		#endregion		
	
		#region ExecuteSQL
		public static object ExecuteSQL(string sqlQuery)
		{
			object returnObject = null;
			SqlConnection sqlConnection = new SqlConnection();			// SQL Connection to retrieve data												// xml reader to return data as xml

			// open the connection
			sqlConnection.ConnectionString = ConnectionString;
			sqlConnection.Open();
		
			// set the command to execute on SQL
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection = sqlConnection;
			sqlCommand.CommandType = CommandType.Text;
			sqlCommand.CommandText = sqlQuery;
			sqlCommand.CommandTimeout = _sqlConnTimeout;

			try
			{
				// return the data
				returnObject = sqlCommand.ExecuteScalar();
			}
			catch (SqlException e)
			{				
				throw e;
			}	
			finally
			{
				// close the connection
				sqlConnection.Close();	
			}
			return returnObject;	
		}

		public object PExecuteSQL(string sqlQuery)
		{
			object returnObject = null;
					
			// set the command to execute on SQL
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection = _connection;
			sqlCommand.CommandType = CommandType.Text;
			sqlCommand.CommandText = sqlQuery;
			sqlCommand.CommandTimeout = _sqlConnTimeout;

			try
			{
				// return the data
				returnObject = sqlCommand.ExecuteScalar();
			}
			catch (SqlException e)
			{				
				throw e;
			}
			return returnObject;	
		}
		#endregion			

		#region Bulk Upload CSV
		public static void BulkUploadCSV(string FileName, string DataTable)
		{
			BulkUpload(FileName,DataTable,",",null,2);
		}
		#endregion

		#region Bulk Upload File
		public static void BulkUpload(string FileName, string DataTable, string FieldTerminator)
		{
			BulkUpload(FileName,DataTable,FieldTerminator,null,1);
		}

        public static void BulkUpload(string FileName, string DataTable, string FieldTerminator, string RowTerminator, int FirstRow)
        {
            BulkUpload(FileName, DataTable, FieldTerminator, RowTerminator, FirstRow, null);
        }

		public static void BulkUpload(string FileName, string DataTable, string FieldTerminator, string RowTerminator, int FirstRow, string formatFile)
		{
			SqlConnection sqlConnection = new SqlConnection(ConnectionString);

			string 	Sql = string.Format("BULK INSERT [{0}].dbo.[{1}]",sqlConnection.Database,DataTable);
			Sql += "FROM '" + FileName + "'";
			Sql += " WITH ";
			Sql += " (";
			Sql += " FIELDTERMINATOR = '" + FieldTerminator + "'";
			Sql += ", FIRSTROW  = " + FirstRow.ToString();
			if(RowTerminator != null)
			{
				Sql += ", ROWTERMINATOR  = '" + RowTerminator + "'";
			}
            if (formatFile != null)
            {
                Sql += ", FORMATFILE  = '" + formatFile + "'";
            }
			Sql += ")";

			// Execute 
			ExecuteNonReturnSQL(Sql);
		}
		#endregion
//
		#region Execute DTS Package
//		public static void ExecuteDTS(string server, string username, string password, string packageName)
//		{
//			try
//			{
//				Package2Class package = new Package2Class();
//				object pVarPersistStgOfHost = null;
//
//				/* if you need to load from file
//				package.LoadFromStorageFile(
//					"c:\\TestPackage.dts",
//					null,
//					null,
//					null,
//					"Test Package",
//					ref pVarPersistStgOfHost);
//				*/
//
//				package.LoadFromSQLServer(
//					server,
//					username,
//					password,
//					DTSSQLServerStorageFlags.DTSSQLStgFlag_UseTrustedConnection,
//					null,
//					null,
//					null,
//					packageName,
//					ref pVarPersistStgOfHost);
//
//				package.Execute();
//				package.UnInitialize();
//
//				// force Release() on COM object
//				// 
//				System.Runtime.InteropServices.Marshal.ReleaseComObject(package);
//				package = null;
//			}
//			catch(System.Runtime.InteropServices.COMException e)
//			{	
//				throw e;
//			}
//			catch(System.Exception e)
//			{
//				throw e;
//			}
//
//		}
		#endregion
		
		#region GetData
		public static DataTable GetData(string query)
		{
			DataTable table = new DataTable();
			SqlConnection sqlConnection = new SqlConnection();

			// open the connection
			sqlConnection.ConnectionString = ConnectionString;
			sqlConnection.Open();

			// Create the command
			SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
			command.SelectCommand.CommandTimeout = _sqlConnTimeout;
			try
			{
				command.Fill(table);
			}
			catch(Exception e)
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
				catch(Exception e)
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
				catch(Exception e)
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
	}
}
