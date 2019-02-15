using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Windows;
using System.IO;
using System.Xml.Linq;
using System.Data;
using System.Data.SqlTypes;

namespace UniversalImporter.DAL
{
    public class DataAccessLayer
    {
        private string TableName;
        private string FileName;
        private string ConnectionString;
        public DataAccessLayer(string tableName, string fileName, string connectionString)
        {
            TableName = tableName;
            FileName = fileName;
            ConnectionString = connectionString;
        }
        public void SaveXML(DateTime dateBeg, DateTime dateEnd)
        {
            var excel = new ExcelQueryFactory(FileName);
            var sheetName = excel.GetWorksheetNames().FirstOrDefault();
            var columnsCount = excel.GetColumnNames(sheetName).Count();
            if (TableColumnsCount() != columnsCount)
            {
                MessageBox.Show("Не совпадает кол-во колонок в таблицах!");
                return;
            }
            
            // Filtering data
            var rows = excel.WorksheetNoHeader(sheetName)
                .Skip(1)
                //.Where(w => w[0].Cast<DateTime>() >= dateBeg && w[0].Cast<DateTime>() <= dateEnd)
                .ToList();

            var columnNames = GetColumnNames();

            // Creating XML
            var content = new XElement("Rows",
              rows.Select(line => new XElement("Row",
                  line.Select((column, index) => new XElement(columnNames[index], column)))));
            var xml = new XDocument(new XElement("ROOT", content));

            //xml.Save(@"D:\out.xml");

            // Inserting
            InsertFromXML(xml);
        }

        #region MS SQL helpers
        private int TableColumnsCount()
        {
            int result = 0;
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return result;
                }
                string query = "select count(*) from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='" + TableName + "'";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    result = (Int32)command.ExecuteScalar();
                }
            }
            return result;
        }
        public static ObservableCollection<string> RefreshTables(string ConnectionString)
        {
            List<String> columnData = new List<String>();
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return null;
                }
                string query = "SELECT top 50 TABLE_NAME FROM INFORMATION_SCHEMA.TABLES order by TABLE_NAME";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnData.Add(reader.GetString(0));
                        }
                    }
                }
            }
            return new ObservableCollection<string>(columnData);
        }
        private List<string> GetColumnNames()
        {
            List<String> columnNames = new List<String>();
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return null;
                }
                string query = "select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='"+TableName+"'";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnNames.Add(reader.GetString(0));
                        }
                    }
                }
            }
            return columnNames;
        }
        private void InsertFromXML(XDocument doc)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                using (SqlCommand command = new SqlCommand("[dbo].[sp_ImportDataFromXML]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar, 50) { Value=TableName } );
                    command.Parameters.Add(new SqlParameter("@Xml", SqlDbType.Xml) { Value = new SqlXml(doc.CreateReader()) });
                    command.Parameters.Add(new SqlParameter("@RowsInserted", SqlDbType.BigInt) { Direction= ParameterDirection.Output });
                    command.Parameters.Add(new SqlParameter("@CurrentStatement", SqlDbType.NVarChar, 5000) { Direction = ParameterDirection.Output });
                    try
                    {
                        var timeBeg = DateTime.Now;
                        var result = command.ExecuteScalar();
                        var timeEnd = DateTime.Now;

                        var rowsInserted = (Int64)command.Parameters["@RowsInserted"].Value;
                        MessageBox.Show("Вставлено " + rowsInserted + " записей. За : "+(timeEnd-timeBeg).TotalSeconds+" секунд.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка:"+ex.Message+" в выражении:"+ command.Parameters["@CurrentStatement"].Value);
                    }
                }
            }
        }



        #endregion


    }
}
