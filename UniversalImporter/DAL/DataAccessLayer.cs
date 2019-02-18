using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Windows;
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
        private List<string> ColumnNames;
        private SqlConnection Connection;

        public DataAccessLayer(string tableName, string fileName, string connectionString)
        {
            TableName = tableName;
            FileName = fileName;
            ConnectionString = connectionString;
            Connection = new SqlConnection(ConnectionString);
            Connection.Open();
        }
        public void Save(IExcelReader reader, DateTime dateBeg, DateTime dateEnd, IProgress<int> Progress, bool Bulk)
        {

            ColumnNames = SqlTableColumnNames();
            var schemaTable = GetSqlTableSchema();

            if (ColumnNames.Count != reader.ColumnsCount)
            {
                MessageBox.Show("Не совпадает кол-во колонок в таблицах!");
                return;
            }
            var timeBeg = DateTime.Now;

            var rowsCount = reader.RowsCount;
            int step = (rowsCount / 100);
            if (step == 0)
                step = rowsCount;
            while (reader.ReadedRows < rowsCount)
            {
                if (Bulk)
                {
                    var table = reader.ReadNext(step);
                    BulkSave(table);
                    Progress.Report(reader.ReadedRows);
                }
                else
                {
                    var doc = reader.ReadNextX(step);
                    InsertFromXML(doc);
                    Progress.Report(reader.ReadedRows);
                }
            }
            Connection.Close();

            var timeEnd = DateTime.Now;
            var result = (timeEnd - timeBeg).TotalSeconds;
            var realRows = rowsCount - reader.Errors.Count;
            MessageBox.Show(
                string.Format("Вставлено {0} из {1} ошибок {2} за {3} секунд.", realRows, rowsCount, reader.Errors.Count, result)
            );
        }


        #region MS SQL helpers
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
        private List<string> SqlTableColumnNames()
        {
            List<String> columnNames = new List<String>();
            string query = "select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='"+TableName+"'";
            using (SqlCommand command = new SqlCommand(query, Connection))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columnNames.Add(reader.GetString(0));
                    }
                }
            }
            return columnNames;
        }
        private void InsertFromXML(XDocument doc)
        {
            using (SqlCommand command = new SqlCommand("[dbo].[sp_ImportDataFromXML]", Connection))
            {
                command.CommandTimeout = 360;
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar, 50) { Value=TableName } );
                command.Parameters.Add(new SqlParameter("@Xml", SqlDbType.Xml) { Value = new SqlXml(doc.CreateReader()) });
                command.Parameters.Add(new SqlParameter("@RowsInserted", SqlDbType.BigInt) { Direction= ParameterDirection.Output });
                command.Parameters.Add(new SqlParameter("@CurrentStatement", SqlDbType.NVarChar, 5000) { Direction = ParameterDirection.Output });
                try
                {
                    var result = command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка:"+ex.Message+" в выражении:"+ command.Parameters["@CurrentStatement"].Value);
                }
            }
        }
        private void BulkSave(DataTable table)
        {
            if (table.Rows.Count == 0) return;
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(Connection))
            {
                bulkCopy.DestinationTableName = "[" + TableName + "]";
                bulkCopy.WriteToServer(table);
            }
        }
        public DataTable GetSqlTableSchema()
        {
            var result = new DataTable();
            string query = String.Format("SELECT TOP 1 * FROM {0}", TableName);
            using (SqlCommand command = new SqlCommand(query, Connection))
            {
                SqlDataReader reader = command.ExecuteReader(CommandBehavior.SchemaOnly);

                result.Load(reader);
            }
            return result;
        }
        #endregion


    }
}
