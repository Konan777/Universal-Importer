using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Windows;
using System.Xml.Linq;
using System.Data;
using System.Data.SqlTypes;
using Z.EntityFramework.Extensions;

namespace UniversalImporter.DAL
{
    public enum SaveType { SaveXML, SaveBulk, SaveEF }
    public class DataAccessLayer
    {
        private string _tableName;
        private string _connectionString;
        private List<string> _columnNames;
        private SqlConnection _connection;
        private MSSQL_Database database;

        public DataAccessLayer(string tableName, string connectionString)
        {
            _tableName = tableName;
            _connectionString = connectionString;
            _connection = new SqlConnection(_connectionString);
            _connection.Open();
        }
        public void Save(IReader reader, DateTime dateBeg, DateTime dateEnd, IProgress<int> progress, SaveType saveType)
        {

            _columnNames = SqlTableColumnNames();

            if (_columnNames.Count != reader.ColumnsCount)
                throw new Exception("Не совпадает кол-во колонок в таблицах!");

            var timeBeg = DateTime.Now;

            var rowsCount = reader.RowsCount;
            int step = (rowsCount / 100);
            if (step == 0)
                step = rowsCount;
            while (reader.ReadedRows < rowsCount)
            {
                switch (saveType)
                {
                    case SaveType.SaveBulk:
                        var table = reader.ReadNext(step);
                        BulkSave(table);
                        progress.Report(reader.ReadedRows);
                        break;
                    case SaveType.SaveXML:
                        var doc = reader.ReadNextX(step);
                        InsertFromXML(doc);
                        progress.Report(reader.ReadedRows);
                        break;
                    case SaveType.SaveEF:
                        var table2 = reader.ReadNext(step);
                        ExtensionsSave(table2);
                        progress.Report(reader.ReadedRows);
                        break;
                }
            }
           _connection.Close();

            var timeEnd = DateTime.Now;
            var result = (timeEnd - timeBeg).TotalSeconds;
            var realRows = rowsCount - reader.Errors.Count;
            MessageBox.Show(
                string.Format("Вставлено {0} из {1} ошибок {2} за {3} секунд.", realRows, rowsCount, reader.Errors.Count, result)
            );
        }


        #region MS SQL helpers
        public static ObservableCollection<string> RefreshTables(string connectionString)
        {
            List<String> columnData = new List<String>();
            using (SqlConnection connection = new SqlConnection(connectionString))
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
                string query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES order by TABLE_NAME";
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
            string query = "select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='"+_tableName+"'";
            using (SqlCommand command = new SqlCommand(query, _connection))
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
            using (SqlCommand command = new SqlCommand("[dbo].[sp_ImportDataFromXML]", _connection))
            {
                command.CommandTimeout = 360;
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar, 50) { Value=_tableName } );
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
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(_connection))
            {
                bulkCopy.DestinationTableName = "[" + _tableName + "]";
                bulkCopy.WriteToServer(table);
            }
        }

        private void ExtensionsSave(DataTable table)
        {
            database = database ?? new MSSQL_Database();
            var context = database.GetNewContext();
            var rows = new List<AAA_300>();
            foreach (DataRow row in table.Rows)
            {
                var _row = new AAA_300()
                {
                    Id = 0,
                    DT01 = (row["DT01"] == null) ? (DateTime?)null : DateTime.Parse(row["DT01"].ToString()),
                    TXT01 = row["TXT01"].ToString(),
                    TXT02 = row["TXT02"].ToString(),
                    TXT03 = row["TXT03"].ToString(),
                    TIME01 = row["TIME01"].ToString(),
                    TIME02 = row["TIME02"].ToString(),
                    COST = (row["COST"] == null) ? (double?)null : double.Parse(row["COST"].ToString())
                };
                rows.Add(_row);
            }

            //context.AAA_100.AddRange(rows);
            //context.SaveChanges();
            context.BulkInsert(rows);


        }
        public DataTable GetSqlTableSchema()
        {
            var result = new DataTable();
            string query = String.Format("SELECT TOP 1 * FROM {0}", _tableName);
            using (SqlCommand command = new SqlCommand(query, _connection))
            {
                SqlDataReader reader = command.ExecuteReader(CommandBehavior.SchemaOnly);

                result.Load(reader);
            }
            return result;
        }
        #endregion


    }
}
