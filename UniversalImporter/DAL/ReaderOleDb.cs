using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Xml.Linq;

/// <summary>
/// Жрет память т.к. "reader = cmd.ExecuteReader();" читает всю таблицу в память
/// </summary>

namespace UniversalImporter.DAL
{
    public class ReaderOleDb : IExcelReader
    {
        private OleDbConnection connection;
        private OleDbDataReader reader;
        private List<DataColumn> columns;
        private DataTable SchemaTable;
        private string sheetName;
        public int ReadedRows { get; private set; }
        private List<string> errors = new List<string>();
        public List<string> Errors
        {
            get { return errors; }
        }

        public bool Init(string fileName, DataTable schemaTable)
        {
            SchemaTable = schemaTable;
            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties =\"Excel 12.0;HDR={1};IMEX=1\"", fileName, (true ? "YES" : "NO"));
            connection = new OleDbConnection(connectionString);
            connection.Open();
            var sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (schemaTable == null)
                return false;
            foreach (DataRow row in sheets.Rows)
            {
                sheetName = row["TABLE_NAME"].ToString();
                break;
            }
            OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", sheetName), connection);
            reader = cmd.ExecuteReader();

            return true;
        }
        public DataTable ReadNext(int count)
        {
            int readed = 0;
            var result = SchemaTable.Clone();
            while (reader.Read())
            {
                try
                {
                    var dataRow = result.NewRow();
                    for (int i = 0; i < result.Columns.Count; i++)
                    {
                        dataRow[result.Columns[i]] = reader[i];
                    }
                    result.Rows.Add(dataRow);
                }
                catch (Exception ex)
                {
                    errors.Add(string.Format("Ошибка в строке {0}:{1}", ReadedRows+1, ex.Message));
                }
                readed++;
                ReadedRows++;
                if (readed == count)
                    break;
            }
            return result;
        }
        public XDocument ReadNextX(int count)
        {
            var table = ReadNext(count);

            var content = new XElement("Rows", 
                table.AsEnumerable().Select(row => new XElement("Row", 
                    table.Columns.Cast<DataColumn>().Select(col => new XElement(col.ColumnName, row[col])))));
            return new XDocument(new XElement("ROOT", content));
        }
        public int RowsCount()
        {
            OleDbCommand cmd = new OleDbCommand(string.Format("select count(*) from [{0}]", sheetName), connection);
            return (int)cmd.ExecuteScalar();
        }
        public int ColumnsCount()
        {
            return reader.GetSchemaTable().Rows.Count;
        }

    }
}
