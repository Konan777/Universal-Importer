﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Xml.Linq;

/// <summary>
/// Кушает память т.к. "reader = cmd.ExecuteReader();" читает всю таблицу в память
/// </summary>

namespace UniversalImporter.DAL
{
    public class ReaderOleDb : IReader
    {
        private OleDbConnection _connection;
        private OleDbDataReader _reader;
        private List<DataColumn> _columns;
        private DataTable _schemaTable;
        private string _sheetName;
        private List<string> _errors = new List<string>();
        public int ReadedRows { get; private set; }
        public List<string> Errors
        {
            get { return _errors; }
        }

        public bool Init(string fileName, DataTable schemaTable)
        {
            _schemaTable = schemaTable;
            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties =\"Excel 12.0;HDR={1};IMEX=1\"", fileName, (true ? "YES" : "NO"));
            _connection = new OleDbConnection(connectionString);
            _connection.Open();
            var sheets = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (schemaTable == null)
                return false;
            foreach (DataRow row in sheets.Rows)
            {
                _sheetName = row["TABLE_NAME"].ToString();
                break;
            }
            OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", _sheetName), _connection);
            _reader = cmd.ExecuteReader();

            return true;
        }
        public DataTable ReadNext(int count)
        {
            int readed = 0;
            var result = _schemaTable.Clone();
            while (_reader.Read())
            {
                try
                {
                    var dataRow = result.NewRow();
                    for (int i = 0; i < result.Columns.Count; i++)
                    {
                        dataRow[result.Columns[i]] = _reader[i];
                    }
                    result.Rows.Add(dataRow);
                }
                catch (Exception ex)
                {
                    _errors.Add(string.Format("Ошибка в строке {0}:{1}", ReadedRows+1, ex.Message));
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
        public int RowsCount 
        {
            get
            {
                OleDbCommand cmd = new OleDbCommand(string.Format("select count(*) from [{0}]", _sheetName), _connection);
                return (int)cmd.ExecuteScalar();
            }
        }
        public int ColumnsCount  
        {
            get { return _reader.GetSchemaTable().Rows.Count; }
        }

    }
}
