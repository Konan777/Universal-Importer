using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Xml.Linq;

// EPPlus package. Разницы в производительности с ACE.OLEDB нет. Кушает память.
// ExcelDataReader - то что доктор прописал. память не кушает!

namespace UniversalImporter.DAL
{
    public class ReaderEPPlus : IExcelReader
    {
        private FileStream stream;
        private IExcelDataReader reader;
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

            stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            reader = ExcelReaderFactory.CreateReader(stream);
            reader.Read(); // skip header
            ReadedRows = 1;
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
                    var row = result.NewRow();
                    for (int i=0; i < result.Columns.Count; i++) row[result.Columns[i]] = reader.GetValue(i);
                    result.Rows.Add(row);
                }
                catch (Exception ex)
                {
                    errors.Add(string.Format("Ошибка в строке {0}:{1}", ReadedRows + 1, ex.Message));
                }
                readed++;
                ReadedRows++;
                if (readed == count)
                    break;
            }
            if (ReadedRows == RowsCount())
                stream.Close();
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
            return reader.RowCount;
        }
        public int ColumnsCount()
        {
            return reader.FieldCount;
        }

    }
}

