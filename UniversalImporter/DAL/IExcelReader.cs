using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace UniversalImporter.DAL
{
    public interface IReader
    {
        bool Init(string fileName, DataTable shemaTable);
        DataTable ReadNext(int count);
        XDocument ReadNextX(int count);
        int RowsCount { get; }
        int ReadedRows { get; }
        int ColumnsCount { get; }
        List<string> Errors { get; }
    }

}
