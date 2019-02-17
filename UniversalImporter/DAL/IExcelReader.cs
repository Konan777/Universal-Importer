using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace UniversalImporter.DAL
{
    public interface IExcelReader
    {
        bool Init(string fileName, DataTable shemaTable);
        DataTable ReadNext(int count);
        XDocument ReadNextX(int count);
        int RowsCount();
        int ReadedRows { get; }
        int ColumnsCount();
        List<string> Errors { get; }
    }

}
