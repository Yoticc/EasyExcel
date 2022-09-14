using EasyExcel.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcel.Struct;
public class ExcelCollection<T> where T : ExcelStruct
{
    public ExcelCollection(Sheet sheet)
    {
        Sheet = sheet;
        Collection = new StructCollection<T>(sheet);
    }
    public Sheet Sheet { get; set; }
    public StructCollection<T> Collection { get; private set; }
    public void Save() => Sheet.Parent.Save();
}