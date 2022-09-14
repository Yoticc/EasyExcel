using EasyExcel.Excel;
using EasyExcel.Struct;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test;
public class PlayerCollection : ExcelCollection<Player>
{
    public PlayerCollection(Sheet sheet) : base(sheet)
    {

    }
}