using EasyExcel.Struct;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test;
public class Player : ExcelStruct
{
    [ExcelVariable]
    public int ID { get; set; }
    [ExcelVariable]
    public string Name { get; set; }
}