using NPOI.SS.UserModel;

namespace EasyExcel;
public record Cell(ICell original)
{
    public string Formula { get => original.CellFormula; set => original.SetCellFormula(value); }
    public CellType Type { get => original.CellType; set => original.SetCellType(value); }
    public string StringValue { get => original.StringCellValue; set => original.SetCellValue(value); }
    public bool BoolValue { get => original.BooleanCellValue; set => original.SetCellValue(value); }
    public double DoubleValue { get => original.NumericCellValue; set => original.SetCellValue(value); }
    public float FloatValue { get => (float)original.NumericCellValue; set => original.SetCellValue(value); }
    public int IntValue { get => (int)original.NumericCellValue; set => original.SetCellValue(value); }
    public long LongValue { get => long.Parse(original.StringCellValue); set => original.SetCellValue(value.ToString()); }

    public object Value
    {
        set
        {
            if (value.GetType() == typeof(string))
                StringValue = (string)value;
            else if (value.GetType() == typeof(bool))
                BoolValue = (bool)value;
            else if (value.GetType() == typeof(double))
                DoubleValue = (double)value;
            else if (value.GetType() == typeof(float))
                FloatValue = (float)value;
            else if (value.GetType() == typeof(int))
                IntValue = (int)value;
            else if (value.GetType() == typeof(long))
                LongValue = (long)value;
        }
    }
}