using EasyExcel.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcel.Struct;
public class StructCollection<T> where T : ExcelStruct
{
    public StructCollection(Sheet parent)
    {
        Parent = parent;
        UpdateStructVariables();
        if (parent.IsExists(0, 0))
        {
            int count = Parent[0, 0].IntValue;
            for (int i = 0; i < count; i++)
            {
                
            }
        }
        else
        {
            Parent[0, 0].Value = 0;
            for (int i = 0; i < Variables.Count; i++)
            {
                Parent[i + 1, 0].StringValue = Variables[i];
            }
        }
    }
    public Sheet Parent { get; private set; }
    private List<T> Values = new List<T>();
    public int Count { get => Parent[0, 0].IntValue; private set => Parent[0, 0].IntValue = value; }
    public void Add(T value)
    {
        Values.Add(value);
        int row = Count;
        Type type = value.GetType();
        for(int i = 0; i < Variables.Count; i++)
            Parent[i + 1, row].Value = type.GetField(Variables[i]).GetValue(value);
        Count = Values.Count;
    }
    public T this[int index]
    {
        get => Values[index];
        set => Values[index] = value;
    }
    private List<string> Variables;
    private void UpdateStructVariables()
    {
        Type type = typeof(T);
        Variables = type.GetFields().Where(z => z.CustomAttributes.ToList().Select(z => z.AttributeType).Contains(typeof(ExcelVariableAttribute))).Select(z => z.Name).ToList();
    }
}