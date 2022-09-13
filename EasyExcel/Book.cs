using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace EasyExcel;
public class Book
{
    public Book()
    {
        book = new XSSFWorkbook();
    }
    public Book(string path)
    {
        book = new XSSFWorkbook(path);
        Load();
    }
    private void Load()
    {        
        for (int i = 0; i < book.NumberOfSheets; i++)
            sheets.Add(new Sheet(book.GetSheetAt(i)));
    }
    private IWorkbook book;
    private List<Sheet> sheets = new List<Sheet>();
    public Sheet this[string name]
    {
        get
        {
            int index = sheets.ToList().FindIndex(z => z.Name.Equals(name));
            if(index == -1)
            {
                sheets.Add(new Sheet(book.CreateSheet(name)));
                index = sheets.ToList().FindIndex(z => z.Name.Equals(name));
            }
            return sheets[index];
        }
    }
    public void Save(string path) => Save(new FileStream(path.EndsWith(".xlsx") ? path : (path + ".xlsx"), FileMode.OpenOrCreate));
    public void Save(Stream stream) => book.Write(stream);
}