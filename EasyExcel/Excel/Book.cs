using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace EasyExcel;
public class Book
{
    public Book(string pathToSave = "")
    {
        Path = pathToSave;
        book = new XSSFWorkbook();
    }

    private IWorkbook book;
    private List<Sheet> sheets = new();

    public string Path;

    public static Book Load(string path)
    {
        var book = new Book(path);
        book.LoadBook(path);
        return book;
    }

    private void LoadBook(string path)
    {
        book = new XSSFWorkbook(path);
        for (int i = 0; i < book.NumberOfSheets; i++)
            sheets.Add(new Sheet(book.GetSheetAt(i), this));
    }

    public void Save(string path) => Save(new FileStream(path.EndsWith(".xlsx") ? path : (path + ".xlsx"), FileMode.OpenOrCreate));
    public void Save(Stream stream) => book.Write(stream);
    public void Save() => Save(Path);

    public Sheet this[string name]
    {
        get
        {
            int index = sheets.ToList().FindIndex(z => z.Name.Equals(name));
            if(index == -1)
            {
                sheets.Add(new Sheet(book.CreateSheet(name), this));
                index = sheets.ToList().FindIndex(z => z.Name.Equals(name));
            }
            return sheets[index];
        }
    }
}