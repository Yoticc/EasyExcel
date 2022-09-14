using NPOI.SS.UserModel;

namespace EasyExcel.Excel;
public class Sheet
{
    public Sheet(ISheet original, Book parent)
    {
        this.original = original;
        Parent = parent;
    }
    private ISheet original;
    public string Name => original.SheetName;
    public Book Parent { get; private set; }
    public bool IsExists(int row, int column)
    {
        IRow irow = original.GetRow(row);
        if (irow == null)
            return false;
        ICell icell = irow.GetCell(column);
        if (icell == null)
            return false;
        return true;
    }
    public Cell this[int row, int column]
    {
        get
        {
            IRow irow = original.GetRow(row);
            if (irow == null)
                irow = original.CreateRow(row);
            ICell icell = irow.GetCell(column);
            if (icell == null)
                icell = irow.CreateCell(column);
            return new Cell(icell);
        }
    }
}