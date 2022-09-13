using NPOI.SS.UserModel;

namespace EasyExcel;
public class Sheet
{
    public Sheet(ISheet original)
    {
        this.original = original;
    }
    private ISheet original;
    public string Name => original.SheetName;
    public Cell this[int row, int column]
    {
        get
        {
            IRow irow = original.GetRow(row);
            if(irow == null)
                irow = original.CreateRow(row);
            ICell icell = irow.GetCell(column);
            if (icell == null)
                icell = irow.CreateCell(column);
            return new Cell(icell);
        }
    }
}