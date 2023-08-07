using NPOI.SS.UserModel;

namespace EasyExcel;
public record Sheet(ISheet original, Book Parent)
{
    public string Name => original.SheetName;

    public bool IsExists(int row, int column)
    {
        var irow = original.GetRow(row);
        if (irow == null)
            return false;

        var icell = irow.GetCell(column);
        if (icell == null)
            return false;

        return true;
    }

    public Cell this[int row, int column]
    {
        get
        {
            var irow = original.GetRow(row);
            if (irow == null)
                irow = original.CreateRow(row);

            var icell = irow.GetCell(column);
            if (icell == null)
                icell = irow.CreateCell(column);

            return new(icell);
        }
    }
}