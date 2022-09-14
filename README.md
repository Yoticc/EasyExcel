Example:

```
using EasyExcel;
Book book = new Book();
book["List #1"][0, 0].Value = "My money (debt ;( )";
book["List #1"][1, 0].Value = long.MaxValue;
book.Save("C:\\report"); // Will create C:\\report.xlsx
```
