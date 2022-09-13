using EasyExcel;
Book book = new Book();
book["List #1"][0, 0].Value = "test";
book["List #1"][0, 1].Value = long.MaxValue;
book.Save("C:\\test");