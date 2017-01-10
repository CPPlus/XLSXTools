# XLSXTools
OpenXML solution to read/write xlsx files.

## How to use?
XLSXTools contains four usable classes - XLSXReader, XLSXRowReader, XLSXWriter and XLSXUtils. 
The first three are used for reading and writing xlsx files respectively and the fourth one 
has a few utility methods. 

<i>Just make sure to add references to DocumentFormat.OpenXml and WindowsBase.</i>

## XLSXReader
This class is pretty straightforward. 
Keep in mind that the open xml format is such that if a cell is empty then it is omitted
entirely. Track cell addresses to find if a given cell has been skipped.

### Basic use:
```C#
XLSXReader reader = new XLSXReader("file.xlsx");
reader.SetSheet("Sheet1"); // Not needed. The first sheet is the default one.
while (reader.ReadNextCell())
{
    Console.WriteLine(reader.GetCellValue(reader.CurrentCell));
    Console.WriteLine(reader.GetCellReference(reader.CurrentCell));
}
reader.Close();
```

### Notable properties.
These are the actual row and column counts respectively. Excel sometimes calculates a wrong
used range, be it formatted cells without a value or some other reason. These properties are 
calculated on instantiation by reading the whole file. If you don't need them and think that
the instantiation is slow when you read big files you can comment this calculation out in the
constructor.
```C#
int actualRowCount = reader.RowCount;
int actualColumnCount = reader.ColumnCount;
```

## XLSXRowReader
This class bypasses the empty cell skipping and reads the file row by row.
Empty cells are replaced with string.Empty.
```C#
XLSXRowReader reader = new XLSXRowReader(@"file.xlsx");
reader.SetSheet("Sheet1"); // Not needed. The first sheet is the default one.
string[] record;
while (reader.ReadNextRecord(out record))
{
    foreach (string field in record)
        Console.Write(field + ", ");
    Console.WriteLine();
}
reader.Close();
```

## XLSXWriter
The writer is also simple. Just make sure to call the Start() method before
writing anything. Also call the Finish() and Close() methods after writing out
your data or the file will become corrupt.

```C#
XLSXWriter writer = new XLSXWriter("write.xlsx");
writer.Start();

writer.Write("This is a shared string text.");
writer.WriteInline("This is an inline text");
writer.Write(5);
writer.Write("This is a colored cell", Styles.BLUE); // Add more styles in XLSXWriter's "WriteWorkbookStylesPart" method.

writer.Finish();
writer.Close();
```
### Note:
The WriteInline() method is a lot faster, but the resulting file can sometimes get 
bigger because shared strings are not used.


