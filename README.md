# XLSXTools
OpenXML solution to read/write xlsx files.

## How to use?
XLSXTools contains three usable classes - XLSXReader, XLSXWriter and XLSXUtils. 
The first two are used for reading and writing xlsx files respectively and the third one 
has a few utility methods. 

<i>Just make sure to add references to DocumentFormat.OpenXml andWindowsBase.</i>

## XLSXReader
This class is pretty straightforward. 
Keep in mind that the open xml format is such that if a cell is empty then it is omitted
entirely. Track cell addresses to find if a given cell has been skipped.

```C#
XLSXReader reader = new XLSXReader("file.xlsx");
while (reader.ReadNextCell())
{
    Console.WriteLine(reader.GetCellValue(reader.CurrentCell));
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

writer.Write("Test");
writer.WriteInline("more test");
writer.Write(5);

writer.Finish();
writer.Close();
```
### Note:
The WriteInline() method is a lot faster, but the resulting file gets bigger 
because shared strings are not used.
