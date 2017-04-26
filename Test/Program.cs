using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXTools;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            XLSXReader reader = new XLSXReader("read.xlsx", "Sheet2");
            while (reader.ReadNextCell())
            {
                Console.WriteLine(reader.GetCellValue(reader.CurrentCell));
            }
            reader.Close();
            */

            /*
            XLSXWriter writer = new XLSXWriter("write.xlsx");

            // Write a sheet.
            writer.SetWorksheet("MyTestSheet1");

            writer.Write("Id");
            writer.WriteInline("Product");
            writer.Write("Price");
            writer.NewRow();

            writer.Write(1);
            writer.WriteInline("Apple");
            writer.Write(2.3M);
            writer.NewRow();

            // Write another sheet.
            writer.SetWorksheet("MyTestSheetTWO");

            writer.Write("Id");
            writer.NewRow();

            writer.Write(1);
            writer.WriteInline("LOL");
            writer.NewRow();

            writer.SetWorksheet("MyTestSheet1");
            writer.WriteInline("ADDED TEST");

            writer.Finish();
            writer.Close();
            */
            
            XLSXRowReader reader = new XLSXRowReader(@"read.xlsx", "Sheet1", false);

            string[] record;
            while (reader.ReadNextRecord(out record))
            {
                foreach (string field in record)
                    Console.Write(field + ", ");
                Console.WriteLine();
            }

            reader.Close();
        }
    }
}
