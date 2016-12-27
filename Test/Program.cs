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
            XLSXReader reader = new XLSXReader("read.xlsx");
            while (reader.ReadNextCell())
            {
                Console.WriteLine(reader.GetCellValue(reader.CurrentCell));
            }
            reader.Close();

            XLSXWriter writer = new XLSXWriter("write.xlsx");
            writer.Start();

            writer.Write("Test");
            writer.WriteInline("more test");
            writer.Write(5);

            writer.Finish();
            writer.Close();
        }
    }
}
