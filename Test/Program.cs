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
            XLSXReader reader = new XLSXReader("Calculations.xlsx");
            while (reader.ReadNextCell())
            {
                // Console.WriteLine(reader.GetCellValue(reader.CurrentCell));
            }
            reader.Close();
            */

            XLSXWriter writer = new XLSXWriter("write.xlsx");
            writer.Start();

            writer.Write("Id");
            writer.WriteInline("Product");
            writer.Write("Price");
            writer.NewRow();

            writer.Write(1);
            writer.WriteInline("Apple");
            writer.Write(2.3M);
            writer.NewRow();

            writer.Finish();
            writer.Close();

            /*
            XLSXRowReader reader = new XLSXRowReader(@"C:\Users\Grigorov\Documents\User Created\Apollo\QA\Mass Validation\Validations\Copies\18. Ceded Losses ITD - Current_VE.xlsx");
            string[] record;
            int recordCount = 0;
            while (reader.ReadNextRecord(out record))
            {
                Console.WriteLine(++recordCount);
            }
            reader.Close();
            */
        }
    }
}
