using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using withdrawingDataExcel;

namespace ExcelTest
{
    class Testing
    {
        
        static void Main(string[] args)
        {
            Program pr = new Program("C:\\Users\\kirill\\Downloads\\расписание с 6.03.xlsx", 3, "ИП31");
            List<string> strings = pr.sheetsList();
            //Console.WriteLine(pr.findStrings("ВТР"));


            //Console.WriteLine(pr.findIndex("L7"));

            Console.WriteLine(pr.findRowAndColumn(pr.findStrings("ИП31")));


            //Console.WriteLine(pr.returnListPair(pr.findIndexWeek()));
            //Console.WriteLine(pr.findIndex("O7"));

            foreach (var item in pr.returnIndex("ИП31"))
            {
                Console.WriteLine(item);
            }

            foreach (var item in pr.findIndexWeek())
            {
                Console.WriteLine(item);
            }
        }

    }
}
