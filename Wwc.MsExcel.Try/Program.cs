using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Wwc.MsExcel;

namespace Wwc.MsExcel.Try
{
    class Program
    {
        static void Main(string[] args)
        {
            MsExcelParser parser = new MsExcelParser();
            parser.Parse();

            foreach (var cust in Customer.CustomerList)
            {
                Console.WriteLine(cust.ToString());
            }
            Console.WriteLine("Hit any key to continue");
            Console.ReadKey();
        }
    }
}
