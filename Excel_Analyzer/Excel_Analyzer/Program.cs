using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Analyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelApplication IVS = new ExcelApplication();

            IVS.ExcelOpen();
        }
    }
}
