#define CONSOLE
#define EXCEL
#define FILE

using AQUA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALS.AQUA
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelOps.InitializeWorkbook();
        }
    }
}
