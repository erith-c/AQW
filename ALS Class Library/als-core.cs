using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ALS
{
    namespace APC
    {
        public class DataRecord
        {
            // Properties
            public string Name { get; }
            public string Source { get; set; }
            public string CellRef { get; set; }
            public string Value { get; set; }
            // Methods

            // Constructors
            public DataRecord()
            {
                Name = "";
                Source = "";
                CellRef = "";
                Value = "";
            }
        }
    }
}
