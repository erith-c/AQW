#define CONSOLE
#define EXCEL
#define FILE

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALS
{
    public class Error
    {
        public class MissingDirectory : Exception
        {
            public MissingDirectory()
            {

            }
            public MissingDirectory(string message)
                : base(message)
            {

            }
            public MissingDirectory(string message, Exception inner)
                : base(message, inner)
            {

            }
        }
    }
}
