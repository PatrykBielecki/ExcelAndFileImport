using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.Exceptions
{
    class IncorrectGuidSizeException : Exception
    {
        public IncorrectGuidSizeException(string message) : base("ERROR: " + message) { }
    }
}
