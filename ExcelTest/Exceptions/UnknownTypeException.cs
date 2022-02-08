using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.Exceptions
{
    class UnknownTypeException : Exception
    {

        public UnknownTypeException(string message) : base("ERROR: " + message) { }

    }
}
