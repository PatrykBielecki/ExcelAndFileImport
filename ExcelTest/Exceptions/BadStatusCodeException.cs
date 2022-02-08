using System;

namespace ExcelTest
{
    class BadStatusCodeException : Exception
    {
        public BadStatusCodeException(string message) : base("ERROR:" + message)
        {
        }
    }
}
