using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.API_Classes.Body_Elements.Types
{
    class Date
    {
        [JsonProperty("date")]
        public DateTime DateValue { get; set; }

        public Date(string date) { this.DateValue = DateTime.Parse(date).Date; }

        public override string ToString()
        {
            return this.DateValue.ToString();
        }
    }
}
