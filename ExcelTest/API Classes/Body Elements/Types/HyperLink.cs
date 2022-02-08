using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.API_Classes.Body_Elements.Types
{
    class HyperLink :IBodyElement
    {

        [JsonProperty("url")]
        public string Url { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        public HyperLink(string url) { this.Url = url; this.Name = this.Url; }


        public override string ToString()
        {
            return this.Url;
        }

        public string GetElementName()
        {
            return null;
        }
    }
}
