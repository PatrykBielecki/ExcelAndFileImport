using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.API_Classes.Body_Elements.Types
{
    class PickerValueElement : IBodyElement
    {

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        public PickerValueElement(string choice, string name) { Id = choice; Name = name; }
        public string GetElementName()
        {
            return null;
        }
    }
}
