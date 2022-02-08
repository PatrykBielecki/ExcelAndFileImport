using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.API_Classes.Body_Elements.Types
{
    class SurveyChooseField : IBodyElement
    {

        [JsonProperty("choices")]
        public PickerValue Choices { get; set; }

        [JsonProperty("other")]
        public string Other { get; set; }



        public SurveyChooseField(string id, string name) { this.Choices = new PickerValue(id, name); this.Other = ""; }

        public string GetElementName()
        {
            return null;
        }
    }
}
