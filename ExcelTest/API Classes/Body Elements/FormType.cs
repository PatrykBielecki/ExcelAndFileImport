using Newtonsoft.Json;

namespace ExcelTest.API_Classes
{
    class FormType : IBodyElement
    {
        [JsonProperty("guid")]
        public string Guid { get; set; }

        public FormType(string guid) { this.Guid = guid; }

        public string GetElementName()
        {
            return "formType";
        }
    }
}
