using Newtonsoft.Json;

namespace ExcelTest.API_Classes
{
    class Workflow : IBodyElement
    {
        [JsonProperty("guid")]
        public string Guid { get; set; }

        public Workflow(string guid) { this.Guid = guid; }

        public string GetElementName()
        {
            return "workflow";
        }
    }
}
