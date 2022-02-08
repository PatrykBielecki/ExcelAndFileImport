using Newtonsoft.Json;

namespace ExcelTest.API_Classes
{
    class Company : IBodyElement
    {
        [JsonProperty("guid")]
        public string Guid { get; set; }

        public Company(string guid) { this.Guid = guid; }

        public string GetElementName()
        {
            return "company";
        }
    }
}
