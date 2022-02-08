using Newtonsoft.Json;

namespace ExcelTest.API_Classes.Result
{
    class PostStartsNewElement
    {
        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("instanceNumber")]
        public string InstanceNumber { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }
    }
}
