using Newtonsoft.Json;

namespace ExcelTest.API_Classes.Body_Elements
{
    class Comments : IBodyElement
    {
        [JsonProperty("newComment")]
        public string NewComment { get; set; }

        public Comments(string newComment) { this.NewComment = newComment; }

        public string GetElementName()
        {
            return "comments";
        }
    }
}
