using Newtonsoft.Json;

namespace ExcelTest.API_Classes.Body_Elements
{
    class FormFieldElement<T> : IBodyElement
    {

        [JsonProperty("value")]
        public T Value { get; set; }

        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("guid")]
        public string Guid { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("svalue")]
        public string Svalue { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }



        public FormFieldElement(string guid, string type, string name, T value)
        {

            Guid = guid;
            Type = type;
            Value = value;
            Svalue = value.ToString();
            Name = name;
        
        }

        public FormFieldElement(string guid, string type, string name, T value, string sValue)
        {

            this.Guid = guid;
            this.Type = type;
            this.Value = value;
            this.Svalue = sValue;
            this.Name = name;

        }
        public string GetElementName()
        {
            return null;
        }
    }
}
