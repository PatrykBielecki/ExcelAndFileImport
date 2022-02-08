namespace ExcelTest.API_Classes.Body_Elements
{
    class Attachments : IBodyElement
    {
        public string content { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string group { get; set; }

        public string GetElementName()
        {
            return null;
        }
    }
}
