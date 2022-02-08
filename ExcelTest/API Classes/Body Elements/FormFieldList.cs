using System.Collections.Generic;

namespace ExcelTest.API_Classes.Body_Elements
{
    class FormFieldList : List<object>, IBodyElement
    {
        public string GetElementName()
        {
            return "formFields";
        }
    }
}
