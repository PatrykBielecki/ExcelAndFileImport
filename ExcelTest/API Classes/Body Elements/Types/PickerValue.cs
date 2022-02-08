using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.API_Classes.Body_Elements.Types
{
    class PickerValue : List<object>, IBodyElement
    {
        //public List<PickerValueElement> list = new List<PickerValueElement>();
        
        public PickerValue(string id, string name)
        {
            Add(new PickerValueElement(id, name));
        }
        public string GetElementName()
        {
            return "choices";
        }
    }
}
