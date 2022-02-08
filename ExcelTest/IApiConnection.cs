using System.Collections.Generic;

namespace ExcelTest
{
    interface IApiConnection
    {
        /// <summary>
        /// Send GET request to API
        /// </summary>
        /// <param name="parameters">The rest of the url address</param>
        /// <returns>String received in response</returns>
        string Request(string parameters);

        /// <summary>
        /// Send POST request to API
        /// </summary>
        /// <param name="parameters">The rest of the url address</param>
        /// <param name="elements">List of elements to send in body</param>
        /// <returns>String received in response</returns>
        public string Request(string parameters, List<object> elements, HttpRequestMethod httpRequestMethod = HttpRequestMethod.POST);
        public string RequestTEST(string parameters, List<object> elements, HttpRequestMethod httpRequestMethod = HttpRequestMethod.POST);
    }
}
