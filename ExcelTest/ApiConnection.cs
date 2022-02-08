using ExcelTest.API_Classes;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace ExcelTest
{
    public enum HttpRequestMethod { POST, PATCH }
    class ApiConnection : IApiConnection
    {
        AuthorizationToken AuthorizationToken { get; set; }
        readonly Uri apiUri;
        readonly HttpClient HttpClient;

        public ApiConnection(string URL, Account account)
        {
            apiUri = new Uri(URL);
            HttpClient = CreateHttpClient();
            GetAuthorizationToken(account);
        }

        /// <summary>
        /// Send POST and PATCH (only theoretical) request to API
        /// </summary>
        /// <param name="parameters">The rest of the url address</param>
        /// <param name="elements">List of elements to send in body</param>
        /// <returns>String received in response</returns>
        public string Request(string parameters, List<Object> elements, HttpRequestMethod httpRequestMethod = HttpRequestMethod.POST)
        {
            HttpResponseMessage response = null;

            string account = "{";

            foreach (var item in elements)
            {
                string elementName = ((IBodyElement)item).GetElementName();
                if (elementName != null)
                    account += $"\"{elementName}\": {JsonConvert.SerializeObject(item)},";
                else
                    account += JsonConvert.SerializeObject(item) + ",";

            }
            account = account[0..^1];
            account += "}";
            if (elements.Count == 1)
                account = account.Substring(1, account.Length - 1);

            var httpContent = new StringContent(account, Encoding.UTF8, "application/json");
            switch (httpRequestMethod)
            {
                case HttpRequestMethod.POST:
                    response = HttpClient.PostAsync(parameters, httpContent).Result;
                    break;
                case HttpRequestMethod.PATCH:
                    response = HttpClient.PatchAsync(parameters, httpContent).Result;
                    break;
            }


            try
            {
                return CheckResponce(response);
            }
            catch (BadStatusCodeException e)
            {

                throw new BadStatusCodeException(e.Message);
            }

        }

        public string RequestTEST(string parameters, List<Object> elements, HttpRequestMethod httpRequestMethod = HttpRequestMethod.POST)
        {

            string account = "{";

            foreach (var item in elements)
            {
                string elementName = ((IBodyElement)item).GetElementName();
                if (elementName != null)
                    account += $"\"{elementName}\": {JsonConvert.SerializeObject(item)},";
                else
                    account += JsonConvert.SerializeObject(item) + ",";

            }
            account = account.Substring(0, account.Length - 1);
            account += "}";
            if (elements.Count == 1)
                account = account.Substring(1, account.Length - 1);
            _ = new StringContent(account, Encoding.UTF8, "application/json");
            switch (httpRequestMethod)
            {
                case HttpRequestMethod.POST:
                    //response = this.HttpClient.PostAsync(parameters, httpContent).Result;
                    break;
                case HttpRequestMethod.PATCH:
                    //response = this.HttpClient.PatchAsync(parameters, httpContent).Result;
                    break;
            }

            return account;

        }

        /// <summary>
        /// Send GET request to API
        /// </summary>
        /// <param name="parameters">The rest of the url address</param>
        /// <returns>String received in response</returns>
        public string Request(string parameters)
        {
            HttpResponseMessage response = HttpClient.GetAsync(parameters).Result;
            try
            {
                return CheckResponce(response);
            }
            catch (BadStatusCodeException e)
            {

                throw new BadStatusCodeException(e.Message);
            }
        }

        /// <summary>
        /// Download Autorization Token
        /// </summary>
        /// <returns>Token</returns>
        void GetAuthorizationToken(Account account)
        {
            if (this.AuthorizationToken != null)
            {
                Console.WriteLine("Autorization token exist");
                return;
            }
            try
            {
                string data = Request("/api/login", new List<object>() { account });

                AuthorizationToken = System.Text.Json.JsonSerializer.Deserialize<AuthorizationToken>(data);
                HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", this.AuthorizationToken.token);
                return;
            }
            catch (BadStatusCodeException e)
            {
                Console.WriteLine("Autorization Error");
                throw new BadStatusCodeException(e.Message);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

        }

        /// <summary>
        /// Create http client to API with/without token 
        /// </summary>
        HttpClient CreateHttpClient()
        {
            HttpClient client = new();
            client.BaseAddress = apiUri;
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            return client;
        }

        /// <summary>
        /// Checking the response status 
        /// </summary>
        static string CheckResponce(HttpResponseMessage response)
        {
            if (response.IsSuccessStatusCode)
            {
                var dataObjects = response.Content.ReadAsStringAsync().Result;
                return dataObjects;
            }
            else
            {
                throw new BadStatusCodeException((int)response.StatusCode + " " + response.ReasonPhrase + "  -  " + response.Content.ReadAsStringAsync().Result);
            }
        }

        ~ApiConnection()
        {
            HttpClient.Dispose();
        }
    }
}
