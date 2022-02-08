using ExcelTest.API_Classes;
using Newtonsoft.Json;
using System.Configuration;

namespace ExcelTest
{
    class Account : IBodyElement
    {
        [JsonProperty("ClientID")]
        public string ClientID  {get; set;}
        [JsonProperty("ClientSecret")]
        public string ClientSecret { get; set; }
        [JsonProperty("Impersonation")]
        public ImpresonationClass Impersonation { get; set; }

        public Account()
        {

            ClientID = ConfigurationManager.AppSettings.Get("ClientID");
            ClientSecret = ConfigurationManager.AppSettings.Get("ClientSecret");
            Impersonation = new ImpresonationClass();
        }


        public Account(string clientID, string clientSecret, string login)
        {

            ClientID = clientID;
            ClientSecret = clientSecret;
            Impersonation = new ImpresonationClass(login);
        }

        public string GetElementName()
        {
            return null;
        }
    }

    internal class ImpresonationClass
    {
        [JsonProperty("Login")]
        public string Login
        { get; set; }

        public ImpresonationClass(string Login)
        {
            this.Login = Login;
        }

        public ImpresonationClass()
        {
            Login = ConfigurationManager.AppSettings.Get("login");
        }

    }
}
