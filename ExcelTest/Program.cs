using ExcelTest.API_Classes;
using ExcelTest.API_Classes.Result;
using ExcelTest.Excel;
using ExcelTest.Exceptions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using ExcelTest.FilesImporter;
using ExcelTest.API_Classes.Body_Elements;

namespace ExcelTest
{
    class Program
    {
        //URL, path to directory, path to excel file and login values should be set in the App.config filee
        private static readonly string serverUrl = ConfigurationManager.AppSettings.Get("serverUrl");
        private static readonly string directoryPath = ConfigurationManager.AppSettings.Get("directoryPath");
        private static readonly string FilePath = ConfigurationManager.AppSettings.Get("FilePath");

        //Create a guid list that will be used for fill data 
        private static List<string> Guids;
        private static Account Account;
        private static FormType FormType;
        private static Workflow Workflow;

        //Api connection to communicate with WEBCON
        private static IApiConnection Connection;

        private static void Main()
        {

            //_ = new Importer(directoryPath, serverUrl);



            Account = new Account(
                ConfigurationManager.AppSettings.Get("clientID"),
                ConfigurationManager.AppSettings.Get("clientSecret"),
                ConfigurationManager.AppSettings.Get("login"));

            FormType = new FormType(ConfigurationManager.AppSettings.Get("formTypeKontrachent"));
            Workflow = new Workflow(ConfigurationManager.AppSettings.Get("workflowKontrachent"));
            string guids = ConfigurationManager.AppSettings.Get("guids").Replace("\r\n", string.Empty).Trim();
            guids = string.Concat(guids.Where(c => !Char.IsWhiteSpace(c)));
            Guids = guids.Split(new char[] { ',' }).ToList();

            //Create connection with API server
            try
            {
                Connection = new ApiConnection(serverUrl, Account);
            }
            catch (Exception)
            {

                Console.WriteLine("Cannot connect to the server");
                Console.ReadKey();
                return;
            }

            //Setup ExcelImporter
            ExcelImporter excel = new(FilePath);

            try
            {

                string columnsNamesToRemoveDuplications = ConfigurationManager.AppSettings.Get("ColumnsNamesToRemoveDuplications").Replace("\r\n", string.Empty).Trim();
                columnsNamesToRemoveDuplications = String.Concat(columnsNamesToRemoveDuplications.Where(c => !Char.IsWhiteSpace(c)));
                List<string> columnsNamesToRemoveDuplicationsList = columnsNamesToRemoveDuplications.Split(new char[] { ',' }).ToList();

                string columnsToRemove = ConfigurationManager.AppSettings.Get("ColumnsToRemove").Replace("\r\n", string.Empty).Trim();
                columnsToRemove = columnsToRemove.Replace("\t", "");
                List<string> columnsToRemoveList = columnsToRemove.Split(new char[] { ',' }).ToList();

                //Load data from excel to list
                var formFieldLists = excel.LoadData(
                    Guids,
                    Convert.ToInt32(ConfigurationManager.AppSettings.Get("sheetID")),
                    columnsNamesToRemoveDuplicationsList,
                    columnsToRemoveList
                    );

                //Create elements to connect with API
                List<object> requestBodyElemetsList = new();

                requestBodyElemetsList.Add(FormType);
                requestBodyElemetsList.Add(Workflow);

                //Sending request to API (singly)
                Console.Write("Ready to send requests!\n0 - Test (Show only JSON, no send)\n1 - Send without show JSON\n2 - Send and show JSON\n");
                int choice = -1;
                int delay = 0;
                try
                {
                    choice = Convert.ToInt32(Console.ReadLine());
                }
                catch (Exception)
                {
                    Console.WriteLine("Bad choice");
                    Console.ReadKey();

                    return;
                }
                try
                {
                    delay = Convert.ToInt32(ConfigurationManager.AppSettings.Get("delay"));
                }
                catch (Exception)
                {
                    Console.WriteLine("Bad delay in config");
                    Console.ReadKey();

                    return;
                }
                if (choice < -1)
                    return;
                Console.WriteLine();

                int i = 1;
                foreach (var item in formFieldLists)
                {
                    Console.WriteLine("Sending " + i);
                    string guid = "84643a1c-8687-4f6b-ac4d-dd27c6e0ee8e";

                    var nipZwykly = item[0] as FormFieldElement<string>;
                    var nipZagraniczny = item[1] as FormFieldElement<string>;
                    if (nipZwykly.Svalue.ToString() != "") item[1] = nipZwykly;

                    
                    requestBodyElemetsList.Add(item);
                    string result = "";

                    //POST new element instance

                    if (choice > 0)
                    {

                        result = Connection.Request("/api/data/v3.0/db/1/elements?path=" + guid + "", requestBodyElemetsList, HttpRequestMethod.POST);

                        PostStartsNewElement postStartsNewElement = JsonConvert.DeserializeObject<PostStartsNewElement>(result);


                        Console.WriteLine(postStartsNewElement.Id + " - " + postStartsNewElement.Status);
                    }

                    
                    if (choice == 0 || choice > 1)
                        result = Connection.RequestTEST("/api/data/v3.0/db/1/elements?path=" + guid + "", requestBodyElemetsList, HttpRequestMethod.POST);

                    if (choice == 0 || choice > 1)
                        ShowJson(result);

                    requestBodyElemetsList.Remove(item);

                    Console.WriteLine("\n" + new string('=', 40) + "\n");

                    i++;

                    _ = Task.Delay(delay);
                }
            }

            catch (IncorrectGuidSizeException e)
            {
                Console.WriteLine(e.Message);
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("Excel file imported\n" + new string('=', 40) + "\n");
        }


        static private void ShowJson(string jsonText)
        {
            JToken parsedJson = JToken.Parse(jsonText);
            var beautified = parsedJson.ToString(Formatting.Indented);
            Console.WriteLine(beautified);
        }
    }
}