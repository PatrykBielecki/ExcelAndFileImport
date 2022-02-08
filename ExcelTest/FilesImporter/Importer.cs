using ExcelTest.API_Classes.Body_Elements.Types;
using ExcelTest.API_Classes.Body_Elements;
using ExcelTest.API_Classes.Result;
using System.Collections.Generic;
using ExcelTest.API_Classes;
using System.Configuration;
using Newtonsoft.Json;
using System.Linq;
using System.IO;
using System;

namespace ExcelTest.FilesImporter
{
    class Importer
    {
        private readonly string directory;
        private static IApiConnection connection;
        private readonly string[] notAllowedFileType = ConfigurationManager.AppSettings["notAllowedFileType"].Split(',');

        public Importer(string directoryPath, string url)
        {
            directory = directoryPath;
            Console.WriteLine("Reading files from " + directoryPath);

            try
            {
                connection = new ApiConnection(url, new Account());
            }
            catch (Exception)
            {

                Console.WriteLine("Cannot connect to the server");
                return;
            }

            Console.WriteLine("Import started, connected to " + url + "\n");

            DirSearch(directory);

            Console.WriteLine("\nFiles imported\n" + new string('=', 40) + "\n");
        }


        private void DirSearch(string sDir, long id = 0)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    if (Directory.GetParent(Directory.GetParent(d).ToString()).ToString() == directory)
                    {
                        id = PostClient(Path.GetFileName(Directory.GetParent(d).ToString())[11..], Path.GetFileName(d));
                    }

                    foreach (string f in Directory.GetFiles(d))
                    {

                        if (notAllowedFileType.Contains(Path.GetExtension(f)))
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("File " + f + " is not allowed type, not imported");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        else PostAttachment(f, id);
                    }

                    DirSearch(d, id);
                }
            }

            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }



        private static void PostAttachment(string path, long id)
        {
            List<object> requestBodyElemetsList2 = new();
            Attachments attachments = new();

            string fileContent = Convert.ToBase64String(File.ReadAllBytes(path));
            string fileName = Path.GetFileName(path);
            string description = Path.GetFileName(Directory.GetParent(path).ToString());

            attachments.content = fileContent;
            attachments.name = fileName;
            attachments.description = description;

            requestBodyElemetsList2.Add(attachments);

            if (!string.IsNullOrEmpty(fileContent))
            {
                Console.WriteLine("Sending document " + fileName + ", document type " + description); //put into app.config
                connection.Request("/api/data/v3.0/db/1/elements/" + id + "/attachments?path=" + "0b7be5d3-be51-4fec-a4c6-d0f1411f04da" + "", requestBodyElemetsList2, HttpRequestMethod.POST);
            }
        }


        private static long PostClient(string name, string catName)
        {
            List<object> requestBodyElemetsList = new();

            FormType formType = new(ConfigurationManager.AppSettings.Get("formTypeUmowa"));
            Workflow workflow = new(ConfigurationManager.AppSettings.Get("workflowUmowa"));
            FormFieldList formFieldList = new();


            formFieldList.Add(new FormFieldElement<string>(ConfigurationManager.AppSettings.Get("nazwaKontrachenta"), "SingleLine", "", name));
            formFieldList.Add(new FormFieldElement<string>(ConfigurationManager.AppSettings.Get("nazwaKategorii"), "SingleLine", "", catName));
            formFieldList.Add(new FormFieldElement<string>(ConfigurationManager.AppSettings.Get("nazwaPodfolderu"), "SingleLine", "", catName));
            formFieldList.Add(new FormFieldElement<string>(ConfigurationManager.AppSettings.Get("checkboxImport"), "Boolean", "", "1"));



            requestBodyElemetsList.Add(formType);
            requestBodyElemetsList.Add(workflow);
            requestBodyElemetsList.Add(formFieldList);

            string guid = "0b7be5d3-be51-4fec-a4c6-d0f1411f04da";

            Console.WriteLine("\nCreating client: " + name);
            string result = connection.Request("/api/data/v3.0/db/1/elements?path=" + guid + "", requestBodyElemetsList, HttpRequestMethod.POST);
            Console.WriteLine("\nCreated client: " + name);
            PostStartsNewElement postStartsNewElement = JsonConvert.DeserializeObject<PostStartsNewElement>(result);

            return postStartsNewElement.Id;
        }
    }
}
