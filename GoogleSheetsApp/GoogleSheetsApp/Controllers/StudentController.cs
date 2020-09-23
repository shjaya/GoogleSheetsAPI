using GoogleSheetsApp.Models;
using System.Collections.Generic;
using System.Web.Mvc;
using Google.Apis.Auth.OAuth2;
using System;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace GoogleSheetsApp.Controllers
{
    public class StudentController : Controller
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        public static string sheetId = null;
        // GET: Student
        public ActionResult Index()
        {                                                  
            var studentList = GetStudents();
            return View(studentList);
        }
        public ActionResult OpenGoogleSheets()
        {
            var service = AuthorizeGoogleAppForSheetsService();

            string newRange = GetRange(service, sheetId);
            IList<IList<Object>> objNeRecords = GenerateData();
            UpdatGoogleSheetinBatch(objNeRecords, sheetId, newRange, service);

            string sheetsURL = "https://docs.google.com/spreadsheets/d/" + sheetId + "/edit#gid=0";
            return Redirect(sheetsURL);
        }

        public List<Student> GetStudents()
        {
            List<Student> students = new List<Student>();
            for (int i = 0; i <= 6; i++)
            {
                Student student = new Student();
                student.Id = i;
                student.Name = "Student" + i;
                student.Roll = "Roll" + i;

                students.Add(student);
            }

            return students;
        }

        private static SheetsService AuthorizeGoogleAppForSheetsService()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //Creating the sheet now in Google Sheets after being authorised.
            Guid obj = Guid.NewGuid();
            string sheetName = string.Format("{0} - {1}", "Google Sheets Testing", obj.ToString());
            var myNewSheet = new Google.Apis.Sheets.v4.Data.Spreadsheet();
            myNewSheet.Properties = new SpreadsheetProperties();
            myNewSheet.Properties.Title = sheetName;
            var newSheet = service.Spreadsheets.Create(myNewSheet).Execute();
            sheetId = newSheet.SpreadsheetId;

            return service;
        }        
        protected static string GetRange(SheetsService service, string SheetId)
        {
            // Define request parameters.  
            String spreadsheetId = SheetId;
            String range = "A:A";

            SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                       service.Spreadsheets.Values.Get(spreadsheetId, range);
            System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
            ValueRange getResponse = getRequest.Execute();
            IList<IList<Object>> getValues = getResponse.Values;
            if (getValues == null)
            {
                // spreadsheet is empty return Row A Column A  
                return "A:A";
            }

            int currentCount = getValues.Count() + 1;
            String newRange = "A" + currentCount + ":A";
            return newRange;
        }
        private static IList<IList<Object>> GenerateData()
        {
            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            int maxrows = 7;  
            for (var i = 0; i <= maxrows; i++)
            {
                IList<Object> obj = new List<Object>();
                if (i==0)
                {
                    obj.Add("ID");
                    obj.Add("Name");
                    obj.Add("Roll");
                }
                else if (i > 0)
                {
                    obj.Add(i - 1);
                    obj.Add("Student" + (i - 1));
                    obj.Add("Roll" + (i - 1));
                }
                objNewRecords.Add(obj);
            }
            return objNewRecords;
        }
        private static void UpdatGoogleSheetinBatch(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            SpreadsheetsResource.ValuesResource.AppendRequest request =
               service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }
    }
}
