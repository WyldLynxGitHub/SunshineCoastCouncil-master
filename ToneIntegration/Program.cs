using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using HP.HPTRIM.SDK;
using Newtonsoft.Json;
using System.IO;
using HP.HPTRIM.Service.Client;
using HP.HPTRIM.ServiceModel;
using System.Configuration;

namespace TestSdk
{
    public class T1Integration
    {
        public int ApplicationNo { get; set; }
        public int PropertyNo { get; set; }
        public int AnimalNo { get; set; }
        public int InfringementNo { get; set; }
        public string PropertyAddress { get; set; }
        public int RequestNumber { get; set; }
        public long lEddieUri { get; set; }
        public string propertyLoc { get; set; }
        public string EddieRecordUrl { get; set; }
    }
    public class T1Application
    {
        public string ApplicationNumber { get; set; }
        public long PropertyNumber { get; set; }
        public string PropertyAddress { get; set; } //moved from the T1Property text file
        public string documentLoc { get; set; }
        public long EddieUri { get; set; } //Application document URI
        public string EddieRecordNumber { get; set; }
        public string EddieRecordUrl { get; set; } //Application document URL
        public int DocumentID { get; set; } //ID for T1 return management
        public string ApplicationTitle { get; set; } //Set Eddie title property
        public string PropertyTitle { get; set; } //Set Eddie title property
        public string FolderStructure { get; set; }
    }
    public class T1Config
    {
        public string pathout { get; set; }
        public string pathin { get; set; }
        public string serviceapi { get; set; }

    }
    class Program
    {
        public static string pathout;
        public static string pathin;
        public static string serviceapi;
        static void Main(string[] args)
        {
            JsonSerializer s = new JsonSerializer();
            T1Config config = null;
            using (StreamReader sr = new StreamReader("t1Config.txt"))
            {
                using (JsonTextReader reader = new JsonTextReader(sr))
                {
                    config = (T1Config)s.Deserialize(reader, typeof(T1Config));
                }
            }
            if (config != null)
            {
                pathout = config.pathout;
                pathin = config.pathin;
                serviceapi = config.serviceapi;
            }
            else
            {
                Console.WriteLine("T1 config file error.");
                return;
            }

                //t1Config.txt
                //
                FileSystemWatcher watcher = new FileSystemWatcher();

            watcher.Path = pathin;//assigning path to be watched
            watcher.EnableRaisingEvents = true;//make sure watcher will raise event in case of change in folder.
            watcher.IncludeSubdirectories = true;//make sure watcher will look into subfolders as well.
            watcher.Filter = "*.*"; //watcher should monitor all types of file.

            watcher.Created += ProcessT1toRM;
            Console.WriteLine("T1 property integration simulater loaded and listening");
            while (true) ;
        }
        static void watcher_Created(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("File : " + e.FullPath + " is created.");

        }
        static void ProcessT1toRM(object sender, FileSystemEventArgs e)
        {
            //Console.WriteLine("Test");

            var inputfile = e.Name;
            switch (inputfile)
            {
                case "t1Application.txt":
                    //ProcessApplication(e.FullPath);
                    ApplicationService(e.FullPath);
                    break;
                //case "t1Propery.txt":
                //    ProcessProperty(e.FullPath);
                //    break;
            }
        }
        static void ApplicationService(string Name)
        {

            try
            {
                TrimClient trimClient = new TrimClient(serviceapi);
                trimClient.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //
                Locations request = new Locations();
                request.q = "me";
                request.Properties = new List<string>() { "SortName" };

                var Loc = trimClient.Get<LocationsResponse>(request);

                Console.WriteLine("Logged onto Eddie as " + Loc.Results[0].SortName);

                JsonSerializer s = new JsonSerializer();

                T1Application t1 = null;
                using (StreamReader sr = new StreamReader(Name))
                {
                    using (JsonTextReader reader = new JsonTextReader(sr))
                    {
                        //T1Integration t1 = new T1Integration();
                        t1 = (T1Application)s.Deserialize(reader, typeof(T1Application));
                    }
                }
                if (t1 != null)
                {
                    Console.WriteLine("T1 export file processed.");
                    if (t1.EddieUri == 0)
                    {
                        Record record = new Record();
                        
                        record.RecordType = new RecordTypeRef() { Uri = 1 };
                        record.Title = t1.PropertyTitle + " - " + DateTime.Now.ToString();
                        record.Classification = new ClassificationRef() { Uri=2598 };
                        record.SetCustomField("PropertyNumber", t1.PropertyNumber);
                        record.SetCustomField("ApplicationNumber", t1.ApplicationNumber);
                        record.SetCustomField("PropertyAddress", t1.PropertyAddress);
                        RecordsResponse response = trimClient.Post<RecordsResponse>(record);

                        if(response.Results.Count==1)
                        {
                            Console.WriteLine("New property container created in Eddie - RM Uri: " + response.Results.First().Uri.ToString());
                        }
                        //
                        FileInfo f = new FileInfo(t1.documentLoc);
                        if (f.Exists)
                        {
                            Record doc = new Record();
                            doc.Properties = new List<string>() { "Number" , "Url"};
                            doc.RecordType = new RecordTypeRef() { Uri = 2 };
                            doc.Container = new RecordRef() { Uri = response.Results.First().Uri };
                            doc.Title= t1.ApplicationTitle + " - " + DateTime.Now.ToString();
                            doc.SetCustomField("PropertyNumber", t1.PropertyNumber);
                            doc.SetCustomField("ApplicationNumber", t1.ApplicationNumber);

                            RecordsResponse response11= trimClient.PostFileWithRequest<RecordsResponse>(f, doc);
                            if(response11.Results.Count==1)
                            {
                                Record rr = response11.Results.First();
                                Console.WriteLine("New application document created in Eddie - RM Uri: " + rr.Uri.ToString());
                                t1.EddieUri = rr.Uri;
                                t1.EddieRecordNumber = rr.Number;
                                t1.EddieRecordUrl = rr.Url;
                            }
                        }
                    }
                    else
                    {
                        //Add amend function
                    }
                    using (StreamWriter sw = new StreamWriter(pathout + "T1 Application - " + t1.PropertyNumber.ToString() + ".txt"))
                    {
                        using (JsonWriter writer = new JsonTextWriter(sw))
                        {
                            s.Serialize(writer, t1);
                            Console.WriteLine("Application file returned");
                        }
                    };
                    FileInfo f2 = new FileInfo(Name);
                    if (f2.Exists)
                        f2.Delete();
                }
            }
            //catch (ServiceStack.ServiceClient.Web.WebServiceException ex)
            //{
            //    Console.WriteLine("Service stack error: "+ex.ErrorMessage);
            //}
            catch (Exception exp)
            {
                Console.WriteLine("General error: " + exp.Message.ToString());
                throw;
            }

        }
        static void SendBacktoT1(T1Integration t)
        {
            JsonSerializer s = new JsonSerializer();
            using (StreamWriter sw = new StreamWriter(pathout + t.PropertyNo.ToString()))
            {
                using (JsonWriter writer = new JsonTextWriter(sw))
                {
                    s.Serialize(writer, t);
                }
            }
            //T1Integration t = new T1Integration();
            //t.ApplicationNo = 123;
            //t.PropertyNo = 544;
            //t.InfringementNo = 656565;
            //t.PropertyAddress = "45 Somewhere st, suburb";
            //t.RequestNumber = 7676;
            //t.AnimalNo = 7676;
            //
            //string output = JsonConvert.SerializeObject(t);
            //using(StreamWriter sw = new StreamWriter(@"c:\TonePropertyTest.txt"))
            //{
            //    using(JsonWriter writer = new JsonTextWriter(sw))
            //    {
            //        s.Serialize(writer, t);
            //    }
            //}
        }
    }
}
