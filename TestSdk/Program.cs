using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using Newtonsoft.Json;
using System.IO;

namespace TestSdk
{
    public class T1Integration
    {
        public int ApplicationNo{get;set;}
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
        public string documentLoc { get; set; }
        public long EddieUri { get; set; }
        public string EddieRecordUrl { get; set; }
    }
    public class T1Property
    {
        public long PropertyNumber { get; set; }
        public string PropertyAddress { get; set; }
        public long EddieUri { get; set; }
        public string EddieRecordUrl { get; set; }
        //public List<T1Animal> lstAnimals { get; set; }
    }
    public class T1Infringement
    {
        public long InfringementNumber { get; set; }
        public string documentLoc { get; set; }
    }
    public class T1Animal
    {
        public long AnimalNumber { get; set; }
        public string documentLoc { get; set; }
    }
    public class T1Request
    {
        public string RequestNumber { get; set; }
        public string documentLoc { get; set; }
    }
    class Program
    {
        public static string pathout = @"C:\temp\T1Watcher\out\";
        public static string pathin = @"C:\temp\T1Watcher\in\";
        static void Main(string[] args)
        {
            FileSystemWatcher watcher = new FileSystemWatcher();

            watcher.Path = pathin;//assigning path to be watched
            watcher.EnableRaisingEvents = true;//make sure watcher will raise event in case of change in folder.
            watcher.IncludeSubdirectories = true;//make sure watcher will look into subfolders as well.
            watcher.Filter = "*.*"; //watcher should monitor all types of file.

            watcher.Created += ProcessT1toRM;
            //watcher.Created += watcher_Created; //register event to be called when a file is created in specified path
            //watcher.Changed += watcher_Changed;//register event to be called when a file is updated in specified path
            //watcher.Deleted += watcher_Deleted;//register event to be called when a file is deleted in specified path

            while (true) ;//run infinite loop so program doesn't terminate untill we force it.
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
                            ProcessApplication(e.FullPath);
                            break;
                        case "t1Propery.txt":
                            ProcessProperty(e.FullPath);
                            break;
                    }
        }
        static void ProcessProperty(string Name)
        {
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                db.Id = "10";
                //db.WorkgroupServerName = "https://SCDB98:9443";
                db.Connect();
                //Console.WriteLine("DB test: " + db.CurrentUser.Name);
                //
                JsonSerializer s = new JsonSerializer();
                try
                {
                    T1Property t1 = null;
                    using (StreamReader sr = new StreamReader(Name))
                    {
                        using (JsonTextReader reader = new JsonTextReader(sr))
                        {
                            //T1Integration t1 = new T1Integration();
                            t1 = (T1Property)s.Deserialize(reader, typeof(T1Property));
                        }
                    }
                    if (t1 != null)
                    {
                        FieldDefinition fdPropertyAddress = new FieldDefinition(db, 502);
                        if (t1.EddieUri == 0)
                        {
                            RecordType rt = new RecordType(db, 1);
                            Record rec = new Record(rt);
                            FieldDefinition fdPropertyNo = new FieldDefinition(db, 504);                           
                            rec.Title = "Property folder - " + t1.PropertyNumber.ToString();
                            rec.SetFieldValue(fdPropertyNo, new UserFieldValue(t1.PropertyNumber));
                            rec.SetFieldValue(fdPropertyAddress, new UserFieldValue(t1.PropertyAddress));
                            Classification c = new Classification(db, 2616);
                            rec.Classification = c;
                            Location loc = new Location(db, 43);
                            rec.SetAssignee(loc);
                            //rec.assig




                            rec.Save();
                            t1.EddieUri = rec.Uri;
                            t1.EddieRecordUrl = rec.WebURL;
                            Console.WriteLine("New property record created ok - number: " + rec.Number);
                        }
                        else
                        {
                            //RecordType rt = new RecordType(db, 2);
                            Record rec = new Record(db, t1.EddieUri);
                            //Record cont = new Record(db, t1.lEddieUri);
                            //rec.Container = cont;
                            //rec.Title = "Application - " + t1.ApplicationNo;
                            //if (t1.propertyLoc != null)
                            //{
                            //    rec.SetDocument(t1.propertyLoc);
                            //}
                            rec.SetFieldValue(fdPropertyAddress, new UserFieldValue(t1.PropertyAddress));
                            rec.Save();
                            //t1.EddieUri = rec.Uri;
                            t1.EddieRecordUrl = rec.WebURL;
                            Console.WriteLine("Property record upgrades - number: " + rec.Number);
                        }
                        using (StreamWriter sw = new StreamWriter(pathout + "T1 property - "+t1.PropertyNumber.ToString()+".txt"))
                        {
                            using (JsonWriter writer = new JsonTextWriter(sw))
                            {
                                s.Serialize(writer, t1);
                                Console.WriteLine("Property file returned");
                            }
                        };
                        FileInfo f = new FileInfo(Name);
                        if (f.Exists)
                            f.Delete();
                    }
                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error: " + exp.Message.ToString());
                    throw;
                }

            }

        }
        static void ProcessApplication(string Name)
        {
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                db.Id = "10";
                db.Connect();
                JsonSerializer s = new JsonSerializer();
                try
                {
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
                        FieldDefinition fdPropertyAddress = new FieldDefinition(db, 502);
                        if (t1.EddieUri == 0)
                        {
                            RecordType rt = new RecordType(db, 2);
                            Record rec = new Record(rt);
                            if (t1.PropertyNumber != 0)
                            {
                                Record cont = new Record(db, t1.PropertyNumber);
                                rec.Container = cont;
                            }
                            //FieldDefinition fdPropertyNo = new FieldDefinition(db, 504);
                            rec.Title = "Application - " + t1.ApplicationNumber;
                            //rec.SetFieldValue(fdPropertyNo, new UserFieldValue(t1.PropertyNumber));
                            //rec.SetFieldValue(fdPropertyAddress, new UserFieldValue(t1.PropertyAddress));
                            //Classification c = new Classification(db, 2616);
                            //rec.Classification = c;
                            FileInfo f = new FileInfo(t1.documentLoc);
                            if (f.Exists)
                            {
                                rec.SetDocument(new InputDocument(t1.documentLoc), false, false, "");
                            }
                            rec.Save();
                            t1.EddieUri = rec.Uri;
                            t1.EddieRecordUrl = rec.WebURL;
                            Console.WriteLine("New application record created ok - number: " + rec.Number);
                        }
                        else
                        {
                            //RecordType rt = new RecordType(db, 2);
                            Record rec = new Record(db, t1.EddieUri);
                            //Record cont = new Record(db, t1.lEddieUri);
                            //rec.Container = cont;
                            //rec.Title = "Application - " + t1.ApplicationNo;
                            //if (t1.propertyLoc != null)
                            //{
                            //    rec.SetDocument(t1.propertyLoc);
                            //}
                            //rec.SetFieldValue(fdPropertyAddress, new UserFieldValue(t1.PropertyAddress));
                            FileInfo f1 = new FileInfo(t1.documentLoc);
                            if (f1.Exists)
                            {
                                rec.SetDocument(new InputDocument(t1.documentLoc), true, false, "");
                            }
                            rec.Save();
                            //t1.EddieUri = rec.Uri;
                            t1.EddieRecordUrl = rec.WebURL;
                            Console.WriteLine("Application record upgrades - number: " + rec.Number);
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
                catch (Exception exp)
                {
                    Console.WriteLine("Error: " + exp.Message.ToString());
                    throw;
                }

            }

        }
        static void SendBacktoT1(T1Integration t)
        {
            JsonSerializer s = new JsonSerializer();
            using (StreamWriter sw = new StreamWriter(pathout+t.PropertyNo.ToString()))
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
