using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using Newtonsoft.Json;
using System.IO;

namespace SearchTitle
{
    public class Sheets
    {
        public string SheetName { get; set; }
        public List<Rows> Rows { get; set; }
    }
    public class Rows
    {
        public long SheetLine { get; set; }
        public string RMuri { get; set; }
        public string RMString { get; set; }
        public string MigrateLoc { get; set; }
        public int OriginalFileCount { get; set; }
        public int RMFileCount { get; set; }
        public DateTime Timestamp { get; set; }
        public string Error { get; set; }
        public string level1 { get; set; }
        public string level2 { get; set; }
        public string level3 { get; set; }
        public string Folder { get; set; }
        public string Subfolder { get; set; }
    }
    public class RM
    {
        public long Uri { get; set; }
        public string title { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            if (!TrimApplication.HasBeenInitialized)
            {
                TrimApplication.Initialize();
            }
            TrimApplication.HasUserInterface = true;
            //
            List<Sheets> myMessage = JsonConvert.DeserializeObject<List<Sheets>>(File.ReadAllText(@"C:\temp\SSprocess.json"));
            List<RM> lstRm = new List<RM>();
            using (Database db = new Database())
            {
                db.Connect();
                //10030-Westview Cres NAMBOUR
                //var p = "10030-Westview Cres NAMBOURR";
                Record rec = null;
                TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                objS.SetSearchString("classification:7726");
                //if (objS.FastCount == 0)
                //Console.WriteLine("No results for this folder under the class: " + cl3.Title + "/" + p);

                foreach (Record r in objS)
                {
                    RM rm = new RM();
                    rm.Uri = r.Uri;
                    rm.title = r.Title;
                    lstRm.Add(rm);
                    //Console.WriteLine("Loop through folders under class:" + p + "/" + r.Title);
                    //if (r.Title == p.Trim())
                    //{
                    //    rec = r;
                    //}
                }
                //if (rec != null)
                //{
                //    Console.WriteLine("Found record:" + rec.Uri);
                //}
                //else
                //{
                //    Console.WriteLine("Record is null");
                //}
            }
            Console.WriteLine("Record count: " + lstRm.Count().ToString());
            Console.ReadLine();

            foreach (var file in myMessage)
            {
                //file.Rows.Skip(9880).Take(30)
                //file.Rows.Where(x => x.RMuri == "0" && x.Error == null).Take(30)
                //.Where(x=>x.Error==null && x.RMuri=="0")
                Console.WriteLine("SS row count: " + file.Rows.Count().ToString());
                foreach (Rows r in file.Rows.Where(x => x.Folder !=null && x.RMuri == "0").Take(30))
                {
                    //.Trim()
                    var d = r.Folder.Trim();
                    var h = "Not found";
                    if(lstRm.Exists(x=>x.title.Contains(d)))
                    {
                        h = lstRm.Find(x=>x.title==d).title;
                    }

                    Console.WriteLine("Process test: " + r.Folder + " / "+ h);
                    // (x => x.title == d).FirstOrDefault().title);
                }
            }
            Console.ReadLine();
            
        }
    }
}