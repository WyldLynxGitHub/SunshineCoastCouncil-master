using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace ReadJsonTest
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
    class Program
    {
        private static long toplevelUri = 0;
        private static string toplevel = null;
        private static string json = null;
        static void Main(string[] args)
        {
            long recuri = 0;
            long folderUri = 0;
            long classUri = 0;
            List<Sheets> myMessage = JsonConvert.DeserializeObject<List<Sheets>>(File.ReadAllText(@"C:\temp\SSprocess.json"));
            if (!TrimApplication.HasBeenInitialized)
            {
                TrimApplication.Initialize();
            }
            TrimApplication.HasUserInterface = true;
            //
            using (Database db = new Database())
            {
                db.Connect();
                toplevelUri = Convert.ToInt64(Console.ReadLine());
                if (getToplevel(toplevelUri, db) != null)
                {
                    toplevel = getToplevel(toplevelUri, db).Title;
                    Console.WriteLine("Confirm that this is the folder stating point for this process: " + toplevel + " Hit q to quit or any other key to continue.");
                    if (Console.ReadLine() == "q")
                    {
                        Environment.Exit(0);
                    }
                }
                else
                {
                    Console.WriteLine("Please check the high level classification Uri and try again:");
                    Console.ReadLine();
                    Environment.Exit(0);
                }
                foreach (var file in myMessage)
                {
                    //file.Rows.Skip(9880).Take(30)
                    //file.Rows.Where(x => x.RMuri == "0" && x.Error == null).Take(30)
                    //.Where(x=>x.Error==null && x.RMuri=="0")
                    foreach (Rows r in file.Rows.Where(x => x.SheetLine > 2 && x.RMuri == "0").Take(30))
                    {
                        Console.WriteLine(r.SheetLine.ToString() + " " + r.RMuri + " (" + r.level1 + " - " + r.level2 + " - " + r.level3 + " - " + r.Folder + " - " + r.Subfolder + ") " + r.Error);
                        //
                        Classification cl1 = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, toplevel + " - " + r.level1);
                        if (cl1 == null)
                        {
                            //cl1 = CreateFolderStructure(r.level1, toplevel, "c", db, FindStartingNumber(toplevelUri, db));
                            //classUri = cl1.Uri;
                        }

                        Classification cl2 = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, cl1.Title + " - " + r.level2);
                        if (cl2 == null)
                        {
                            //cl2 = CreateFolderStructure(r.level2, cl1.Title, "c", db, FindStartingNumber(cl1.Uri, db));
                            //classUri = cl2.Uri;
                        }

                        Classification cl3 = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, cl2.Title + " - " + r.level3);
                        if (cl3 == null)
                        {
                            //cl3 = CreateFolderStructure(r.level3, cl2.Title, "r", db, FindStartingNumber(cl2.Uri, db));
                            //classUri = cl3.Uri;
                        }
                        //Classification cl3 = (Classification)db.FindTrimObjectByUri(BaseObjectTypes.Classification, cl3.Uri);
                        if (r.Folder != null)
                        {
                            Record chkRec = CheckifFolderAlreadyExists(cl3, r.Folder, db);
                            if (chkRec == null)
                            {
                                Console.WriteLine("Did not find existing folder: " + cl3.Title + " " + r.Folder);
                                try
                                {
                                    recuri = CreateNewRMfolder(r.Folder, cl3.Uri, db);
                                }
                                catch (Exception exp)
                                {
                                    Console.WriteLine("Error creating folder: " + exp.Message.ToString() + " - " + r.Folder + " (" + cl3.Title + ")");
                                    r.Error = exp.Message.ToString();
                                }
                                folderUri = recuri;
                            }
                            else
                            {
                                Console.WriteLine("Found existing folder: " + chkRec.Number);
                                recuri = chkRec.Uri;
                                //r.RMuri=recuri.ToString();
                                folderUri = recuri;
                            }
                        }

                        if (r.Subfolder != null && folderUri != 0)
                        {
                            Record chkRecSub = CheckSubFolderAlreadyExists(folderUri, r.Subfolder, db);
                            if (chkRecSub == null)
                            {
                                Console.WriteLine("Did not find existing sub folder: " + r.Subfolder);
                                recuri = CreateNewRMSubFolder(folderUri, r.Subfolder, db);
                            }
                            else
                            {
                                Console.WriteLine("Found existing sub folder: " + chkRecSub.Number);
                                recuri = chkRecSub.Uri;
                                //r.RMuri=
                            }
                        }
                        //if (r.Folder != null)
                        //{
                        //    recuri = CreateNewRMfolder(r.Folder, cl3.Uri, db);
                        //   folderUri = recuri;

                        //}
                        //if (r.Subfolder != null)
                        //{
                        //    recuri = CreateNewRMSubFolder(folderUri, r.Subfolder, db);
                        //}
                        //if (r.Folder == null && r.Subfolder == null)
                        //{
                        //    recuri = classUri;
                        //}

                        //r.RMuri = recuri.ToString();
                        ////r.Error = "All good";
                        //Console.WriteLine("Processed SS line " + r.SheetLine.ToString() + " with Uri " + recuri.ToString());

                    }
                    Console.ReadLine();
                }
            }
        }
        private static Classification getToplevel(long toplevelUri, Database db)
        {
            Classification c = db.FindTrimObjectByUri(BaseObjectTypes.Classification, toplevelUri) as Classification;

            return c;
        }
        private static Classification CreateFolderStructure(string p, string parent, string lowertype, Database db, int lvlvalue)
        {
            string strnew = null;
            Classification par = new Classification(db, parent);
            Classification c = new Classification(db);
            c = par.NewLowerLevel();
            c.Name = CleanupFolderName(p);
            if (lowertype == "c")
            {
                c.ChildPattern = "/gggg";
            }
            else
            {
                c.RecordPattern = "/gggg";
            }
            c.LevelNumber = "/" + lvlvalue.ToString().PadLeft(4, '0');
            //c.Save();
            strnew = c.Title;
            toplevelUri = c.Uri;
            return c;
        }
        private static string CleanupFolderName(string name)
        {
            string strNewName = null;
            strNewName = name.Replace("-", "~");
            return strNewName;
        }
        private static int FindStartingNumber(long toplevelUri, Database db)
        {
            int intChildCnt = 0;
            List<int> classNumb = new List<int>();
            TrimMainObjectSearch obj = new TrimMainObjectSearch(db, BaseObjectTypes.Classification);
            obj.SetSearchString("parent:" + toplevelUri);
            if (obj.FastCount > 0)
            {
                foreach (Classification c in obj)
                {
                    var kkkk = c.LevelNumberUncompressed.Remove(0, 2);
                    classNumb.Add(Convert.ToInt32(kkkk));
                }
                intChildCnt = classNumb.Max();
            }
            else
            {
                intChildCnt = 0;
            }
            return intChildCnt + 10;
        }
        private static long CreateNewRMfolder(string str4, long classification, Database db)
        {
            long uri = 0;
            Classification objclas = new Classification(db, classification);
            long clasUri = objclas.Uri;
            RecordType rt = new RecordType(db, 1);
            Record r = new Record(rt);
            Classification c = new Classification(db, clasUri);
            Location lo = new Location(db, 5503);
            r.OwnerLocation = lo;
            r.SetAssignee(lo);

            r.Classification = c;
            r.Title = str4;
            //r.Save();
            uri = r.Uri;

            return uri;
        }
        private static long CreateNewRMSubFolder(long conturi, string str5, Database db)
        {
            long uri = 0;
            Record rCont = new Record(db, conturi);
            RecordType rtsub = new RecordType(db, 5);
            Record rsub = new Record(rtsub);
            rsub.Title = str5;
            rsub.Container = rCont;
            rsub.SetAssignee(rCont.Assignee);
            //rsub.Save();
            uri = rsub.Uri;
            return uri;
        }
        private static Record CheckifFolderAlreadyExists(Classification cl3, string p, Database db)
        {
            Record rec = null;
            TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
            objS.SetSearchString("classification:" + cl3.Uri);
            //if (objS.FastCount == 0)
            //    Console.WriteLine("No results for this folder under the class: " + cl3.Title + "/" + p);
            foreach (Record r in objS)
            {
                //Console.WriteLine("Loop through folders under class:"+p+"/"+r.Title);
                if (r.Title == p.Trim())
                {
                    rec = r;
                }
            }
            if (rec != null)
            {
                Console.WriteLine("Found record:" + rec.Uri);
            }
            else
            {
                Console.WriteLine("Record is null");
            }
            return rec;
        }
        private static Record CheckSubFolderAlreadyExists(long folder, string p, Database db)
        {
            Record rec = null;

            TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
            objS.SetSearchString("container:" + folder);
            foreach (Record r in objS)
            {
                if (r.Title == p.Trim())
                {
                    rec = r;
                }
                //else
                //{
                //    rec = null;
                //}
            }


            return rec;
        }
    }
}
