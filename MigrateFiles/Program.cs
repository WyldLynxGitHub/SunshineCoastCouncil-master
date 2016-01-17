using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office;
using System.Data;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using HP.HPTRIM.SDK;
using Newtonsoft.Json;

namespace MigrateFiles
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
        public int OriginalFileCount { get; set; }
        public int RMFileCount { get; set; }
        public DateTime Timestamp { get; set; }
        public string Error { get; set; }
        public string level1 { get; set; }
        public string level2 { get; set; }
        public string level3 { get; set; }
        //public string Folder { get; set; }
        //public string Subfolder { get; set; }
        public List<Folder> Folders { get; set; }
    }
    public class Folder
    {
        public long SheetLine { get; set; }
        public string folder { get; set; }
        public string RMuri { get; set; }
        public string RMString { get; set; }
        public string Error { get; set; }
        public string MigrateLoc { get; set; }
        public List<Subfolder> SubFolders { get; set; }
    }
    public class Subfolder
    {
        public long SheetLine { get; set; }
        public string subfolder { get; set; }
        public string RMuri { get; set; }
        public string RMString { get; set; }
        public string Error { get; set; }
        public string MigrateLoc { get; set; }
    }
    class Program
    {
        private static string FileLocation = null;
        private static bool isInteger = false;
    
        private static void Repeatmenu()
        {
            Console.WriteLine("");
            Console.WriteLine("1. Migrate files from SS");
            Console.WriteLine("2. Migrate files from individual location");
            Console.WriteLine("3. Syncronise migrated email names");
            Console.WriteLine("4. Exit");
        }
        private static void ManageMenu(int menuitem)
        {
            switch (menuitem)
            {
                case (1):
                    processfiles();
                    break;
                case (2):
                    processSingleFolder();
                    break;
                case (3):
                    Environment.Exit(0);
                    break;
                default:

                    break;
            }

        }
        static void Main(string[] args)
        {
            //Form1 frm = new Form1();
            //frm.Activate();
           // ApplicationExcepti
                int imenuItem = 0;
                Console.WriteLine("1. Migrate files from SS");
                Console.WriteLine("2. Migrate files from individual location");
                //Console.WriteLine("3. Syncronise migrated email names");
            Console.WriteLine("4. Exit");
            isInteger = int.TryParse(Console.ReadLine(), out imenuItem);
            if (isInteger)
            {
                ManageMenu(imenuItem);
            }
        }
        private static void processSingleFolder()
        {
            if (!TrimApplication.HasBeenInitialized)
            {
                TrimApplication.Initialize();
            }
            TrimApplication.HasUserInterface = true;
            using (Database db = new Database())
            {
                try
                {
                    db.Connect();
                    Console.WriteLine("Connected to RM successfully as : " + db.CurrentUser.Name + " on " + db.Id + "/" + db.WorkgroupServerName);
                    Console.WriteLine("Hit q to quit or any other key to continue");
                    if (Console.ReadLine() == "q")
                    {
                        Environment.Exit(0);
                    }

                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error: " + exp);
                    return;
                }
                //try
                //{
                    //db.Connect();
                bool bDelete = false;
                Console.WriteLine("Enter location of files to be migrated");
                string loc = Console.ReadLine();
                Console.WriteLine("Enter RM uri folder.");
                var rmUri = Console.ReadLine();
                Console.WriteLine("Confirm Y to delete source or N to leave.");
                switch (Console.ReadLine())
                {
                    case "Y":
                    case "y":
                        bDelete = true;
                        break;
                    case "N":
                    case "n":
                        bDelete = false;
                        break;
                    default:
                        bDelete = false;
                        break;

                }
                try
                {
                    if (Directory.Exists(loc))
                    {
                        var txtFiles = Directory.EnumerateFiles(loc, "*.*", SearchOption.TopDirectoryOnly);
                        foreach (var ff in txtFiles)
                        {
                            if (CreateNewRMRecord(ff, Convert.ToInt64(rmUri), db) > 0)
                            {
                                Console.WriteLine("File created in RM: " + ff);
                                if (bDelete)
                                {
                                    File.Delete(ff);
                                    Console.WriteLine("File deleted: " + ff);
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Could not find supplied folder.");
                    }
                }
                catch(Exception exp)
                {
                    Console.WriteLine("Error: "+exp.Message.ToString());
                }
            }
        }
        private static void processfiles()
        {
            if (!TrimApplication.HasBeenInitialized)
            {
                TrimApplication.Initialize();
            }
            TrimApplication.HasUserInterface = true;
            using (Database db = new Database())
            {
                try
                {
                    db.Connect();
                    //db.PreventDuplicatedDocuments = true;
                    Console.WriteLine("Connected to RM successfully as : " + db.CurrentUser.Name + " on " + db.Id + "/" + db.WorkgroupServerName);
                    Console.WriteLine("Hit q to quit or any other key to continue");
                    if (Console.ReadLine() == "q")
                    {
                        Environment.Exit(0);
                    }

                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error: " + exp);
                    return;
                }

                Console.WriteLine("Enter location of Json file");
                FileLocation = Console.ReadLine();

                int intMigratedfiles = 0;
                bool bDelete = false;
                Console.WriteLine("Confirm Y to delete source or N to leave.");
                switch (Console.ReadLine())
                {
                    case "Y":
                    case "y":
                        bDelete = true;
                        break;
                    case "N":
                    case "n":
                        bDelete = false;
                        break;
                    default:
                        bDelete = false;
                        break;

                }
                List<Sheets> myMessage = JsonConvert.DeserializeObject<List<Sheets>>(File.ReadAllText(@FileLocation));
                try
                {
                    foreach (var file in myMessage)
                    {
                        foreach (Rows r in file.Rows)
                        {
                            if (r.Folders != null)
                            {
                                foreach (Folder f in r.Folders)
                                {
                                    if (f.MigrateLoc != null)
                                    {
                                        if (Directory.Exists(f.MigrateLoc))
                                        {
                                            intMigratedfiles = 0;
                                            var txtFiles = Directory.EnumerateFiles(f.MigrateLoc, "*.*", SearchOption.TopDirectoryOnly);
                                            r.OriginalFileCount = txtFiles.Count();
                                            foreach (var ff in txtFiles)
                                            {
                                                if (CreateNewRMRecord(ff, Convert.ToInt64(f.RMuri), db) > 0)
                                                {
                                                    Console.WriteLine("File created in RM: " + ff);
                                                    intMigratedfiles++;

                                                    if (bDelete)
                                                    {
                                                        try
                                                        {
                                                            File.Delete(ff);
                                                            Console.WriteLine("File deleted: " + ff);
                                                        }
                                                        catch (Exception exp)
                                                        {
                                                            Console.WriteLine("Error deleting File: " + exp.Message.ToString());
                                                        }
                                                    }
                                                }
                                            }
                                            db.LogExternalEvent("There were "+txtFiles.Count().ToString()+" files in " + f.MigrateLoc + " and "+intMigratedfiles.ToString()+" successfully migrated. Failed file migrations are listed individually.", BaseObjectTypes.Record, Convert.ToInt64(f.RMuri), true);
                                        }
                                        else
                                        {
                                            db.LogExternalEvent("File location passed for migrating incorrect: " + f.MigrateLoc, BaseObjectTypes.Record, Convert.ToInt64(f.RMuri), true);
                                        }
                                    }
                                    if (f.SubFolders != null)
                                    {
                                        foreach (Subfolder sf in f.SubFolders)
                                        {
                                            if (sf.MigrateLoc != null)
                                            {
                                                Console.WriteLine("File Location: " + sf.MigrateLoc);
                                                if (Directory.Exists(sf.MigrateLoc))
                                                {
                                                    intMigratedfiles = 0;
                                                    var txtFiles = Directory.EnumerateFiles(sf.MigrateLoc, "*.*", SearchOption.TopDirectoryOnly);
                                                    foreach (var ff in txtFiles)
                                                    {
                                                        if (CreateNewRMRecord(ff, Convert.ToInt64(sf.RMuri), db) > 0)
                                                        {
                                                            Console.WriteLine("File created in RM: " + ff);
                                                            intMigratedfiles++;
                                                            if (bDelete)
                                                            {
                                                                try
                                                                {
                                                                    File.Delete(ff);
                                                                    Console.WriteLine("File deleted: " + ff);
                                                                }
                                                                catch (Exception exp)
                                                                {
                                                                    Console.WriteLine("Error deleting File: " + exp.Message.ToString());
                                                                }
                                                            }
                                                        }
                                                    }
                                                    db.LogExternalEvent("There were " + txtFiles.Count().ToString() + " files in " + f.MigrateLoc + " and " + intMigratedfiles.ToString() + " successfully migrated. Failed file migrations are listed individually.", BaseObjectTypes.Record, Convert.ToInt64(f.RMuri), true);
                                                }
                                            }
                                            else
                                            {
                                                db.LogExternalEvent("File location passed for migrating incorrect: " + f.MigrateLoc, BaseObjectTypes.Record, Convert.ToInt64(f.RMuri), true);
                                            }
                                        }
                                    }
                                }
                                r.RMFileCount = intMigratedfiles;
                            } 
                        }
                    }
                    Console.WriteLine("Completed file migration.");
                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error Migrating files, Error: " + exp.Message.ToString());
                }
            }
            //}
        }
        private static long CreateNewRMRecord(string FileLoc, long cont, Database db)
        {
            long uri = 0;
            string newtitle = null;
            try
            {
                if (File.Exists(FileLoc))
                {
                    RecordType rt = null;
                   // Console.WriteLine("Document location: " + FileLoc);
                    string strExt = Path.GetExtension(FileLoc);
                    switch (strExt)
                    {
                        //Documents
                        case ".PDF":
                        case ".pdf":
                        case ".DOC":
                        case ".doc":
                        case ".XLS":
                        case ".xls":
                        case ".DOCX":
                        case ".docx":
                        case ".XLSX":
                        case ".xlsx":
                        case ".ppt":
                        case ".PPT":
                        case ".pptx":
                        case ".PPTX":
                        case ".DOTX":
                        case ".dotx":
                        case ".TXT":
                        case ".txt":
                        case ".CSV":
                        case ".csv":
                        case ".PUB":
                        case ".pub":
                        case ".HTML":
                        case ".html":
                        case ".HTM":
                        case ".htm":
                        case ".DWG":
                        case ".dwg":
                        case ".RTF":
                        case ".rtf":
                        case ".XLSM":
                        case ".xlsm":
                        case ".MP4":
                        case ".mp4":
                        case ".DOT":
                        case ".dot":
                        case ".VSD":
                        case ".vsd":
                        case ".ZIP":
                        case ".zip":
                        case ".PSD":
                        case ".psd":
                        case ".RDL":
                        case ".rdl":
                        case ".XML":
                        case ".xml":
                        case ".INDD":
                        case ".MHT": //NS
                        case ".MPP": //NS
                        case ".CFM": //NS
                        case ".XLTM": //NS
                        case ".SHX": //NS   
                        case ".PPSX": //NS
                        case ".CSS": //NS
                        case ".XLTX": //NS
                        case ".JS": //NS"
                        case ".OTF": //NS
                        case ".VST": //NS
                        case ".WMV": //NS
                        case ".WMZ": //NS
                        case ".OFT": //NS
                        case ".MOV": //NS
                        case ".PPS": //NS
                        case ".MDB": //NS
                        case ".EPS": //NS
                        case ".XPS": //NS
                        case ".XLSB": //NS
                        case ".MID": //NS
                        case ".THMX": //NS
                        case ".TMP": //NS 
                        case ".TTF": //NS
                        case ".MXD": //NS
                        case ".ICS": //NS
                        case ".DOCM": //NS
                        case ".BPMN": //NS
                        case ".AI": //NS
                        case ".AVI": //NS
                        case ".DXF": //NS
                        case ".INI": //NS
                        case ".LOG": //NS
                        case ".LST": //NS
                        case ".ODT": //NS
                        case ".PARTIAL": //NS
                        case ".POT": //NS
                        case ".POTX": //NS
                        case ".VSDX": //NS
                        case ".SKP": //NS
                        case ".SHS": //NS
                        case ".RPT": //NS
                        case ".XLA": //NS

                            rt = new RecordType(db, 2);
                            break;
                        //Images
                        case ".JPG":
                        case ".TIF":
                        case ".jpg":
                        case ".tif":
                        case ".PNG":
                        case ".png":
                        case ".JPEG":
                        case ".jpeg":
                        case ".GIF":
                        case ".gif":
                        case ".BMP":
                        case ".bmp":
                        case ".MPEG":
                        case ".TIFF": 
                            rt = new RecordType(db, 14);
                            break;
                        //Mail
                        case ".EML":
                        case ".eml":
                        case ".msg":
                        case ".MSG":
                            rt = new RecordType(db, 3);
                            newtitle = SetEmailTitle(FileLoc);
                            break;
                    }


                        Record rcont = new Record(db, cont);
                        Record r = new Record(rt);
                        Location lo = new Location(db, 5503);
                        if (newtitle != null)
                            r.Title = newtitle;
                        r.SetDocument(FileLoc);
                        r.SetAssignee(lo);
                        r.SetOwnerLocation(lo);
                        r.Container = rcont;
                    
                        r.Save();
                        uri = r.Uri;
                        Console.WriteLine("Migrated document from: " + FileLoc + " and saved as " + r.Number);
                        db.LogExternalEvent("Document migrated from: " + FileLoc, BaseObjectTypes.Record, uri, true);
                }
                else
                {
                    Console.WriteLine("Document location does not exist: " + FileLoc);
                    db.LogExternalEvent("Document location does not exist: "+FileLoc, BaseObjectTypes.Record, Convert.ToInt64(cont), true);
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message.ToString());
                db.LogExternalEvent("Error migrating: " + FileLoc+" Error: "+exp.Message.ToString(), BaseObjectTypes.Record, Convert.ToInt64(cont), true);
            }
            return uri;
        }
        private static string SetEmailTitle(string docdetails)
        {
            string FileName = docdetails.Substring(docdetails.LastIndexOf('\\') + 1);
            //
            int fileExtPos = FileName.LastIndexOf(".");
            if (fileExtPos >= 0)
                FileName = FileName.Substring(0, fileExtPos);

            return FileName;
        }
    }
}
