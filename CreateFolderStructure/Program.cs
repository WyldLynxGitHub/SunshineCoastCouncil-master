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

namespace CreateFolderStructure
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
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Worksheet MySheetV = null;
        private static int y = 1;
        private static long toplevelUri = 0;
        private static string toplevel = null;
        private static string FileLocation = null;
        private static long ProcessfromRow = 0;
        private static long ProcesstoRow = 0;
        private static bool isInteger = false;

    
        private static void Repeatmenu()
        {
            Console.WriteLine("");
            Console.WriteLine("1. Process Spreadsheet");
            Console.WriteLine("2. Create and Sync folders");
            Console.WriteLine("3. Exit");
        }
        private static void ManageMenu(int menuitem)
        {
            switch (menuitem)
            {
                case (1):
                    Console.WriteLine("Enter location of excel worksheet");
                    FileLocation = Console.ReadLine();
                    ProcessClassFolders();
                    break;
                case (2):                  
                    processtoRM();
                    break;
                case (500):
                    Console.WriteLine("Not releaseed.");
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
                int imenuItem = 0;
            Console.WriteLine("1. Process Spreadsheet to JSon");
            Console.WriteLine("2. Create and Sync folders");
            Console.WriteLine("3. Exit");
            isInteger = int.TryParse(Console.ReadLine(), out imenuItem);
            if (isInteger)
            {
                ManageMenu(imenuItem);
            }
        }
        private static void ProcessClassFolders()
        {
            string iRow = null;
            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(@FileLocation);
                MySheet = (Excel.Worksheet)MyBook.Sheets["Folder Structure"];
                MySheetV = (Excel.Worksheet)MyBook.Sheets["Variables"];
                Console.WriteLine("Connected to Spread Sheet");
                //Console.ReadLine();
                Console.WriteLine("Processing, please wait. If its a large SS then go make a coffee.");
                ProcessfromRow = Convert.ToInt64(MySheetV.Cells[11, 2].Value);
                ProcesstoRow = Convert.ToInt64(MySheetV.Cells[12, 2].Value);

                string str1 = null, str2 = null, str3 = null, str4 = null, str5 = null;
                bool createnew = false;
                long fileconturi = 0, recuri = 0;
                //EmpList.Add(emp);

                List<Sheets> sheets = new List<Sheets>();
                Sheets sheet = new Sheets();
                sheet.SheetName = FileLocation;
                List<Rows> rows = new List<Rows>();
                Rows r = new Rows();
                
                for (var i = 2; i < ProcesstoRow + 1; i++)
                {
                    if(i==350)
                    {
                        var dsd = "sfdsf";
                    }
                    
                    iRow = i.ToString();


                    //First child level
                    bool bLevel3New = false;
                    createnew = false;
                    if (MySheet.Cells[i, 1].Value2 != null)
                    {
                        str1 = MySheet.Cells[i, 1].Value.ToString();

                        createnew = true;

                    }

                    if (MySheet.Cells[i, 2].Value2 != null)
                    {
                        str2 = MySheet.Cells[i, 2].Value.ToString();
                        createnew = true;
                    }

                    if (MySheet.Cells[i, 3].Value2 != null)
                    {
                        str3 = MySheet.Cells[i, 3].Value.ToString();
                        createnew=true;
                        bLevel3New = true;
                    }
                    if(createnew)
                    {
                        r = new Rows();
                        r.level1 = str1;
                        r.level2 = str2;
                        r.level3 = str3;

                        r.SheetLine = i;
                        rows.Add(r);
                    }
                    
                    // New
                    if (MySheet.Cells[i, 4].Value2 != null)
                    {
                        if (bLevel3New == true)
                        {
                            List<Folder> lstFolder = new List<Folder>();
                            rows.LastOrDefault().Folders = lstFolder;
                        }
                        str4 = MySheet.Cells[i, 4].Value.ToString();
                        //int lvl4LoopCount = 0;

                        //Check to see if the group identifyer is not blank and has a g
                        if (MySheet.Cells[i, 12].value2 != null && MySheet.Cells[i, 12].value.ToString() == "g")
                        {
                            Console.WriteLine("Grouped string found in level 4");
                            List<string> lstLvl4 = str4.Split('/').ToList();
                            foreach (var kkk in lstLvl4.Where(x => x.Length > 2))
                            {
                                Folder fol = new Folder();
                                fol.folder = kkk.Trim();
                                fol.SheetLine = i;
                                rows.LastOrDefault().Folders.Add(fol);
                                //lstFolder.Add(fol);
                                fileconturi = recuri;
                                if (MySheet.Cells[i, 5].Value2 != null)
                                {
                                    str5 = MySheet.Cells[i, 5].Value.ToString();
                                    List<Subfolder> lstsubfolder = new List<Subfolder>();
                                    if (MySheet.Cells[i, 13].value2 != null && MySheet.Cells[i, 13].value.ToString() == "g")
                                    {
                                        List<string> lstLvl5 = str5.Split('/').ToList();
                                        foreach (var lll in lstLvl5.Where(x => x.Length > 2))
                                        {
                                            Subfolder sf = new Subfolder();
                                            sf.subfolder = lll.Trim();
                                            sf.SheetLine = i;

                                            lstsubfolder.Add(sf);
                                        }
                                        fol.SubFolders = lstsubfolder;
                                    }
                                    else
                                    {
                                        Subfolder sf = new Subfolder();
                                        sf.subfolder = str5.Trim();
                                        sf.SheetLine = i;
                                        lstsubfolder.Add(sf);
                                        fol.SubFolders = lstsubfolder;
                                    }
                                }
                            }
                        }
                        else
                        {
                            Folder fol = new Folder();
                            fol.folder = str4;
                            fol.SheetLine = i;
                            fileconturi = recuri;
                            if (MySheet.Cells[i, 5].Value2 != null)
                            {
                                List<Subfolder> lstsf = new List<Subfolder>();
                                str5 = MySheet.Cells[i, 5].Value.ToString();
                                //r.Subfolder = str5;
                                if (MySheet.Cells[i, 13].value2 != null && MySheet.Cells[i, 13].value.ToString() == "g")
                                {
                                    List<string> lstLvl5 = str5.Split('/').ToList();
                                    foreach (var lll in lstLvl5.Where(x => x.Length > 2))
                                    {
                                        Subfolder sf = new Subfolder();
                                        sf.subfolder = lll.Trim();
                                        sf.SheetLine = i;
                                        lstsf.Add(sf);
                                    }
                                }
                                else
                                {
                                    Subfolder sf = new Subfolder();
                                    sf.subfolder = str5;
                                    sf.SheetLine = i;
                                    if (MySheet.Cells[i, 7].value2 != null)
                                    {
                                        sf.MigrateLoc = MySheet.Cells[i, 7].Value.ToString();
                                    }

                                    lstsf.Add(sf);
                                }
                                fol.SubFolders = lstsf;
                            }
                            else
                            {
                                if (MySheet.Cells[i, 7].value2 != null)
                                {
                                    fol.MigrateLoc = MySheet.Cells[i, 7].Value.ToString();
                                }
                            }
                            rows.LastOrDefault().Folders.Add(fol);
                        }
                    }
                    else
                    {
                        if (MySheet.Cells[i, 5].Value2 != null)
                        {
                            if (rows.LastOrDefault().Folders.LastOrDefault().SubFolders==null)
                            {
                                List<Subfolder> lstsf = new List<Subfolder>();
                                rows.LastOrDefault().Folders.LastOrDefault().SubFolders = lstsf;
                            }
                            str5 = MySheet.Cells[i, 5].Value.ToString();
                            if (MySheet.Cells[i, 13].value2 != null && MySheet.Cells[i, 13].value.ToString() == "g")
                            {
                                List<string> lstLvl5 = str5.Split('/').ToList();
                                foreach (var lll in lstLvl5.Where(x => x.Length > 2))
                                {
                                    Subfolder sf = new Subfolder();
                                    sf.subfolder = lll.Trim();
                                    sf.SheetLine = i;
                                    if (MySheet.Cells[i, 7].value2 != null)
                                    {
                                        sf.MigrateLoc = MySheet.Cells[i, 7].Value.ToString();
                                    }
                                    rows.LastOrDefault().Folders.LastOrDefault().SubFolders.Add(sf);
                                }   
                            }
                            else
                            {
                                Subfolder sf = new Subfolder();
                                sf.subfolder = str5.Trim();
                                sf.SheetLine = i;
                                if (MySheet.Cells[i, 7].value2 != null)
                                {
                                    sf.MigrateLoc = MySheet.Cells[i, 7].Value.ToString();
                                }
                                //lstsubfolder.Add(sf);
                                //fol.SubFolders = lstsubfolder;
                                //rows.LastOrDefault().Folders.LastOrDefault().SubFolders = lstsubfolder;
                                rows.LastOrDefault().Folders.LastOrDefault().SubFolders.Add(sf);
                            }
                        }

                    }




                    //End New
                }




                sheet.Rows = rows;
                sheets.Add(sheet);
                //
                string json = JsonConvert.SerializeObject(sheets.ToArray());
                //FileLocation
                var loc = Path.GetDirectoryName(FileLocation);
                System.IO.File.WriteAllText(@loc + "\\SSprocessTIP.json", json);
                //
                MyBook.Save();
                MyBook.Close(null, null, null);
                MyApp.Quit();
                Console.WriteLine("Successful in building Json file.");
                Repeatmenu();

                //}
            }
            catch (Exception exc)
            {
                MyBook.Save();
                MyBook.Close(null, null, null);
                MyApp.Quit();
                //
                Console.WriteLine("Error: " + exc.Message.ToString()+" - "+iRow);
                return;
            }
        }
        private static void processFolders()
        {
            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(@FileLocation);
                MySheet = (Excel.Worksheet)MyBook.Sheets["Folder Structure"];
                MySheetV = (Excel.Worksheet)MyBook.Sheets["Variables"];
                Console.WriteLine("Connected to Spread Sheet");
                Console.ReadLine();
                ProcessfromRow = Convert.ToInt64(MySheetV.Cells[11, 2].Value);
                ProcesstoRow = Convert.ToInt64(MySheetV.Cells[12, 2].Value);
                    string str1 = null, str2 = null, str3 = null, str4 = null, str5 = null;
                    bool createnew = false;
                    long fileconturi = 0, recuri = 0;
                    int lvl1Value = 10, lvl2Value = 10;
                    //EmpList.Add(emp);
                    
                    List<Sheets> sheets = new List<Sheets>();
                    Sheets sheet = new Sheets();
                    sheet.SheetName = FileLocation;
                    List<Rows> rows = new List<Rows>();
                    Rows r = new Rows();
                    for (var i = 2; i < ProcesstoRow+1; i++)
                    {
                        
                        //First child level
                        if (MySheet.Cells[i, y].Value2 != null)
                        {
                            str1 = MySheet.Cells[i, y].Value.ToString();

                                createnew = true;
                                lvl1Value = lvl1Value + 10;
                                //lvl2Value = 10;
                                //lvl3Value = 10;
                            y = y + 1;
                        }
                        else
                        {
                            y = y + 1;
                        }
                        if (MySheet.Cells[i, y].Value2 != null)
                        {
                            str2 = MySheet.Cells[i, y].Value.ToString();
                            if (createnew)
                            {
                                lvl2Value = lvl2Value + 10;
                                //lvl3Value = 10;
                            }

                            y = y + 1;
                        }
                        else
                        {
                            y = y + 1;
                        }
                        if (MySheet.Cells[i, y].Value2 != null)
                        {
                            str3 = MySheet.Cells[i, y].Value.ToString();
                            if (createnew)
                            {
                                r = new Rows();
                                    r.level1 = str1;
                                    r.level2 = str2;
                                    r.level3 = str3;
                            }
                            y = y + 1;
                        }
                        else
                        {
                            y = y + 1;
                            r = rows.LastOrDefault();
                        }

                            List<Folder> lstFolder = new List<Folder>();

                            if (MySheet.Cells[i, 4].Value2 != null)
                            {

                                str4 = MySheet.Cells[i, 4].Value.ToString();

                                //Check to see if the group identifyer is not blank and has a g
                                if (MySheet.Cells[i, 12].value2 != null && MySheet.Cells[i, 12].value.ToString() == "g")
                                {
                                    Console.WriteLine("Grouped string found in level 4");
                                    List<string> lstLvl4 = str4.Split('/').ToList();
                                    foreach (var kkk in lstLvl4.Where(x => x.Length > 2))
                                    {
                                        Folder fol = new Folder();
                                        fol.folder = kkk.Trim();
                                        lstFolder.Add(fol);
                                        fileconturi = recuri;
                                        if (MySheet.Cells[i, 5].Value2 != null)
                                        {
                                            str5 = MySheet.Cells[i, 5].Value.ToString();
                                            List<Subfolder> lstsubfolder = new List<Subfolder>();
                                            if (MySheet.Cells[i, 13].value2 != null && MySheet.Cells[i, 13].value.ToString() == "g")
                                            {
                                                List<string> lstLvl5 = str5.Split('/').ToList();
                                                foreach (var lll in lstLvl5.Where(x => x.Length > 2))
                                                {
                                                    Subfolder sf = new Subfolder();
                                                    sf.subfolder = lll.Trim();
                                                    lstsubfolder.Add(sf);
                                                }
                                                fol.SubFolders = lstsubfolder;
                                            }
                                            else
                                            {
                                                Subfolder sf = new Subfolder();
                                                sf.subfolder = str5.Trim();
                                                lstsubfolder.Add(sf);
                                                fol.SubFolders = lstsubfolder;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Folder fol = new Folder();
                                    fol.folder = str4;
                                    fileconturi = recuri;
                                    if (MySheet.Cells[i, 5].Value2 != null)
                                    {
                                        List<Subfolder> lstsf = new List<Subfolder>();
                                        str5 = MySheet.Cells[i, 5].Value.ToString();
                                        //r.Subfolder = str5;
                                        if (MySheet.Cells[i, 13].value2 != null && MySheet.Cells[i, 13].value.ToString() == "g")
                                        {
                                            List<string> lstLvl5 = str5.Split('/').ToList();
                                            foreach (var lll in lstLvl5.Where(x => x.Length > 2))
                                            {
                                                Subfolder sf = new Subfolder();
                                                sf.subfolder = lll.Trim();
                                                lstsf.Add(sf);
                                                //recsuburi = CreateNewRMSubFolder(recuri, lll.Trim(), db, i);
                                            }
                                        }
                                        else
                                        {
                                            Subfolder sf = new Subfolder();
                                            sf.subfolder = str5;

                                            if (MySheet.Cells[i, 7].value2 != null)
                                            {
                                                sf.MigrateLoc = MySheet.Cells[i, 7].Value.ToString();
                                            }

                                            lstsf.Add(sf);
                                            //fol.SubFolders=lstsf;
                                            //
                                            //recsuburi = CreateNewRMSubFolder(recuri, str5, db, i);
                                        }
                                        fol.SubFolders = lstsf;
                                    }
                                    else
                                    {
                                        if (MySheet.Cells[i, 7].value2 != null)
                                        {
                                            fol.MigrateLoc = MySheet.Cells[i, 7].Value.ToString();
                                        }
                                    }

                                    lstFolder.Add(fol);
                                }
                               
                                r.Folders = lstFolder;
                            }
                            y = 1;

                            if (str4 != null)
                            {

                                MySheet.Cells[i, 8] = fileconturi.ToString();


                            }
                            else
                            {
                                //Remove for basic testing
                                //long Newuri = CreateNewRMfolder(str1, str2, str3, str4, str5, str3Title);

                                //
                                //Console.WriteLine("New folder created (Uri: " + fileconturi.ToString() + ")");
                            }
                            r.SheetLine = i;

                        //7
                            //r.RMuri = fileconturi.ToString();
                            //r.Timestamp = DateTime.Now;
                            rows.Add(r);
                    }
                    sheet.Rows = rows;
                    sheets.Add(sheet);
                    //
                    string json = JsonConvert.SerializeObject(sheets.ToArray());
                    //FileLocation
                    var loc = Path.GetDirectoryName(FileLocation);
                    System.IO.File.WriteAllText(@loc+"\\SSprocessTIP.json", json);
                    //
                    MyBook.Save();
                    MyBook.Close(null, null, null);
                    MyApp.Quit();

                //}
            }
            catch (Exception exc)
            {
                MyBook.Save();
                MyBook.Close(null, null, null);
                MyApp.Quit();
                //
                Console.WriteLine("Error: "+exc.Message.ToString());
                return;
            }
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
                }
                else
                {
                    Console.WriteLine("Document location does not exist: " + FileLoc);
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message.ToString());
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
        private static bool CheckFolderlevel(long toplevel, Database db)
        {
            bool bExists = false;
            //try
            //{
            Classification c = db.FindTrimObjectByUri(BaseObjectTypes.Classification, toplevel) as Classification;
                    //TrimMainObjectSearch objClas = new TrimMainObjectSearch(db, BaseObjectTypes.Classification);
                    //objClas.SetSearchString("title:" + toplevel);
                    if (c!=null)
                    {
                        Console.WriteLine("First Level Folder Exists");
                        bExists = true;
                    }
                    else
                    {
                        Console.WriteLine("First Level Folder could not be found, check the following folder exist in RM: "+toplevel+" in "+db.Database.Name+" with the dataset ID: "+db.Id);
                        bExists = false;
                    }
                   
                //}
            //}
            //catch (Exception exp)
            //{
            //    Console.WriteLine("Connect to RM SDK issue, Error: " + exp.Message.ToString());
            //}
            return bExists;
        }
        private static string CleanupFolderName(string name)
        {
            string strNewName = null;
            strNewName = name.Replace("-", "~");
            return strNewName;
        }
        private static void processtoRM()
        {
            long recuri = 0;
            long folderUri = 0;
            long classUri = 0;
            
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
                Console.WriteLine("Enter location of Json file");
                FileLocation = Console.ReadLine();
                List<Sheets> myMessage = JsonConvert.DeserializeObject<List<Sheets>>(File.ReadAllText(@FileLocation));
                Console.WriteLine("Enter the Classification Uri for this process:");
                foreach (var file in myMessage)
                {

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

                    Console.WriteLine("Enter the SS line number to process from, minimum should be 2.");
                    long linenumber = Convert.ToInt64(Console.ReadLine());
                    long tl = toplevelUri;
                    foreach (Rows r in file.Rows.Where(x => x.SheetLine >= linenumber))
                    {

                        try
                        {
                            Classification cl1 = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, toplevel + " - " + CleanupFolderName(r.level1));
                            if (cl1 == null)
                            {
                                cl1 = CreateFolderStructure(r.level1, toplevel, "c", db, FindStartingNumber(tl, db));
                                classUri = cl1.Uri;
                            }

                            Classification cl2 = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, cl1.Title + " - " + CleanupFolderName(r.level2));
                            if (cl2 == null)
                            {
                                cl2 = CreateFolderStructure(r.level2, cl1.Title, "c", db, FindStartingNumber(cl1.Uri, db));
                                classUri = cl2.Uri;
                            }

                            Classification cl3 = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, cl2.Title + " - " + CleanupFolderName(r.level3));
                            if (cl3 == null)
                            {
                                cl3 = CreateFolderStructure(r.level3, cl2.Title, "r", db, FindStartingNumber(cl2.Uri, db));
                                classUri = cl3.Uri;
                            }

                            if (r.Folders != null)
                            {
                                if (r.Folders.Count() > 0)
                                {
                                    foreach (var f in r.Folders)
                                    {
                                        Record chkRec = CheckifFolderAlreadyExists(cl3, f.folder, db);
                                        //
                                        if (chkRec == null)
                                        {
                                            //Console.WriteLine("Did not find existing folder: " + cl3.Title + " " + f.folder);
                                            try
                                            {
                                                recuri = CreateNewRMfolder(f.folder, cl3.Uri, db);
                                                f.RMuri = recuri.ToString();
                                                Console.WriteLine("Created new folder: Uri " + recuri.ToString());
                                            }
                                            catch (Exception exp)
                                            {
                                                Console.WriteLine("Error creating folder: " + exp.Message.ToString() + " - " + f.folder + " (" + cl3.Title + ")");
                                                f.Error = exp.Message.ToString();
                                            }
                                            folderUri = recuri;
                                        }
                                        else
                                        {
                                            //Console.WriteLine("Found existing folder: " + chkRec.Number);
                                            recuri = chkRec.Uri;
                                            //r.RMuri=recuri.ToString();
                                            folderUri = recuri;
                                            if (f.RMuri == null)
                                                f.RMuri = chkRec.Uri.ToString();
                                        }
                                        if (f.SubFolders != null)
                                        {
                                            if (f.SubFolders.Count() > 0)
                                            {
                                                foreach (var s in f.SubFolders)
                                                {
                                                    Record chkRecSub = CheckSubFolderAlreadyExists(folderUri, s.subfolder, db);
                                                    if (chkRecSub == null)
                                                    {
                                                        //Console.WriteLine("Did not find existing sub folder: " + s.subfolder);
                                                        try
                                                        {
                                                            recuri = CreateNewRMSubFolder(folderUri, s.subfolder, db);
                                                            s.RMuri = recuri.ToString();
                                                            Console.WriteLine("Created new sub folder: Uri " + recuri.ToString());
                                                        }
                                                        catch (Exception exp)
                                                        {
                                                            Console.WriteLine("Error creating sub folder: " + exp.Message.ToString() + " - " + s.subfolder + " (" + cl3.Title + ")");
                                                            s.Error = exp.Message.ToString();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //Console.WriteLine("Found existing sub folder: " + chkRecSub.Number);
                                                        recuri = chkRecSub.Uri;
                                                        if (s.RMuri == null)
                                                            s.RMuri = chkRecSub.Uri.ToString();
                                                        //r.RMuri=
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("No sub folders to create: ");
                                        }


                                    }

                                }
                                else
                                {
                                    r.RMuri = classUri.ToString();
                                }
                            }


                            r.RMuri = recuri.ToString();
                            r.Timestamp = DateTime.Now;
                            //r.Error = "All good";
                            Console.WriteLine("Processed SS line " + r.SheetLine.ToString());
                            //string jsont = JsonConvert.SerializeObject(myMessage.ToArray());
                            //System.IO.File.WriteAllText(@Path.GetDirectoryName(FileLocation)+"\\Temptemp.json", jsont);
                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine("Error: " + exp.Message.ToString());
                            //Console.ReadLine();
                            r.Error = exp.Message.ToString();
                            //break;
                        }
                        //}

                    }
                }

            Console.WriteLine("Completed processing.");
            string json = JsonConvert.SerializeObject(myMessage.ToArray());
            string strFileWrite = Guid.NewGuid().ToString();
            System.IO.File.WriteAllText(@Path.GetDirectoryName(FileLocation) + "\\" + strFileWrite + ".json", json);
            Console.WriteLine("Json file update: " + strFileWrite);
            Repeatmenu();

            }
            //}


            
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
            c.Save();
            strnew = c.Title;
            toplevelUri = c.Uri;
            return c;
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
            }
            return rec;
        }
        private static Record CheckifFolderAlreadyExists(Classification cl3, string p, Database db)
        {
            Record rec = null;
            TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
            objS.SetSearchString("classification:" + cl3.Uri);
            foreach (Record r in objS)
            {
                if (r.Title == p.Trim())
                {
                    rec = r;
                }
            }
            return rec;
        }
        private static int FindStartingNumber(long toplevelUri, Database db)
        {
            int intChildLast = 0;
            List<int> classNumb = new List<int>();
            List<Classification> cl = new List<Classification>();
            
            TrimMainObjectSearch obj = new TrimMainObjectSearch(db, BaseObjectTypes.Classification);
            
            obj.SetSearchString("parent:" + toplevelUri);
            obj.SetSortString("uri");
            
            if (obj.FastCount > 0)
            {
                foreach (Classification c in obj)
                {
                    cl.Add(c);
                    var kkkk = c.LevelNumberUncompressed.Remove(0, 1);
                    classNumb.Add(Convert.ToInt32(kkkk));
                }
                intChildLast= classNumb.Last();
            }
            Console.WriteLine("Calculating next folder number: "+toplevelUri.ToString()+" Container count: "+obj.FastCount.ToString()+ " Last number: "+(intChildLast+10).ToString());
            return intChildLast + 10;
        }
        private static Classification getToplevel(long toplevelUri, Database db)
        {
            Classification c = db.FindTrimObjectByUri(BaseObjectTypes.Classification, toplevelUri) as Classification;

            return c;
        }
        private static long CreateNewRMfolder(string str4, long classification, Database db)
        {
            long uri = 0;
            Classification objclas = new Classification(db, classification);
            long clasUri = objclas.Uri;
            RecordType rt = new RecordType(db, 1);
            Record r = new Record(rt);
            Classification c = new Classification(db, clasUri);
            //Location lo = new Location(db, 5503);
            //r.OwnerLocation = lo;
            //r.SetAssignee(lo);

            r.Classification = c;
            r.Title = str4;
            r.Save();
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
            rsub.Save();
            uri = rsub.Uri;
            return uri;
        }
        private static void SynctoExcel(string excel, string json)
        {
            List<Sheets> myMessage = JsonConvert.DeserializeObject<List<Sheets>>(File.ReadAllText(@FileLocation));
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(@FileLocation);
            MySheet = (Excel.Worksheet)MyBook.Sheets["Folder Structure"];
            //
            foreach (var file in myMessage)
            {
                foreach (Rows r in file.Rows)
                {
                    var snumber = r.SheetLine;
                    MySheet.Cells[snumber, 8] = r.RMuri;
                    MySheet.Cells[snumber, 9] = r.OriginalFileCount;
                    MySheet.Cells[snumber, 10] = r.RMFileCount;

                }
            }

            MyBook.Save();
            MyBook.Close(null, null, null);
            MyApp.Quit();

        }
    }
}
