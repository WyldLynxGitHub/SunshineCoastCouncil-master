using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HP.HPTRIM.SDK;
using System.IO;
using Newtonsoft.Json;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;


namespace wfMigrateFiles
{
    public partial class Form1 : Form
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
        List<ManageFiles> lstmf = new List<ManageFiles>();
        public int intMigratedfiles { get; set; }
        public static string strWorkingLocation = null;
        public class ManageFiles
        {
            [Display(Name = "SS No")]
            public long sheetno { get; set; }
            [Display(Name = "Source")]
            public int SourceCnt { get; set; }
            [Display(Name = "Migrated")]
            public int MigrateCnt { get; set; }
            public string Location { get; set; }
            public string Comments { get; set; }
             [Display(Name = "Eddie Uri")]
            public long MigrateUri { get; set; }
        }
        public Form1()
        {
            InitializeComponent();
            //
            if (!TrimApplication.HasBeenInitialized)
            {
                TrimApplication.Initialize();
            }
            TrimApplication.HasUserInterface = true;
            Stopwatch watch = new Stopwatch();
            
            using (Database db = new Database())
            {
                try
                {
                    watch.Start();
                   
                    //
                    db.Connect();
                    //
                     watch.Stop();
                    var clock = watch.Elapsed.Seconds.ToString();
                    //MessageBox.Show(string.Format("Finished in {0} seconds", clock));
                    //
                    label1.Text = "Connected to RM successfully as : " + db.CurrentUser.Name + " on " + db.Id + "/" + db.WorkgroupServerName + " - RM Workgroup connection time: " + clock + " seconds";
                    label1.Visible = true;
                }
                catch (Exception exp)
                {
                    label1.Text = "Error: " + exp;
                    return;
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(button1, "Find JSon file to be processed");
            lblJsonFile.Text = "";
            lblError.Text = "";
            label2.Text = "";
            
        }
        public enum LogTypes
        {
            err,
            log,
            dir
        }
        private void populatelistView(string p)
        {
            try
            {
                //dataGridView1.DataSource = null;
                //dataGridView1.Rows.Clear();

                //dataGridView1.Refresh();
                label2.Text = "";
               //List<Sheets> myMessage= JsonConvert.DeserializeObject<List<Sheets>>(File.ReadAllText(@p));

               var reader = new JsonTextReader(new StreamReader(@p, Encoding.GetEncoding(1251)));
               var serializer = new JsonSerializer { CheckAdditionalContent = false };
               List<Sheets> myMessage = (List<Sheets>)serializer.Deserialize(reader, typeof(List<Sheets>));

                //using (StreamWriter sw = new StreamWriter(@"c:\json.txt"))
                // using (JsonWriter writer = new JsonTextWriter(sw))
                //{
                //    myMessage = JsonConvert.DeserializeObject<List<Sheets>>(writer);
                //     //JsonConvert.de
                //    //serializer.Serialize(writer, product);
                //    // {"ExpiryDate":new Date(1230375600000),"Price":0}
                //}

                foreach (var file in myMessage)
                {                   
                    foreach (Rows r in file.Rows)
                    {
                        if (r.Folders != null)
                        {
                            string strDirIssues = null;
                            foreach (Folder f in r.Folders)
                            {                                                            
                                try
                                {
                                    if (f.MigrateLoc != null)
                                    {
                                        if (Directory.Exists(f.MigrateLoc))
                                        {
                                            ManageFiles mf = new ManageFiles();
                                            intMigratedfiles = 0;
                                            var txtFiles = Directory.EnumerateFiles(f.MigrateLoc, "*.*", SearchOption.TopDirectoryOnly);
                                            mf.sheetno = f.SheetLine;
                                            mf.Location = f.MigrateLoc;
                                            mf.SourceCnt = txtFiles.Count();
                                            mf.MigrateCnt = 0;
                                            mf.MigrateUri = Convert.ToInt64(f.RMuri);
                                            lstmf.Add(mf);
                                        }
                                        else
                                        {
                                            ManageFiles mf = new ManageFiles();
                                            mf.sheetno = f.SheetLine;
                                            mf.Location = f.MigrateLoc;
                                            mf.SourceCnt = 0;
                                            mf.MigrateCnt = 0;
                                            mf.Comments = "Not a valid directory";
                                            mf.MigrateUri = Convert.ToInt64(f.RMuri);
                                            lstmf.Add(mf);
                                            //Loggitt("Source directory issure", f.MigrateLoc + " - " + f.SheetLine.ToString(), LogTypes.dir);
                                            strDirIssues += "Not a valid directory("+f.MigrateLoc + ") - SS line No. " + f.SheetLine.ToString() + Environment.NewLine;
                                        }
                                    }
                                }
                                catch (Exception exp)
                                {
                                    ManageFiles mf = new ManageFiles();
                                    mf.sheetno = f.SheetLine;
                                    mf.Location = f.MigrateLoc;
                                    mf.SourceCnt = 0;
                                    mf.MigrateCnt = 0;
                                    mf.Comments = "Error: " + exp.Message.ToString();
                                    mf.MigrateUri = Convert.ToInt64(f.RMuri);
                                    lstmf.Add(mf);
                                    Loggitt("Error", exp.Message.ToString() + " " + f.MigrateLoc + " - " + f.SheetLine.ToString(), LogTypes.err);
                                }
                                if (f.SubFolders != null)
                                {
                                    foreach (Subfolder sf in f.SubFolders)
                                    {
                                        try
                                        {
                                            if (sf.MigrateLoc != null)
                                            {
                                                //Console.WriteLine("File Location: " + sf.MigrateLoc);
                                                if (Directory.Exists(sf.MigrateLoc))
                                                {
                                                    ManageFiles mf = new ManageFiles();
                                                    intMigratedfiles = 0;
                                                    var txtFiles = Directory.EnumerateFiles(sf.MigrateLoc, "*.*", SearchOption.TopDirectoryOnly);
                                                    mf.sheetno = sf.SheetLine;
                                                    mf.Location = sf.MigrateLoc;
                                                    mf.SourceCnt = txtFiles.Count();
                                                    mf.MigrateCnt = 0;
                                                    mf.MigrateUri = Convert.ToInt64(sf.RMuri);
                                                    lstmf.Add(mf);
                                                }
                                                else
                                                {
                                                    ManageFiles mf = new ManageFiles();
                                                    mf.sheetno = sf.SheetLine;
                                                    mf.Location = sf.MigrateLoc;
                                                    mf.SourceCnt = 0;
                                                    mf.MigrateCnt = 0;
                                                    mf.Comments = "Not a valid directory";
                                                    mf.MigrateUri = Convert.ToInt64(sf.RMuri);
                                                    lstmf.Add(mf);
                                                    //Loggitt("Source directory issure", sf.MigrateLoc+" - "+sf.SheetLine.ToString(), LogTypes.dir);
                                                    strDirIssues += "Not a valid directory(" + f.MigrateLoc + ") - SS line No. " + f.SheetLine.ToString() + Environment.NewLine;
                                                }
                                            }
                                        }
                                        catch (Exception exp)
                                        {
                                            ManageFiles mf = new ManageFiles();
                                            mf.sheetno = f.SheetLine;
                                            mf.Location = f.MigrateLoc;
                                            mf.SourceCnt = 0;
                                            mf.MigrateCnt = 0;
                                            mf.Comments = "Error: " + exp.Message.ToString();
                                            mf.MigrateUri = Convert.ToInt64(f.RMuri);
                                            lstmf.Add(mf);
                                            Loggitt("Error", exp.Message.ToString()+" "+sf.MigrateLoc + " - " + sf.SheetLine.ToString(), LogTypes.err);
                                        }
                                    }
                                }
                                
                            }
                            //build directory log here
                            Loggitt("Source directory issures", strDirIssues, LogTypes.dir);
                        }
                    }
                    
                }
                dataGridView1.DataSource = lstmf;
            }
            catch (Exception exp)
            {
                
                lblError.Text = exp.Message.ToString();
                Loggitt("Error", exp.Message.ToString(), LogTypes.err);
            }
        }
        //private void button2_Click(object sender, EventArgs e)
        //{
        //    using(Database TrimDb = new Database())
        //    {
        //   TrimMainObjectSearch baseSearch = new TrimMainObjectSearch(TrimDb, BaseObjectTypes.Classification);

                
        //        // to be replaced by saved search
        //        baseSearch.SetSearchString("Top");
        //        //baseSearch.SetSortString("id");
        //        //MessageBox.Show("Check baseSearch: " + baseSearch.FastCount, "", MessageBoxButtons.OK);
        //        //TrimMainObjectSearch resultSearch = ObjectSelector.SelectOne(this.Handle, baseSearch);  //(this.Handle, baseSearch);
        //        TrimMainObject tmo = ObjectSelector.SelectOne(this.Handle, baseSearch);
        //        if(tmo.TrimType==BaseObjectTypes.Classification)
        //        {
        //            Classification cl = (Classification)tmo;
        //            //lblEddieFolder.Text = cl.Title + "(" + cl.Uri.ToString() + ")";
        //        }
                
        //        //

        //    }
        //}        
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
                        case ".VCARD":
                        case ".vcard":

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
                        default:
                            //uri = 0;
                            return 0;
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
                    db.LogExternalEvent("Document location does not exist: " + FileLoc, BaseObjectTypes.Record, Convert.ToInt64(cont), true);
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message.ToString());
                db.LogExternalEvent("Error migrating: " + FileLoc + " Error: " + exp.Message.ToString(), BaseObjectTypes.Record, Convert.ToInt64(cont), true);
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
        public static void Loggitt(string title,string msg, LogTypes type)
        {
            //if (!(System.IO.Directory.Exists(@"C:\Temp\TrimMail\")))
            //{
            //    System.IO.Directory.CreateDirectory(@"C:\Temp\TrimMail\");
            //}
            string strLogPrefix = null;
            switch (type)
            {
                case LogTypes.log:
                    strLogPrefix = "errlog_";
                    break;
                case LogTypes.err:
                    strLogPrefix = "log_";
                    break;
                case LogTypes.dir:
                    strLogPrefix = "IncorrectDirlog__";
                    break;
            }
            //if (Error)
            //{
            //    strLogPrefix = "errlog_";
            //}
            //else
            //{
            //    strLogPrefix = "log_";
            //}
            
            FileStream fs = new FileStream(strWorkingLocation+ "\\" + strLogPrefix + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(strWorkingLocation + "\\" + strLogPrefix + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            s1.Write("Title: " + title + Environment.NewLine);
            s1.Write("Message: " + msg + Environment.NewLine);
            //s1.Write("StackTrace: " + stkTrace + Environment.NewLine);
            s1.Write("Date/Time: " + DateTime.Now.ToString() + Environment.NewLine);
            s1.Write
("============================================" + Environment.NewLine);
            s1.Close();
            fs1.Close();
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                lblJsonFile.Text = openFileDialog1.FileName;
                strWorkingLocation = Path.GetDirectoryName(openFileDialog1.FileName);
                populatelistView(openFileDialog1.FileName);
                
            }
        }
        private void button3_Click_1(object sender, EventArgs e)
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
                    var sdsd = lstmf.Where(x => x.SourceCnt > 0);
                    foreach (ManageFiles line in sdsd.Where(x=>x.sheetno >= Convert.ToInt32(tbProcessFrom.Text) && x.sheetno <= Convert.ToInt32(tbProcessto.Text)))
                    {
                        var txtFiles = Directory.EnumerateFiles(line.Location, "*.*", SearchOption.TopDirectoryOnly);
                        intMigratedfiles = 0;
                        foreach (var ff in txtFiles)
                        {
                            if (CreateNewRMRecord(ff, line.MigrateUri, db) > 0)
                            {
                                intMigratedfiles++;
                                label2.Text = "Migrating from Sheet line no " + line.sheetno.ToString() + " - file no. " + intMigratedfiles.ToString() + " of " + line.SourceCnt.ToString();
                                if (cbRemoveSource.Checked)
                                {
                                    try
                                    {
                                        File.Delete(ff);
                                        Loggitt("File migrated and deleted", "The file " + ff + " was migrated to RM and deleted.", LogTypes.log);
                                    }
                                    catch (Exception exp)
                                    {
                                        //Console.WriteLine("Error deleting File: " + exp.Message.ToString());
                                        lblError.Text = exp.Message.ToString();
                                        Loggitt("Delete File Error", exp.Message.ToString(), LogTypes.err);
                                    }
                                }
                                else
                                {
                                    Loggitt("File migrated and deleted", "The file " + ff + " was migrated to RM and NOT deleted.", LogTypes.log);
                                }
                            }
                        }
                        db.LogExternalEvent("There were " + line.SourceCnt.ToString() + " files in " + line.Location + " and " + intMigratedfiles.ToString() + " successfully migrated.", BaseObjectTypes.Record, line.MigrateUri, true);
                        Loggitt("Migrate records", "There were " + line.SourceCnt.ToString() + " files in " + line.Location + " and " + intMigratedfiles.ToString() + " successfully migrated.", LogTypes.log);
                        line.MigrateCnt = intMigratedfiles;
                        dataGridView1.DataSource = lstmf;
                        dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                        dataGridView1.Refresh();
                    }
                }
                catch (Exception exp)
                {
                    lblError.Text = exp.Message.ToString();
                    Loggitt("Migrate Error", exp.Message.ToString(), LogTypes.err);
                }

            }
            label2.Text = "Migration completed";

        }
        private void button4_Click(object sender, EventArgs e)
        {
            using (Database TrimDb = new Database())
            {
                TrimMainObjectSearch baseSearch = new TrimMainObjectSearch(TrimDb, BaseObjectTypes.Record);


                // to be replaced by saved search
                baseSearch.SetSearchString("uri:"+textBox1.Text);
                //baseSearch.SetSortString("id");
                //MessageBox.Show("Check baseSearch: " + baseSearch.FastCount, "", MessageBoxButtons.OK);
                //TrimMainObjectSearch resultSearch = ObjectSelector.SelectOne(this.Handle, baseSearch);  //(this.Handle, baseSearch);
                TrimMainObject tmo = ObjectSelector.SelectOne(this.Handle, baseSearch);
                if (tmo != null)
                {
                    if (tmo.TrimType == BaseObjectTypes.Record)
                    {
                        //Classification cl = (Classification)tmo;
                        Record r = (Record)tmo;
                        //label3.Text = cl.Title + "(" + cl.Uri.ToString() + ")";
                        label3.Text = r.Title + "(" + r.Number + ")";
                    }
                    else
                    {
                        MessageBox.Show("You need to select a folder");
                    }
                }
                else
                {
                    MessageBox.Show("Uri not found");
                }

                //

            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "Select a folder in which to save your workspace...";
            folderBrowserDialog1.SelectedPath = Application.StartupPath;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label4.Text = folderBrowserDialog1.SelectedPath;
            }
            //DialogResult result = openFileDialog1.ShowDialog();
            //folderBrowserDialog1 result = folderBrowserDialog1.ShowDialog();
            ////DialogResult result =Directory
            //if (result == DialogResult.OK) // Test result.
            //{
            //    label4.Text = openFileDialog1.FileName;
            //    strWorkingLocation = Path.GetDirectoryName(openFileDialog1.FileName);
            //    //populatelistView(openFileDialog1.FileName);

            //}
        }
    }
}
