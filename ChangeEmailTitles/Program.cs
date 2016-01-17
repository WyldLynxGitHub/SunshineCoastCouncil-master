using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.IO;
using System.Net.Mail;
using Microsoft.Office.Interop;

namespace ChangeEmailTitles
{
    class Program
    {
        static string strFolderLoc = null;
        static void Main(string[] args)
        {
            Console.WriteLine("Select Folder for original email files");
            strFolderLoc = Console.ReadLine();
            var result = new List<string>();
            if(!Directory.Exists(strFolderLoc))
            {
                Console.WriteLine("Folder does not exist");
            }
            string[] extensions = { ".eml", ".EML", ".msg", ".MSG" };

            //foreach (Directory.EnumerateFiles(strFolderLoc, "*.*", SearchOption.TopDirectoryOnly).Where(s => extensions.Any(ext => ext == Path.GetExtension(s))))
            //{

            //}
            var txtFiles = Directory.EnumerateFiles(strFolderLoc, "*.*", SearchOption.TopDirectoryOnly).Where(s => extensions.Any(ext => ext == Path.GetExtension(s)));
            //foreach (string file in Directory.EnumerateFiles(strFolderLoc, "*.*", SearchOption.AllDirectories).Where(s => extensions.Any(ext => ext == Path.GetExtension(s))))
            //{
            //    result.Add(file);
            //    Console.WriteLine(file);
                
            //}
            foreach (var ff in txtFiles)
            {
                FileInfo f = new FileInfo(ff);
                //Console.WriteLine(f.Name + " - " + f.CreationTime.ToString());
                CreateRMRecord(f);
            }

        }

        private static void CreateRMRecord(FileInfo f)
        {

           // msg. //LoadMessage(@"C:\Docs\TestMail.eml");
            var strmessageId = findmessagid(f);
            //Console.WriteLine("message id: " + strmessageId);

            using (Database db = new Database())
            {
                db.Connect();
                TrimMainObjectSearch objs = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                //objs.SetSearchString("messageId = " + strmessageId);
                objs.SetSearchString("messageId= *1C2DFB0563FDEE43AB54DAF2A50508D24475B09E@SCEX13.msc.INTERNAL*");
                if(objs.FastCount>0)
                {
                    Console.WriteLine("Found a record with the same message ID");
                }
                else
                {
                    Console.WriteLine("Did not find a message id");
                }
                //RecordType rt = new RecordType(db,3);
                //Record r = new Record(rt);              
                //r.Container = new Record(db, 172693);
                //r.Title = f.Name;
                //r.Save();
                //Console.WriteLine("New RM email record: " + r.Number);
                //Record r = new Record(db, 172736);


            }
        }

        private static object findmessagid(FileInfo f)
        {
            string msgid = null;
            using(Database db = new Database())
            {
                RecordType rt = new RecordType(db, 3);
                Record r = new Record(rt);
                r.Container = new Record(db, 172739);
                r.Title = f.Name;
                r.SetDocument(f.FullName);
                r.Save();
                msgid = r.MessageId;
                r.Delete();
            }
            return msgid;
        }

    }
}
