
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace caMigrateRevisions
{
    class Program
    {
        private class newRevision
        {
            public string DocSetID { get; set; }
            public string Location { get; set; }
            public string Filename { get; set; }
            public Int16 revision { get; set; }
        }
        static void Main(string[] args)
        {
            //TrimApplication.Initialize();
            //using (Database db = new Database())
            //{
            //Origin m_origin = null;
            //BulkDataLoader m_loader = null;

            //m_origin = db.FindTrimObjectByName(BaseObjectTypes.Origin, "ECM bulk load") as Origin;

            //if (m_origin == null)
            //{
            //    m_origin = new Origin(db, OriginType.Custom1);

            //    m_origin.Name = "ECM bulk load";
            //    m_origin.OriginLocation = "SQL Extracts";
            //    m_origin.DefaultRecordType = db.FindTrimObjectByUri(BaseObjectTypes.RecordType, 2) as RecordType;
            //    m_origin.Save();
            //    Console.WriteLine("origin check, new: " + m_origin.Name);
            //}
            //else
            //{
            //    Console.WriteLine("origin check: " + m_origin.Name);
            //}
            //m_loader = new BulkDataLoader(m_origin);
            //if (!m_loader.Initialise())
            //{
            //    // this sample has no way of dealing with the error.
            //    Console.WriteLine(m_loader.ErrorMessage);
            //    return;
            //}
            //Console.WriteLine("Starting migration run ...");
            //m_loader.StartRun("Input Data", @"C:\Temp");
            ////
            //InputDocument doc = new InputDocument();
            //doc.SetAsFile(@"D:\SCC\ECM\0BF\00FB6A59.001.jpg");

            //bool updateRec = m_loader.UpdateDocument(366008, doc, BulkLoaderCopyMode.WindowsCopy);
            //m_loader.EndRun();
            //if(updateRec)
            //{
            //    m_loader.ProcessAccumulatedData();

            //    Int64 runHistoryUri = m_loader.RunHistoryUri;

            //}
            //}
            try
            {
                List<newRevision> lstRev = new List<newRevision>();
                using (SqlConnection con = new SqlConnection("Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True"))
                {
                    con.Open();
                    //Console.WriteLine("Connected to SQL 1025");
                    Console.WriteLine("Export Batch?");
                    string strBatch = Console.ReadLine();
                    Console.WriteLine("Offset?");
                    string strOffset = Console.ReadLine();
                    Console.WriteLine("Fetch?");
                    string strFetch = Console.ReadLine();
                    string strbatch = "[Extract Remainder DIFF]";
                    //Console.WriteLine("Enter SQL string");
                    //string sqlstring = "Select [STD:DocumentSetID], [STD:version], [Volume:StorageLocation], [Volume:Filename] from " + strBatch + " group by [STD:DocumentSetID], [Volume:StorageLocation], [Volume:Filename], [STD:version] Having Max([STD:version]) > 1 and [Volume:Filename] is not null Order by[STD:DocumentSetID] desc, [STD:version] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                    //string sqlstring = "SELECT dbo.Eri.[STD:DocumentSetID], dbo.Eri.[STD:Version], dbo.Eri.[Volume:StorageLocation], dbo.Eri.[Volume:Filename], M.MaxVersion, M.RmRevisions FROM from INNER JOIN dbo.ECM_MasterCheck AS M ON dbo.Eri.[STD:DocumentSetID] = M.DocumentSetID GROUP BY dbo.Eri.[STD: DocumentSetID], dbo.Eri.[Volume: StorageLocation], dbo.Eri.[Volume: Filename], dbo.Eri.[STD: Version], M.RmRevisions, M.MaxVersion HAVING (MAX(dbo.Eri.[STD: Version]) > 1) AND (M.MaxVersion > M.RmRevisions) ORDER BY dbo.Eri.[STD:DocumentSetID] DESC, dbo.Eri.[STD:Version] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                    //string sqlstring = "SELECT E.[STD:DocumentSetID], E.[STD:Version], E.[Volume:StorageLocation], E.[Volume:Filename] FROM " + strBatch + " AS E INNER JOIN dbo.ECM_MasterCheck AS M ON E.[STD:DocumentSetID] = M.DocumentSetID GROUP BY E.[STD:DocumentSetID], E.[Volume:StorageLocation], E.[Volume:Filename], E.[STD:Version], M.RmRevisions, M.MaxVersion, M.Pass HAVING (MAX(E.[STD:Version]) > 1) AND(M.MaxVersion > M.RmRevisions) AND M.Pass = 1 ORDER BY E.[STD:DocumentSetID] DESC, E.[STD:Version] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                    string sqlstring = "SELECT E.[STD:DocumentSetID], E.[STD:Version], E.[Volume:StorageLocation], E.[Volume:Filename] FROM " + strbatch + " AS E INNER JOIN dbo.ECM_MasterCheck AS M ON E.[STD:DocumentSetID] = M.DocumentSetID GROUP BY E.[STD:DocumentSetID], E.[Volume:StorageLocation], E.[Volume:Filename], E.[STD:Version], M.RmRevisions, M.MaxVersion, M.Pass, Duplicate HAVING (M.MaxVersion > M.RmRevisions) AND M.Pass = 2 and Duplicate = 0 ORDER BY E.[STD:DocumentSetID] DESC, E.[STD:Version]";
                    //List<ECMMigration> lstecm = new List<ECMMigration>();dgdg

                    //
                    using (SqlCommand command = new SqlCommand(sqlstring, con))
                    {
                        DataSet dataSet = new DataSet();

                        DataTable dt = new DataTable();
                        //DataTable dtt = extension.

                        SqlDataAdapter da = new SqlDataAdapter();
                        da.SelectCommand = command;
                        da.Fill(dt);
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            newRevision nr = new newRevision();
                            nr.DocSetID = reader.GetInt32(0).ToString();
                            nr.revision = reader.GetInt16(1);
                            nr.Location = reader.GetString(2);
                            nr.Filename = reader.GetString(3);
                            lstRev.Add(nr);
                        }
                    }
                }
                Console.WriteLine("Sql query transfered to List, total to process: " + lstRev.Count().ToString());
                int icount = 0;
                int iDocSource = 0;
                string strTotalList = lstRev.Count().ToString();
                TrimApplication.Initialize();
                //Record r = new Record(db)eq3
                foreach (newRevision rr in lstRev)
                {
                    icount++;
                    using (Database db = new Database())
                    {
                        //string strStorageLoc = FindMappedStorage(rr.Location) + "\\" + rr.Filename; ;
                        //string strStorageLoc2 = FindMappedStorageNas(rr.Location) + "\\" + rr.Filename;
                        ////string strStorageLoc = @"D:\SCC\ECM\0BF\00FB8C37.001.pdf";
                        //if (File.Exists(strStorageLoc))
                        //{
                        //    //Look for alternative
                        //    iDocSource = 1;
                        //}
                        //else if (File.Exists(strStorageLoc2))
                        //{
                        //    iDocSource = 2;
                        //}
                        //else
                        //{
                        //    iDocSource = 0;
                        //    Console.WriteLine("Document could not be found");
                        //    //Loggitt("Document  " + strStorageLoc + "  could not be found", "Revision load error");
                        //}
                        //if (iDocSource > 0)
                        //{
                        string strStorageLoc = FindMappedAllStorage(rr.Filename);
                        if (strStorageLoc != null)
                        {
                            Record r = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, rr.DocSetID);
                            if (r != null)
                            {
                                if (r.RevisionNumber < rr.revision)
                                {
                                    try
                                    {
                                        if (r.DateRegistered <= r.DateReceived)
                                        {
                                            DateTime tdtDateReceived = r.DateReceived;
                                            r.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                        }
                                        r.Save();

                                        InputDocument doc = new InputDocument();
                                        doc.SetAsFile(strStorageLoc);
                                        r.SetDocument(doc, true, false, "Add new ECM revision - migration");
                                        r.Save();
                                        Console.WriteLine("Added new revision " + rr.revision.ToString() + " to record " + rr.DocSetID + ", item " + icount.ToString() + " of " + strTotalList);
                                        deleteEcmSource(strStorageLoc);
                                        //Console.ReadLine();
                                    }
                                    catch (Exception exp)
                                    {
                                        Console.WriteLine("Record number " + rr.DocSetID + "  could not be added");
                                        Loggitt("Record number " + rr.DocSetID + "  could not be added - Error: " + exp.Message.ToString(), "Revision load error");
                                        //Console.ReadLine();

                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Record number " + rr.DocSetID + "  could not be found");
                                //Console.ReadLine();
                                // Loggitt("Record number " + rr.DocSetID + "  could not be found", "Revision load error");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Record " + rr.DocSetID + "  document could not be found");
                            //Console.ReadLine();
                        }
                    }

                }

            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: "+exp.Message.ToString());
                Loggitt(exp.Message.ToString(), "Revision load error");
                Console.ReadLine();

            }
        }
        private static string FindMappedAllStorage(string eCMVolumeStorageLocation)
        {
            //string dfd = @"\\scas57\ECMIncache18\";

            string skjfdh = null;

            if (File.Exists(@"L:\ECMIncache2\INCACHE\" + eCMVolumeStorageLocation))
            {
                FileInfo fi = new FileInfo(@"L:\ECMIncache2\INCACHE\" + eCMVolumeStorageLocation);
                skjfdh = fi.FullName;
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"L:\ECMIncache17\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache17\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache19\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache19\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache14\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache14\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache15\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache15\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache3\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache3\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache4\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache4\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache5\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache5\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache7\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache7\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache10\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache10\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache11\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache11\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache12\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache12\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache13\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache13\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName;
                }
            }
            return skjfdh;


        }
        private static string FindMappedStorage(string eCMVolumeStorageLocation)
        {
            string skjfdh = null;
            if (eCMVolumeStorageLocation.Length == 21)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 11);
            }
            else if (eCMVolumeStorageLocation.Length == 22)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 12);
            }

            switch (skjfdh)
            {
                case "ECMIncache2":
                    skjfdh = @"T:\ECMIncache2\INCACHE";
                    break;
                case "ECMIncache3":
                    skjfdh = @"Y:\ECMIncache3\INCACHE";
                    break;
                case "ECMIncache4":
                    skjfdh = @"Y:\ECMIncache4\INCACHE";
                    break;
                case "ECMIncache5":
                    skjfdh = @"Y:\ECMIncache5\INCACHE";
                    break;
                case "ECMIncache6":
                    skjfdh = @"Y:\ECMIncache6\INCACHE";
                    break;
                case "ECMIncache7":
                    skjfdh = @"Y:\ECMIncache7\INCACHE";
                    break;
                case "ECMIncache8":
                    skjfdh = @"Y:\ECMIncache8\INCACHE";
                    break;
                case "ECMIncache10":
                    skjfdh = @"W:\ECMIncache10\INCACHE";
                    break;
                case "ECMIncache11":
                    skjfdh = @"W:\ECMIncache11\INCACHE";
                    break;
                case "ECMIncache12":
                    skjfdh = @"V:\ECMIncache12\INCACHE";
                    break;
                case "ECMIncache13":
                    skjfdh = @"V:\ECMIncache13\INCACHE";
                    break;
                case "ECMIncache14":
                    skjfdh = @"U:\ECMIncache14\INCACHE";
                    break;
                case "ECMIncache15":
                    skjfdh = @"U:\ECMIncache15\INCACHE";
                    break;
                case "ECMIncache17":
                    skjfdh = @"T:\ECMIncache17\INCACHE";
                    break;
                case "ECMIncache18":
                    skjfdh = @"T:\ECMIncache18\INCACHE";
                    break;
            }


            return skjfdh;


        }
        private static string FindMappedStorageNas(string eCMVolumeStorageLocation)
        {
            //string dfd = @"\\scas57\ECMIncache18\";
            string skjfdh = null;
            if (eCMVolumeStorageLocation.Length == 21)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 11);
            }
            else if (eCMVolumeStorageLocation.Length == 22)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 12);
            }
            //var skjfdh = eCMVolumeStorageLocation.Substring(9, 11);
            switch (skjfdh)
            {

                case "ECMIncache2":
                    skjfdh = @"X:\ECM_export\ECMIncache2\Incache";
                    break;
                case "ECMIncache3":
                    skjfdh = @"X:\ECM_export\ECMIncache3\Incache";
                    break;
                case "ECMIncache4":
                    skjfdh = @"X:\ECM_export\ECMIncache4\Incache";
                    break;
                case "ECMIncache5":
                    skjfdh = @"X:\ECM_export\ECMIncache5\Incache";
                    break;
                case "ECMIncache6":
                    skjfdh = @"X:\ECM_export\ECMIncache6\Incache";
                    break;
                case "ECMIncache7":
                    skjfdh = @"X:\ECM_export\ECMIncache7\Incache";
                    break;
                case "ECMIncache8":
                    skjfdh = @"X:\ECM_export\ECMIncache8\INCACHE";
                    break;
                case "ECMIncache10":
                    skjfdh = @"X:\ECM_export\ECMIncache10\INCACHE";
                    break;
                case "ECMIncache11":
                    skjfdh = @"X:\ECM_export\ECMIncache11\INCACHE";
                    break;
                case "ECMIncache12":
                    skjfdh = @"X:\ECM_export\ECMIncache12\INCACHE";
                    break;
                case "ECMIncache13":
                    skjfdh = @"X:\ECM_export\ECMIncache13\INCACHE";
                    break;
                case "ECMIncache14":
                    skjfdh = @"X:\ECM_export\ECMIncache14\INCACHE";
                    break;
                case "ECMIncache15":
                    skjfdh = @"X:\ECM_export\ECMIncache15\INCACHE";
                    break;
                case "ECMIncache17":
                    skjfdh = @"X:\ECM_export\ECMIncache17\INCACHE";
                    break;
                case "ECMIncache18":
                    skjfdh = @"X:\ECM_export\ECMIncache18\INCACHE";
                    break;

            }


            return skjfdh;

        }
        private static void deleteEcmSource(string strStorageLoc)
        {
            try
            {
                if (File.Exists(strStorageLoc))
                {
                    File.Delete(strStorageLoc);
                    Console.WriteLine("ECM File deleted ok");
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error deleting ECm file: " + strStorageLoc);
            }
        }
        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\BulkLoader\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\BulkLoader\");
            }
            FileStream fs = new FileStream(@"C:\Temp\BulkLoader\BulkRevisions_log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\BulkLoader\BulkRevisions_log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
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
    }
}
