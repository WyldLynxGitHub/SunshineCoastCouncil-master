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
                SqlConnection con = new SqlConnection("Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True");
                con.Open();
                Console.WriteLine("Enter SQL string");
                //string sqlstring = "Select [STD:DocumentSetID], [STD:version], [Volume:StorageLocation], [Volume:Filename] from [Extract Public Indexed] group by[STD:DocumentSetID], [Volume:StorageLocation], [Volume:Filename], [STD:version] Having Max([STD:version]) > 1 Order by[STD:DocumentSetID], [STD:version] OFFSET 75000 ROWS FETCH NEXT 3184 ROWS ONLY";
                //List<ECMMigration> lstecm = new List<ECMMigration>();
                List<newRevision> lstRev = new List<newRevision>();
                //
                using (SqlCommand command = new SqlCommand(Console.ReadLine(), con))
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
                con.Close();
                con.Dispose();
                Console.WriteLine("Sql query transfered to List, total to process: " + lstRev.Count().ToString());

                TrimApplication.Initialize();
                using (Database db = new Database())
                {
                    foreach (newRevision rr in lstRev)
                    {
                        string strStorageLoc = FindMappedStorage(rr.Location) + "\\" + rr.Filename; ;
                        //string strStorageLoc = @"D:\SCC\ECM\0BF\00FB8C37.001.pdf";
                        Record r = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, rr.DocSetID);
                        if (r != null)
                        {

                            try
                            {
                                InputDocument doc = new InputDocument();
                                doc.SetAsFile(strStorageLoc);
                                r.SetDocument(doc, true, false, "Add new ECM revision - migration");
                                r.Save();
                                Console.WriteLine("Added new revision " + rr.revision.ToString() + " to record " + rr.DocSetID);
                            }
                            catch (Exception exp)
                            {
                                Console.WriteLine("Record number " + rr.DocSetID + "  could not be added");
                                Loggitt("Record number " + rr.DocSetID + "  could not be added", "Revision load error");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Record number " + rr.DocSetID + "  could not be found");
                            Loggitt("Record number " + rr.DocSetID + "  could not be found", "Revision load error");
                        }
                    }

                }

            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: "+exp.Message.ToString());
                Loggitt(exp.Message.ToString(), "Revision load error");

            }
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
                    skjfdh = @"X:\ECMIncache2\INCACHE";
                    break;
                case "ECMIncache3":
                    skjfdh = @"X:\ECMIncache3\INCACHE";
                    break;
                case "ECMIncache4":
                    skjfdh = @"X:\ECMIncache4\INCACHE";
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
