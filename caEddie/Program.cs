using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace caEddie
{
    class RM
    {
        public string recordid { get; set; }
        public long uri { get; set; }
        public int Versions { get; set; }
        public string Batch { get; set; }
        public string Storage { get; set; }
        public string FileName { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            List<RM> lstRm = new List<RM>();
            //string connectionString = "Data Source=MSI-GS60;Initial Catalog=SCC_Dev;Integrated Security=True";
            string connectionString = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            //string Sql = "SELECT R.recordId, R.uri, E.reVersions FROM dbo.TSRECORD AS R LEFT OUTER JOIN dbo.TSRECELEC AS E ON R.uri = E.uri WHERE (R.rcRecTypeUri = 16) and E.reVersions is null";
            string Sql = "SELECT top 5000 dbo.Elai.[STD:DocumentSetID], dbo.Elai.[Volume:StorageLocation], dbo.Elai.[Volume:Filename] FROM dbo.ECM_MasterCheck AS M INNER JOIN dbo.Elai ON M.DocumentSetID = dbo.Elai.[STD:DocumentSetID] WHERE (M.RMUri IS NOT NULL) AND (M.RmRevisions = 0) AND (M.Issue <> 1) AND Pass = 1 order by dbo.Elai.[STD:DocumentSetID] Desc";
            //WHERE(R.rcRecTypeUri = 21)"
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                //Console.WriteLine("Enter SQL string");
                //List<ECMMigration> lstecm = new List<ECMMigration>();
                using (SqlCommand command = new SqlCommand(Sql, con))
                {
                    command.CommandTimeout = 0;
                    //DataSet dataSet = new DataSet();
                    //DataTable dt = new DataTable();
                    //SqlDataAdapter da = new SqlDataAdapter();
                    //da.SelectCommand = command;
                    //da.Fill(dt);
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        //Console.WriteLine("Record: " + reader.GetString(0));
                        RM rm = new RM();
                        rm.recordid = reader.GetInt32(0).ToString();
                        rm.Storage = reader.GetString(1);
                        rm.FileName = reader.GetString(2);
                        lstRm.Add(rm);
                    }
                }
            }
            Console.WriteLine("Closed sql connection");
                using (Database db = new Database())
                {
                int iprocesscnt = 0;
                string iTotalcnt = lstRm.Count.ToString();
                foreach (RM rm in lstRm)
                    {
                    //Console.WriteLine("Record: " + rm.recordid);
                    iprocesscnt++;
                        //try
                        //{
                        Record rec = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, rm.recordid);
                        //if (reccheck == null)
                        //{
                        string strStorageLoc = FindMappedStorage(rm.Storage) + "\\" + rm.FileName;
                    OriginHistory oh = new OriginHistory(db, 34343);
                    
                        if (rec != null)
                        {
                            if (File.Exists(strStorageLoc))
                            {
                                try
                                {
                                    if (rec.DateRegistered <= rec.DateReceived)
                                    {
                                        DateTime tdtDateReceived = rec.DateReceived;
                                        rec.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                        rec.Save();
                                    }
                                    InputDocument doc = new InputDocument();
                                    doc.SetAsFile(strStorageLoc);
                                    rec.SetDocument(doc, false, false, "Imported from ECM");
                                    rec.Save();
                                    deleteEcmSource(strStorageLoc);

                                Console.WriteLine("Eddie record "+iprocesscnt+" of "+iTotalcnt.ToString()+": updated with file: " + rec.Number);
                                    UpdateMasterinRM(rec, 0);
                                }
                                catch (Exception exp)
                                {
                                    Console.WriteLine("Save record attachment error: " + exp.Message.ToString());
                                    //UpdateMasterNoRm(rm.recordid, 2);
                                    UpdateMasterinRM(rec, 2);
                                }
                            }
                            else
                            {
                                UpdateMasterinRM(rec, 1);
                            }
                        }
                        else
                        {
                            UpdateMasterNoRm(rm.recordid, 0);
                        }
                        //}


                        //        UpdateMasterinRM(rm);
                    }
                //}
                //using (var bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepNulls & SqlBulkCopyOptions.KeepIdentity))
                //{
                //    bulkCopy.BatchSize = CustomerList.Count;
                //    bulkCopy.DestinationTableName = "dbo.tCustomers";
                //    bulkCopy.ColumnMappings.Clear();
                //    bulkCopy.ColumnMappings.Add("CustomerID", "CustomerID");
                //    bulkCopy.ColumnMappings.Add("FirstName", "FirstName");
                //    bulkCopy.ColumnMappings.Add("LastName", "LastName");
                //    bulkCopy.ColumnMappings.Add("Address1", "Address1");
                //    bulkCopy.ColumnMappings.Add("Address2", "Address2");
                //    bulkCopy.WriteToServer(CustomerList);
                //}
                //BulkOperation 
                //SqlBulkCopy bc = new SqlBulkCopy(con);

                //var bulk = new BulkOperation(con);
                //bulk.BulkInsert(dt);
                //bulk.BulkUpdate(dt);
            }
            Console.WriteLine("Records: " + lstRm.Count().ToString());
            Console.ReadLine();
            //
            //   TrimApplication.Initialize();
            //using (Database db = new Database())
            //{
            //    db.Connect();

            //    TrimMainObjectSearch objs = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
            //    //objs.SetSearchString("type:21 and not electronic");
            //    objs.SelectAll();
            //    foreach (Record r in objs)
            //    {
            //        Console.WriteLine("Record: " + r.Number);
            //    }

            //}
        }

        private static string getbatch(string v)
        {
            string strBatch = null;
            string connectionString = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            //string Sql = "SELECT R.recordId, R.uri, E.reVersions FROM dbo.TSRECORD AS R LEFT OUTER JOIN dbo.TSRECELEC AS E ON R.uri = E.uri WHERE (R.rcRecTypeUri = 16) and E.reVersions is null";
            var sql = "SELECT OrigBatch FROM dbo.ECM_MasterCheck WHERE (DocumentSetID = " + v + ")";
            //WHERE(R.rcRecTypeUri = 21)"
            using (SqlConnection con1 = new SqlConnection(connectionString))
            {
                //{
                con1.Open();
                //Console.WriteLine("Enter SQL string");
                //List<ECMMigration> lstecm = new List<ECMMigration>();
                using (SqlCommand command = new SqlCommand(sql, con1))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        strBatch = reader.GetString(0);
                    }
                }
            }
            return strBatch;
        }

        private static void UpdateMasterinRM(Record r, int prob)
        {
            try
            {
                //if (bDetailMsg)
                //{
                //    Console.WriteLine("Starting UpdateMasterinRM");
                //    Console.ReadLine();
                //}
                long ruri = r.Uri;
                int rRev = r.RevisionNumber;

                //"Select DocumentSetID, Origbatch, ClassName, maxversion from ECM_MasterCheck"
                string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, RMUri=@rmuri, RmRevisions=@ver, Issue=@issue WHERE DocumentSetID=@docsetid";
                //if (bDetailMsg)
                //{
                //    Console.WriteLine("UpdateMasterinRM query: " + SQL);
                //    Console.ReadLine();
                //}
                using (SqlConnection con2 = new SqlConnection("Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True"))
                {
                    using (SqlCommand cmd = new SqlCommand(SQL, con2))
                    {
                        cmd.Parameters.AddWithValue("@docsetid", r.Number);
                        cmd.Parameters.AddWithValue("@now", DateTime.Now);
                        cmd.Parameters.AddWithValue("@rmuri", ruri);
                        cmd.Parameters.AddWithValue("@ver", rRev);
                        cmd.Parameters.AddWithValue("@issue", prob);
                        con2.Open();
                        cmd.ExecuteNonQuery();
                        //con2.Close();
                        Console.WriteLine("Master List updated OK: ");
                        //if (bDetailMsg)
                        //{
                        //    Console.WriteLine("Query executed ok:");
                        //    Console.ReadLine();
                        //}
                    }
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Query execution failed: Error " + exp.Message.ToString());
            }

        }
        private static void UpdateMasterNoRm(string recnum, int prob)
        {
            try
            {
                string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, Issue=@issue WHERE DocumentSetID=@docsetid";
                using (SqlConnection con = new SqlConnection("Data Source = Cl1025; Initial Catalog = StagingMove; Integrated Security = True"))
                {
                    using (SqlCommand cmd = new SqlCommand(SQL, con))
                    {
                        cmd.Parameters.AddWithValue("@docsetid", recnum);
                        cmd.Parameters.AddWithValue("@now", DateTime.Now);
                        cmd.Parameters.AddWithValue("@issue", prob);
                        con.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
                //con.Close();
            }
            catch (Exception exp)
            {
                Console.WriteLine("Query execution failed: Error " + exp.Message.ToString());
            }
        }
        private static string FindMappedStorage(string eCMVolumeStorageLocation)
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
    }
}
