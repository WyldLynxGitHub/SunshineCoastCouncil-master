using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace cafinalRun
{
    class Program
    {
        static string strConEddie = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
        static int iFound = 0;
        static void Main(string[] args)
        {
            
            string sds = "Select DocumentSetID, OrigBatch, Pass from ECM_MasterCheck where DocumentSetID not in (Select Cast([recordId] AS INT) from EddieRM)";
            using (DataTable dtMaster = new DataTable())
            {
                using (SqlConnection con = new SqlConnection(strConEddie))
                {
                    using (SqlCommand command = new SqlCommand(sds, con))
                    {
                        command.CommandTimeout = 999;
                        con.Open();
                        using (SqlDataAdapter da = new SqlDataAdapter())
                        {
                            da.SelectCommand = command;
                            da.Fill(dtMaster);
                        }
                    }
                }
                Console.WriteLine("Records to process: " + dtMaster.Rows.Count.ToString());
                Console.WriteLine("How many to bulk process?");
                int iProcess = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Which pass?");
                int iPass = Convert.ToInt32(Console.ReadLine());
                foreach (DataRow o in dtMaster.Select("Pass = " + iPass).Take(iProcess))
                {
                    string strBatch = o["OrigBatch"].ToString().Trim();
                    int iDocid = o.Field<int>("DocumentSetID");
                    //Console.WriteLine("Batch: "+ strBatch);
                    switch (strBatch)
                    {
                        case "Extract Public":
                            ProcessBatch("[Extract Public Indexed]", iDocid);
                            break;
                        case "Extract Public Zip":
                            ProcessBatch("[Extract Public Zip Indexed]", iDocid);
                            break;
                        case "Linked To Application":
                            ProcessBatch("Elai", iDocid);
                            break;
                        case "Linked To Property":
                            ProcessBatch("Elpi", iDocid);
                            break;
                        case "Extract Remainder":
                            ProcessBatch("Eri", iDocid);
                            break;

                    }
                }
                Console.WriteLine("Found documents: " + iFound.ToString());
                Console.ReadLine();
                //
                using (Database db = new Database())
                {
                    //db.Id = "45";
                    db.Connect();
                    //Console.WriteLine("Connected to DB: " + db.CurrentUser.Name);
                    bulkLoaderSample bls = new bulkLoaderSample();
                    //Console.ReadLine();
                    bls.run(db);
                }


            }
        }
        private static void ProcessBatch(string v, int i)
        {
            string strSql = "Select [Volume:Filename] from " + v + " Where ID in(SELECT MIN(ID) AS ID FROM " + v + " GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] = "+i;
            //Console.WriteLine("Sql: " + strSql);
            
            try
            {
                using (SqlConnection con = new SqlConnection(strConEddie))
                {
                    using (SqlCommand command = new SqlCommand(strSql, con))
                    {
                        con.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string strStorageLoc = FindMappedAllStorage(reader.GetString(0));
                                if (strStorageLoc != null)
                                {
                                    Console.WriteLine("Found file: " + strStorageLoc);
                                    iFound++;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message.ToString());
            }
        }
        private static string FindMappedAllStorage(string eCMVolumeStorageLocation)
        {
            //string dfd = @"\\scas57\ECMIncache18\";

            string skjfdh = null;

            if(File.Exists(@"T:\ECMIncache2\INCACHE\"+eCMVolumeStorageLocation))
            {
                FileInfo fi = new FileInfo(@"T:\ECMIncache2\INCACHE\" + eCMVolumeStorageLocation);
                skjfdh = fi.FullName + " " + fi.Length.ToString();
            }
            if(skjfdh==null)
            {
                if (File.Exists(@"T:\ECMIncache17\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"T:\ECMIncache17\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"T:\ECMIncache19\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"T:\ECMIncache19\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"U:\ECMIncache14\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"U:\ECMIncache14\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"U:\ECMIncache15\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"U:\ECMIncache15\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"O:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"O:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"W:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"W:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache3\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache3\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache4\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache4\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache5\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache5\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache6\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache7\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache7\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache8\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache10\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache10\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache11\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache11\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache12\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache12\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            if (skjfdh == null)
            {
                if (File.Exists(@"Y:\ECMIncache13\INCACHE\" + eCMVolumeStorageLocation))
                {
                    FileInfo fi = new FileInfo(@"Y:\ECMIncache13\INCACHE\" + eCMVolumeStorageLocation);
                    skjfdh = fi.FullName + " " + fi.Length.ToString();
                }
            }
            return skjfdh;


        }
    }
}
