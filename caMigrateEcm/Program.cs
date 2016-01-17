using HP.HPTRIM.SDK;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace caMigrateEcm
{
    class Program
    {
        static void Main(string[] args)
        {
            //Migrate();
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                db.Id = "45";
                db.Connect();
                bulkLoaderSample bls = new bulkLoaderSample();
                bls.run(db);
            }
        }
        private static void Migrate()
        {
            string connectionString = "Data Source=yfdev.cloudapp.net;Initial Catalog=SCC_ECM;Persist Security Info=True;User ID=EPL;Password=Password1!";
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand command = new SqlCommand("SELECT top 25 * FROM View_1", con))
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    using (Database db = new Database())
                    {
                        db.Id = "03";
                        db.Connect();
                        while (reader.Read())
                        {
                            RecordType rt = new RecordType(db, 16);
                            Record rCont = GetFolderUri(reader.GetString(6), reader.GetString(7), db);
                            if (rCont != null)
                            {

                                Record r = new Record(rt);
                                r.Container = rCont;
                                r.LongNumber = reader.GetInt32(0).ToString();
                                r.Title = reader.GetString(1);
                                r.DateCreated = (TrimDateTime)reader.GetDateTime(2);
                                string loc = "D:\\" + reader.GetString(3);
                                if (File.Exists(loc))
                                {
                                    InputDocument id = new InputDocument(loc);
                                    r.SetDocument(id, false, false, "Migrated from ECM t1");
                                }

                                r.Save();
                                Console.WriteLine("Record Created: " + r.Number);
                            }
                            else
                            {
                                Console.WriteLine("Could not find folder");
                            }
                        }
                    }
                }
            }
        }
        private static Record GetFolderUri(string p1, string p2, Database db)
        {
            Record r = null;
            Classification c = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, "zz Legacy T1 ECM - Business Folders - " + p1);
            if(c!=null)
            {
                TrimMainObjectSearch mso = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                //string ss = "title:" + reader.GetString(7) + " and classification:zz Legacy T1 ECM - Business Folders - " + reader.GetString(6);
                mso.SetSearchString("classification:"+c.Uri);
                foreach(Record rec in mso)
                {
                    if(rec.Title==p2)
                    {
                        r = rec;
                    }
                }
            }
            return r;
        }
        
    }
}
