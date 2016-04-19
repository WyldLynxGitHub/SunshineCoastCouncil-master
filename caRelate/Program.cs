using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.Data.SqlClient;
using System.Data;

namespace caRelate
{
    class RMrelate
    {
        public int recordid1 { get; set; }
        public int recordid2 { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            TrimApplication.Initialize();
            List<RMrelate> lstRm = new List<RMrelate>();
            //string connectionString = "Data Source=MSI-GS60;Initial Catalog=SCC_Dev;Integrated Security=True";
            string connectionString = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            //string Sql = "SELECT R.recordId, R.uri, E.reVersions FROM dbo.TSRECORD AS R LEFT OUTER JOIN dbo.TSRECELEC AS E ON R.uri = E.uri WHERE (R.rcRecTypeUri = 16) and E.reVersions is null";
            //string Sql = "SELECT DocSetID, RelatedDocSetID FROM [Extract Related Docs - Public Class]";
            string Sql = "SELECT DocSetID, RelatedDocSetID FROM [Extract Related Docs - Public Class]";
            //WHERE(R.rcRecTypeUri = 21)"
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                //{
                con.Open();
                //Console.WriteLine("Enter SQL string");
                //List<ECMMigration> lstecm = new List<ECMMigration>();
                using (SqlCommand command = new SqlCommand(Sql, con))
                {
                    DataSet dataSet = new DataSet();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = command;
                    da.Fill(dt);
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        //Console.WriteLine("Record: " + reader.GetString(0));
                        RMrelate rm = new RMrelate();
                        rm.recordid1 = reader.GetInt32(0);
                        rm.recordid2 = reader.GetInt32(1);
                        lstRm.Add(rm);
                    }
                }
            }
            Console.WriteLine("Related row count: " + lstRm.Count.ToString());
            using (Database db = new Database())
            {
                foreach (RMrelate rm in lstRm)
                {
                    Console.WriteLine("Record one: " + rm.recordid1 + " Record two: " + rm.recordid2);
                    try
                    {
                        Record r1 = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, rm.recordid1.ToString());
                        if(r1!=null)
                        {
                            if (r1.DateRegistered <= r1.DateReceived)
                            {
                                DateTime tdtDateReceived = r1.DateReceived;
                                r1.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                r1.Save();
                            }
                        }
                        Record r2 = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, rm.recordid2.ToString());
                        if (r2 != null)
                        {
                            if (r2.DateRegistered <= r2.DateReceived)
                            {
                                DateTime tdtDateReceived = r2.DateReceived;
                                r2.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                r2.Save();
                            }
                        }
                        if (r1 != null && r2 != null)
                        {
                            //if (r1.ChildRelationships.Count == 0)
                            //{


                            r1.AttachRelationship(r2, RecordRelationshipType.IsRelatedTo);

                                //r1.RelatedRecord = r2;
                                r1.Save();
                                Console.WriteLine("Relationship Created ok: ");
                            //}
                        }
                    }
                    catch (Exception exp)
                    {
                        Console.WriteLine("Relationship failed: "+ exp.Message.ToString());
                    }

                    //Record r = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, rm.recordid2);


                        //UpdateMasterinRM(rm);
                }
            }
            Console.ReadLine();
        }
    }
}
