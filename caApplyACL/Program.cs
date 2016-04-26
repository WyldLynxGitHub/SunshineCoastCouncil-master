using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.Data.SqlClient;
using System.Data;

namespace caApplyACL
{
    class ACL
    {
        public string EcmClass { get; set; }
        public string RMlocation { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            List<ACL> lstAcl = new List<ACL>();
            using (SqlConnection con = new SqlConnection("Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True"))
            {
                con.Open();
                using (SqlCommand command = new SqlCommand("SELECT [ECMClass],[EddieAcl] FROM [StagingMove].[dbo].[SetAcl] where Run = 1", con))
                {
                    //DataTable dt = new DataTable();
                    //SqlDataAdapter da = new SqlDataAdapter();
                    //da.SelectCommand = command;
                    //da.Fill(dt);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ACL acl = new ACL();
                            acl.EcmClass = reader.GetString(0).Trim();
                            acl.RMlocation = reader.GetString(1).Trim();
                            lstAcl.Add(acl);
                        }
                    }
                }
            }

            Console.WriteLine("What is the origin uri?");
            int iOuri = Convert.ToInt32(Console.ReadLine());

            foreach (ACL acl in lstAcl)
            {
                try
                {

                    using (Database db = new Database())
                    {
                        //Console.WriteLine("Enter ECM class name");
                        TrimMainObjectSearch sObj = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                        //ECMSTDClassName:"Confidential Transition" and type:16
                        //string str = @""+Console.ReadLine()+"";
                        //sObj.SetSearchString("ECMSTDClassName:\"" + acl.EcmClass + "\" and type:16 and updated>9/03/2016");
                        sObj.SetSearchString("ECMSTDClassName:\"" + acl.EcmClass + "\" and type:16 and originRun:"+iOuri);
                        Console.WriteLine(sObj.SearchString);
                        Console.WriteLine("Record count: " + sObj.FastCount.ToString());
                        
                        if (sObj.FastCount > 0)
                        {
                            FieldDefinition fdDocDate = new FieldDefinition(db, 85);
                            LocationList lstLoc = new LocationList();
                            Location l = new Location(db, acl.RMlocation);
                            lstLoc.Add(l);
                            foreach (Record r in sObj)
                            {
                                if(r.DateRegistered<=r.DateReceived)
                                {
                                    DateTime tdtDateReceived = r.DateReceived;
                                    r.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                }
                                r.Save();
                                //TrimDateTime rmDt = (TrimDateTime)r.GetFieldValue(fdDocDate).ToDotNetObject();
                                
                                //r.DateReceived = rmDt;
                                //r.DateCreated = rmDt;
                                //r.Save();

                                //
                                TrimAccessControlList lstacl = r.AccessControlList;
                                lstacl.set_AccessLocations((int)RecordAccess.ViewRecord, lstLoc);
                                lstacl.set_AccessLocations((int)RecordAccess.ViewDocument, lstLoc);
                                r.AccessControlList = lstacl;
                                try
                                {
                                    r.Save();
                                    Console.WriteLine("ACL applied to " + r.Number);
                                }
                                catch (Exception exp)
                                {
                                    Console.WriteLine("ACL failed, Error: " + exp.Message.ToString());
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









            Console.ReadLine();
        }
    }
}
