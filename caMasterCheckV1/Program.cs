using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace caMasterCheckV1
{
    class Program
    {
        static void Main(string[] args)
        {
            string strConEddie = "Data Source=Cl1025;Initial Catalog=SCRM_PROD;Integrated Security=True";
            //string strConStage = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            //string strConEddie = "Data Source=MSI-GS60;Initial Catalog=SCC_Dev;Integrated Security=True";
            string sds = "SELECT R.recordId, ISNULL(E.reVersions, 0 ), R.Uri as reVersions FROM dbo.TSRECORD AS R LEFT OUTER JOIN dbo.TSRECELEC AS E ON R.uri = E.uri WHERE (R.rcRecTypeUri = 16)";
            //Console.WriteLine("Enter to proceed");
            //Console.ReadLine();
            //List<ECM> lstEcm = new List<ECM>();
//            var results = (from m in dt.AsEnumerable() where m.Field<string>("Code") == "1-1-12" select m).FirstOrDefault();

            DataTable dtEddie = new DataTable();
            using (SqlConnection con = new SqlConnection(strConEddie))
            {
                //con.Open();
                Console.WriteLine("Enter to proceed - conn open: "+con.Database);
                //Console.ReadLine();
                using (SqlCommand command = new SqlCommand(sds, con))
                {
                    command.CommandTimeout = 999;
                    con.Open();
                    //DataSet dataSet = new DataSet();

                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        da.SelectCommand = command;
                        da.Fill(dtEddie);
                    }

                    //Console.WriteLine("Find greater then 1 versions: Rows " + dtEddie.AsEnume);
                    //Console.ReadLine();
                    //int results = (from m in dtEddie.AsEnumerable() where m.Field<int>("reVersions") >1 select m).Count();
                    //Console.WriteLine("Find greater then 1 versions: Rows " + results.ToString());
                    //Console.ReadLine();
                }
            }
            Console.WriteLine("Datatable complete: Rows " + dtEddie.Rows.Count.ToString());
            //
            string strConEcm = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            //string ecmsds = "Select DocumentSetID, maxversion from ECM_MasterCheck where RMUri is null order by DocumentSetID desc";
            Console.WriteLine("Try to bulk load");
            Console.ReadLine();

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strConEcm))
            {
                bulkCopy.DestinationTableName ="dbo.EddieRM";

                try
                {
                    // Write from the source to the destination.
                    bulkCopy.WriteToServer(dtEddie);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            dtEddie.Dispose();
            Console.WriteLine("Done");
            Console.ReadLine();



            //DataTable dtecm = new DataTable();
            //using (SqlConnection con = new SqlConnection(strConEcm))
            //{
            //    con.Open();
            //    Console.WriteLine("Enter to proceed - ECM conn open: " + con.Database);
            //    //Console.ReadLine();
            //    using (SqlCommand command = new SqlCommand(ecmsds, con))
            //    {

            //        DataSet dataSet = new DataSet();

            //        using (SqlDataAdapter da = new SqlDataAdapter())
            //        {
            //            da.SelectCommand = command;
            //            da.Fill(dtecm);
            //        }
            //    }
            //}
            //Console.WriteLine("Datatable complete: Rows " + dtecm.Rows.Count.ToString());
            //Console.WriteLine("Bulk Coly to SQL");
            //Console.ReadLine();


            //Console.ReadLine();
            //int iExists = 0;
            //int iDoesnotExist = 0;
            //foreach (DataRow dr in dtecm.Rows)
            //{
            //    DataRow drr = dtEddie.AsEnumerable().SingleOrDefault(r => r.Field<String>("recordId") == dr.Field<int>(0).ToString());
            //    if(drr !=null)
            //    {
            //        Console.WriteLine(dr.Field<int>(0).ToString() + " Exists");
            //        iExists++;
            //    }
            //    else
            //    {
            //        iDoesnotExist++;
            //    }
            //    //Console.WriteLine(dr.Field<string>(0)+" - "+ dr.Field<Int16>(1));
            //}
            //Console.WriteLine("Not created: " + iDoesnotExist.ToString() + " Already Created: " + iDoesnotExist.ToString());
            //dtEddie.Dispose();
        }
    }
}
