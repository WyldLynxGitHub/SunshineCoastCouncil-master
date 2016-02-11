using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestMigrationProcessingOptions
{
    class Program
    {
        class ECMMigration
        {
            #region Properties  
            //0
            public string DocSetID { get; set; }
            //1
            public string ECMDescription { get; set; }
            //2
            public DateTime DocumentDate { get; set; }
            //3
            public string FileLocation { get; set; }
            //4
            public DateTime DeclaredDate { get; set; }
            //5
            public DateTime DateRegistered { get; set; }
            //6
            public string Bcs { get; set; }
            //7
            public string Folder { get; set; }
            //8
            public string Correspondant { get; set; }
            //9
            public string ExternalReference { get; set; }
            //10
            public int InternalReference { get; set; }
            //11
            public string SummaryText { get; set; }
            //12
            public DateTime DateReceived { get; set; }
            //13
            public string PropertyNo { get; set; }
            //14
            public string InfringementNo { get; set; }
            //15
            public string JobNo { get; set; }
            //16
            public string DocumentType { get; set; }
            //17
            public string ClassName { get; set; }
            //18
            public string BusinessCode { get; set; }

            #endregion
        }
        static void Main(string[] args)
        {
            RunSqlRead();
            //RunLinq();
        }

        private static void RunLinq()
        {
            SqlConnection con = new SqlConnection("Data Source=MSI-GS60;Initial Catalog=ECMNew;Integrated Security=True");
            //{
                con.Open();
                using (SqlCommand command = new SqlCommand("Select * from Query", con))
                {
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = command;
                    da.Fill(dt);
                con.Close();
                con.Dispose();
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    var list = (from tr in dt.AsEnumerable()
                                select new ECMMigration()
                                {
                                    DocSetID = tr.Field<string>("STD:DocumentSetID"),
                                   ECMDescription = tr.Field<string>("STD:DocumentDescription")
                                }).ToList();
                    foreach(ECMMigration h in list)
                    {
                        Console.WriteLine("Result: " + h.DocSetID);
                        //Console.WriteLine("{0}\t{1}", h.DocSetID, h.ECMDescription);
                    }
                    stopwatch.Stop();
                    TimeSpan ts = stopwatch.Elapsed;
                    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Console.WriteLine("Process Time " + elapsedTime);
            }

            //using (var command = new SqlCommand(queryString, connection))
            //{
            //    connection.Open();
            //    DataTable dt = new DataTable();
            //    SqlDataAdapter da = new SqlDataAdapter();
            //    da.SelectCommand = command;
            //    da.Fill(dt);
            //    return dt;
            //}
                //while (_sqlDataReader.Read())
                //{
                //    var employeeModel = new EmployeeModel
                //    {
                //        Eno = _sqlDataReader.GetInt32(_sqlDataReader.GetOrdinal("Empno")),
                //        Ename = _sqlDataReader.GetString(_sqlDataReader.GetOrdinal("Ename")),
                //        Job = _sqlDataReader.GetString(_sqlDataReader.GetOrdinal("Job")),
                //        Salary = _sqlDataReader.GetDecimal(_sqlDataReader.GetOrdinal("Sal"))
                //    };
                //    employeemodellist.Add(employeeModel);
                //}

                //EmployeeList = employeemodellist;
                //_sqlConnection.Close();
            }

        private static void RunSqlRead()
        {
            SqlConnection con = new SqlConnection("Data Source=MSI-GS60;Initial Catalog=ECMNew;Integrated Security=True");
                con.Open();
                using (SqlCommand command = new SqlCommand("Select * from Query", con))
                using (SqlDataReader reader = command.ExecuteReader())
                {

                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                List<ECMMigration> lstecm = new List<ECMMigration>();
                while (reader.Read())
                    {
                    //Console.WriteLine("{0}\t{1}", reader.GetInt32(0), reader.GetString(1));
                    ECMMigration ecm = new ECMMigration();
                    ecm.DocSetID = reader.GetInt32(0).ToString();
                    ecm.ECMDescription = reader.GetString(1);
                    lstecm.Add(ecm);
                }
                con.Close();
                con.Dispose();
                foreach (ECMMigration l in lstecm)
                {
                    Console.WriteLine(l.DocSetID + " - " + l.ECMDescription);
                }

                //SqlDataReader reader = command.ExecuteReader();
                //while (reader.Read())
                //{
                //    //Console.WriteLine(reader.GetString(1));
                //    ECMMigration ecm = new ECMMigration();
                //    ecm.DocSetID = reader.GetInt32(0).ToString();
                //    ecm.ECMDescription = reader.GetString(1);
                //    lstecm.Add(ecm);

                //}

                con.Close();
                con.Dispose();

                foreach (ECMMigration l in lstecm)
                {
                    Console.WriteLine(l.DocSetID + " - " + l.ECMDescription);
                }





                stopwatch.Stop();
                    TimeSpan ts = stopwatch.Elapsed;
                    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Console.WriteLine("Process Time " + elapsedTime+" - Query count: "+ lstecm.Count().ToString());
                }
                 
        }
    }
}
