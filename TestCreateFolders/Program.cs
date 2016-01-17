using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace TestCreateFolders
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //TrimApplication.Initialize();
                //Console.WriteLine("Software version: " + TrimApplication.SoftwareVersion);
                //using (Database db = new Database())
                //{
                //    db.Id = "ST";
                //    db.Connect();
                //    Console.WriteLine("Connected as: " + db.CurrentUser.Name);
                //}
                //Console.WriteLine("Test running console app");
                //TrimApplication.TrimBinariesLoadPath = "D:\\82\\x64\\Debug";
                TrimApplication.HasUserInterface = true;
                TrimApplication.Initialize();

                using (Database database = new Database())
                {
                    database.Id = "ST";
                    //database.WorkgroupServerName = "local";
                    database.AuthenticationMethod = ClientAuthenticationMechanism.Adfs;

                    database.Connect();


                    Console.WriteLine("{0} - {1}", database.CurrentUser.FullFormattedName, database.CurrentUser.Uri);
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message.ToString());
            }
//            using(Database db = new Database())
//            {
//                db.Id = "03";
//                db.Connect();
//                //Console.WriteLine("Username: " + db.CurrentUser.Name);
//                //2549
//                Classification clas = new Classification(db, "Corporate Services - Finance - Credit Management");
//                var u = clas.Uri;
//                Classification c1 = new Classification(db);
//                //c1.Name = clas.Name+" "+ "Projects";
//                c1.Name = "Projects";
//                //c1.SetNotes("Created automatically", NotesUpdateType.AppendWithUserStamp);
//                //c1 = clas.NewLowerLevel();

//                c1.Save();
//                //Classification t = new Classification(db, c1.Uri);
//                //t = clas.NewLowerLevel();
//                //t.Save();
//                //c1 = clas.NewLowerLevel();

//                //c1.Save();
////long dd = c1.Uri;
////Classification ff = new Classification(db, dd);
                
//                //clas.NewLowerLevel();
                
                
                

//            }
        }
    }
}
