using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace SCCImportLocations
{
    internal class BU1
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public List<BU2> bu2 { get; set; }
        public long RMUri { get; set; }
    }
    internal class BU2
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int BU1ID { get; set; }
        public List<BU3> bu3 { get; set; }
        public List<Position> position { get; set; }
        public long RMUri { get; set; }
    }
    internal class BU3
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int BU2ID { get; set; }
        public List<BU4> bu4 { get; set; }
        public List<Position> position { get; set; }
        public long RMUri { get; set; }
    }
    internal class BU4
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int BU3ID { get; set; }
        public List<Position> position { get; set; }
        public long RMUri { get; set; }
    }
    internal class Position
    {
        //Position_Number|Title|BusLevel01|BusLevel02|BusLevel03|BusLevel04|Supervisor_Number|FTE|Position_Status|Award|Classification
        public int Position_Number { get; set; }
        public string Title { get; set; }
        public int BusLevel01 { get; set; }
        public int BusLevel02 { get; set; }
        public int BusLevel03 { get; set; }
        public int BusLevel04 { get; set; }
        public List<User> user { get; set; }
        public long RMUri { get; set; }
    }
    internal class User
    {
        //Employee_Number|FirstName|PreferredName|Surname|Email|Position_Number
        public int Employee_Number { get; set; }
        public string FirstName { get; set; }
        public string PreferredName { get; set; }
        public string Surname { get; set; }
        public string Email { get; set; }
        public int Position_Number { get; set; }
        public long RMUri { get; set; }
    }
    internal class AD
    {
        public string email { get; set; }
        public string DomainLogon { get; set; }
    }
   
    class Program
    {
        public static List<AD> lstDomain;
        static void Main(string[] args)
        {
            //List<BU1> lstbu1 = BuildBU1();
            //List<BU2> lstbu2 = BuildBU2();
            //List<BU3> lstbu3 = BuildBU3();
            //List<BU4> lstbu4 = BuildBU4();
            List<Position> lstPosition = BuildPosition();
            List<User> lstUser = BuildUser();

            foreach (Position p in lstPosition)
            {
                //if (p.Position_Number == 003770)
                //{
                //    var sdsd = p.Title;
                //}
                p.user = lstUser.Where(x => x.Position_Number == p.Position_Number).ToList<User>();
            }

            //foreach (BU1 bu1 in lstbu1)
            //{
            //    bu1.bu2 = lstbu2.Where(x => x.BU1ID == bu1.ID).ToList<BU2>();
            //}
            //foreach (BU2 bu2 in lstbu2)
            //{
            //    bu2.bu3 = lstbu3.Where(x => x.BU2ID == bu2.ID).ToList<BU3>();
            //    bu2.position = lstPosition.Where(x => x.BusLevel02 == bu2.ID && x.BusLevel03==0 && x.BusLevel04==0).ToList<Position>();
            //}
            //foreach (BU3 bu3 in lstbu3)
            //{
            //    bu3.bu4 = lstbu4.Where(x => x.BU3ID == bu3.ID).ToList<BU4>();
            //    bu3.position = lstPosition.Where(x => x.BusLevel03 == bu3.ID && x.BusLevel04==0).ToList<Position>();
            //}
            //foreach(BU4 bu4 in lstbu4)
            //{
            //    //if(bu4.ID==43)
            //    //{
            //    //    var dsds = bu4.Name;
            //    //}
            //    bu4.position = lstPosition.Where(x => x.BusLevel04 == bu4.ID).ToList<Position>();
            //}
            //Count the number of positions 
            int iBU1 = lstPosition.Where(x => x.BusLevel01 == 0).Count();
            int iBU2 = lstPosition.Where(x => x.BusLevel02 == 0).Count();
            int iBU3 = lstPosition.Where(x => x.BusLevel03 == 0).Count();
            int iBU4 = lstPosition.Where(x => x.BusLevel04 == 0).Count();
            //lstPosition.ForEach(item => Console.WriteLine(item.Title+" "+item.Position_Number.ToString()+"/"+item.BusLevel01.ToString() + "/" + item.BusLevel02.ToString() + "/" + item.BusLevel03.ToString() + "/" + item.BusLevel04.ToString()));
            //lstUser.ForEach(item => Console.WriteLine(item.Position_Number.ToString() /*+ "/" + lstPosition.Where(x=>x.Position_Number==item.Position_Number).FirstOrDefault().Title*/ /*+ "/" + item.BusLevel02.ToString() + "/" + item.BusLevel03.ToString() + "/" + item.BusLevel04.ToString()*/));
            //lstPosition.ForEach(item => Console.WriteLine(item.Title + " " + item.user.Count().ToString() /*+ "/" + item.BusLevel01.ToString() + "/" + item.BusLevel02.ToString() + "/" + item.BusLevel03.ToString() + "/" + item.BusLevel04.ToString()*/));
            //lstPosition.Where(x=>x.BusLevel04>0).ForEach(item => Console.WriteLine(item.Title + " " + item.user.Count().ToString()));
            lstDomain = BuildADUpdate();
            BuildRMLocations(lstPosition);
        }
        static List<AD> BuildADUpdate()
        {
            List<AD> lstAd = new List<AD>();
            //
            string line;
            using (StreamReader file = new StreamReader(@"C:\temp\AD_Extract_20150922.txt"))
            {
                while ((line = file.ReadLine()) != null)
                {
                    AD bu1 = new AD();
                    string[] strArray = line.Split('\t');
                    int intArrayCount = strArray.Count();
                    int dsds = 1;
                    foreach (object obj in strArray)
                    {
                        if (dsds == 1)
                        {
                            bu1.DomainLogon = obj.ToString();
                            dsds = 2;
                        }
                        else if (dsds == 2)
                        {
                           // bu1.Name = obj.ToString();
                            dsds = 3;
                        }
                        else if (dsds == 3)
                        {
                            //bu1.email = obj.ToString();
                            dsds = 4;
                        }
                        else if (dsds == 4)
                        {
                            bu1.email = obj.ToString();
                            dsds = 1;
                        }
                    }
                    //bu1.RMUri = 0;
                    lstAd.Add(bu1);
                }
            }



            return lstAd;

        }
        static List<BU1> BuildBU1()
        {
            List<BU1> lstbu1 = new List<BU1>();
            string line;
            using (StreamReader file = new StreamReader(@"C:\temp\SCC\PageUpInitialLoadFiles\BusinessUnit01.txt"))
            {
                while ((line = file.ReadLine()) != null)
                {
                    BU1 bu1 = new BU1();
                    string[] strArray = line.Split('|');
                    int intArrayCount = strArray.Count();
                    int dsds = 1;
                    foreach (object obj in strArray)
                    {
                        if (dsds == 1)
                        {
                            bu1.ID = int.Parse(obj.ToString());
                            dsds = 2;
                        }
                        else
                        {
                            bu1.Name = obj.ToString();
                            dsds = 1;
                        }
                    }
                    bu1.RMUri = 0;
                    lstbu1.Add(bu1);
                }
            }
            return lstbu1;
        }
        static List<BU2> BuildBU2()
        {
            List<BU2> lstbu2 = new List<BU2>();
            using (System.IO.StreamReader filebu2 = new System.IO.StreamReader(@"C:\temp\SCC\PageUpInitialLoadFiles\BusinessUnit02.txt"))
            {
                string line;
                int intLineCount = 1;
                while ((line = filebu2.ReadLine()) != null)
                {
                    if (intLineCount > 1)
                    {
                        BU2 bu2 = new BU2();
                        string[] strArray = line.Split('|');
                        int intArrayCount = strArray.Count();
                        int dsds = 1;
                        foreach (object obj in strArray)
                        {
                            //if(obj[0])

                            if (dsds == 1)
                            {
                                bu2.ID = int.Parse(obj.ToString());
                                dsds = 2;
                            }
                            else if (dsds == 2)
                            {
                                bu2.Name = obj.ToString();
                                dsds = 3;
                            }
                            else if (dsds == 3)
                            {
                                bu2.BU1ID = int.Parse(obj.ToString());
                                dsds = 1;
                            }
                        }
                        bu2.RMUri = 0;
                        lstbu2.Add(bu2);
                    }
                    intLineCount = intLineCount + 1;
                }
            }

            return lstbu2;
        }
        static List<BU3> BuildBU3()
        {
            List<BU3> lstbu3= new List<BU3>();
            using (System.IO.StreamReader filebu2 = new System.IO.StreamReader(@"C:\temp\SCC\PageUpInitialLoadFiles\BusinessUnit03.txt"))
            {
                string line;
                int intLineCount = 1;
                while ((line = filebu2.ReadLine()) != null)
                {
                    if (intLineCount > 1)
                    {
                        BU3 bu3 = new BU3();
                        string[] strArray = line.Split('|');
                        int intArrayCount = strArray.Count();
                        int dsds = 1;
                        foreach (object obj in strArray)
                        {
                            //if(obj[0])

                            if (dsds == 1)
                            {
                                bu3.ID = int.Parse(obj.ToString());
                                dsds = 2;
                            }
                            else if (dsds == 2)
                            {
                                bu3.Name = obj.ToString();
                                dsds = 3;
                            }
                            else if (dsds == 3)
                            {
                                bu3.BU2ID = int.Parse(obj.ToString());
                                dsds = 1;
                            }
                        }
                        bu3.RMUri = 0;
                        lstbu3.Add(bu3);
                    }
                    intLineCount = intLineCount + 1;
                }
            }

            return lstbu3;

        }
        static List<BU4> BuildBU4()
        {
            List<BU4> lstbu4 = new List<BU4>();
            using (System.IO.StreamReader filebu2 = new System.IO.StreamReader(@"C:\temp\SCC\PageUpInitialLoadFiles\BusinessUnit04.txt"))
            {
                string line;
                int intLineCount = 1;
                while ((line = filebu2.ReadLine()) != null)
                {
                    if (intLineCount > 1)
                    {
                        BU4 bu4 = new BU4();
                        string[] strArray = line.Split('|');
                        int intArrayCount = strArray.Count();
                        int dsds = 1;
                        foreach (object obj in strArray)
                        {
                            //if(obj[0])

                            if (dsds == 1)
                            {
                                bu4.ID = int.Parse(obj.ToString());
                                dsds = 2;
                            }
                            else if (dsds == 2)
                            {
                                bu4.Name = obj.ToString();
                                dsds = 3;
                            }
                            else if (dsds == 3)
                            {
                                bu4.BU3ID = int.Parse(obj.ToString());
                                dsds = 1;
                            }
                        }
                        bu4.RMUri = 0;
                        lstbu4.Add(bu4);
                    }
                    intLineCount = intLineCount + 1;
                }
            }

            return lstbu4;

        }
        static List<Position> BuildPosition()
        {
            //Position_Number|Title|BusLevel01|BusLevel02|BusLevel03|BusLevel04|Supervisor_Number|FTE|Position_Status|Award|Classification
            //int|string|int|int|int|int
            List<Position> lstPosition = new List<Position>();
            using (System.IO.StreamReader filebu2 = new System.IO.StreamReader(@"C:\temp\SCC\PageUpInitialLoadFiles\Position.txt"))
            {
                string line;
                int intLineCount = 1;
                while ((line = filebu2.ReadLine()) != null)
                {
                    if (intLineCount > 1)
                    {
                        Position position = new Position();
                        string[] strArray = line.Split('|');
                        int intArrayCount = strArray.Count();
                        int dsds = 1;
                        foreach (object obj in strArray)
                        {
                            //if(obj[0])

                            if (dsds == 1)
                            {
                                position.Position_Number = int.Parse(obj.ToString());
                                dsds = 2;
                            }
                            else if (dsds == 2)
                            {
                                position.Title = obj.ToString();
                                dsds = 3;
                            }
                            else if (dsds == 3)
                            {
                                position.BusLevel01 = int.Parse(obj.ToString());
                                dsds = 4;
                            }
                            else if (dsds == 4)
                            {
                                position.BusLevel02 = int.Parse(obj.ToString());
                                dsds = 5;
                            }
                            else if (dsds == 5)
                            {
                                int temp;
                                if (int.TryParse(obj.ToString(), out temp))
                                {
                                    position.BusLevel03 = int.Parse(obj.ToString());
                                }
                                else
                                {
                                    position.BusLevel03 = 0;
                                }
                                dsds = 6;
                            }
                            else if (dsds == 6)
                            {
                                int temp;
                                if (int.TryParse(obj.ToString(), out temp))
                                {
                                    position.BusLevel04 = int.Parse(obj.ToString());
                                }
                                else
                                {
                                    position.BusLevel04 = 0;
                                }
                                dsds = 7;
                            }
                            else if (dsds == 7)
                            {
                                //position.BusLevel03 = int.Parse(obj.ToString());
                                dsds = 8;
                            }
                            else if (dsds == 8)
                            {
                                //position.BusLevel03 = int.Parse(obj.ToString());
                                dsds = 9;
                            }
                            else if (dsds == 9)
                            {
                                //position.BusLevel03 = int.Parse(obj.ToString());
                                dsds = 10;
                            }
                            else if (dsds == 10)
                            {
                                //position.BusLevel03 = int.Parse(obj.ToString());
                                dsds = 11;
                            }
                            else if (dsds == 11)
                            {
                                //position.BusLevel03 = int.Parse(obj.ToString());
                                dsds = 1;
                            } 
                        }
                        position.RMUri = 0;
                        lstPosition.Add(position);
                    }
                    intLineCount = intLineCount + 1;
                }
            }

            return lstPosition;
        }
        static List<User> BuildUser()
        {
            //Employee_Number|FirstName|PreferredName|Surname|Email|Position_Number
            List<User> lstUser = new List<User>();
            using (System.IO.StreamReader filebu2 = new System.IO.StreamReader(@"C:\temp\SCC\PageUpInitialLoadFiles\User.txt"))
            {
                string line;
                int intLineCount = 1;
                while ((line = filebu2.ReadLine()) != null)
                {
                    if (intLineCount > 1)
                    {
                        User user = new User();
                        string[] strArray = line.Split('|');
                        int intArrayCount = strArray.Count();
                        int dsds = 1;
                        foreach (object obj in strArray)
                        {
                            //if(obj[0])

                            if (dsds == 1)
                            {
                                user.Employee_Number = int.Parse(obj.ToString());
                                dsds = 2;
                            }
                            else if (dsds == 2)
                            {
                                user.FirstName = obj.ToString();
                                dsds = 3;
                            }
                            else if (dsds == 3)
                            {
                                user.PreferredName = obj.ToString();
                                dsds = 4;
                            }
                            else if (dsds == 4)
                            {
                                user.Surname = obj.ToString();
                                dsds = 5;
                            }
                            else if (dsds == 5)
                            {
                                user.Email = obj.ToString();
                                dsds = 6;
                            }
                            else if (dsds == 6)
                            {
                                user.Position_Number = int.Parse(obj.ToString());
                                dsds = 1;
                            }
                            //else if (dsds == 7)
                            //{
                            //    //position.BusLevel03 = int.Parse(obj.ToString());
                            //    dsds = 8;
                            //}
                            //else if (dsds == 8)
                            //{
                            //    //position.BusLevel03 = int.Parse(obj.ToString());
                            //    dsds = 9;
                            //}
                            //else if (dsds == 9)
                            //{
                            //    //position.BusLevel03 = int.Parse(obj.ToString());
                            //    dsds = 10;
                            //}
                            //else if (dsds == 10)
                            //{
                            //    //position.BusLevel03 = int.Parse(obj.ToString());
                            //    dsds = 11;
                            //}
                            //else if (dsds == 11)
                            //{
                            //    //position.BusLevel03 = int.Parse(obj.ToString());
                            //    dsds = 1;
                            //}
                        }
                        user.RMUri = 0;
                        lstUser.Add(user);
                    }
                    intLineCount = intLineCount + 1;
                }
            }

            return lstUser;

        }
        static void BuildRMLocations(List<Position> full)
        {
            TrimApplication.Initialize();

            try
            {
                using (Database db = new Database())
                {
                    //5503
                    Location locUseProfileof = new Location(db, 19509);
                    Location LocHighLevel = new Location(db, 5503);
                    db.Id = "03";
                    db.Connect();
                    FieldDefinition fd = new FieldDefinition(db, 501);
                    //
                    //Classification cla = new Classification(db);
                    
                    foreach (Position p in full)
                    {
                        //if (p.RMUri == 0)
                        //{
                            Location posloc = new Location(db, LocationType.Position);
                            //{
                            posloc.Surname = p.Title + " - " + p.Position_Number.ToString();
                            posloc.SetNotes("Entered by auto process: Source ID: " + p.Position_Number.ToString(), NotesUpdateType.AppendWithNewLine);
                            posloc.AddRelationship(LocHighLevel, LocRelationshipType.MemberOf, false);
                            //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                            posloc.IsWithin = true;
                            //posloc.UseProfileOf = locUseProfileof;
                            posloc.Save();
                            p.RMUri = posloc.Uri;
                            Console.WriteLine("Created " + p.Title + " - " + p.Position_Number.ToString());
                            //
                            foreach (User u in p.user.Where(x => x.RMUri == 0))
                            {
                                //if (u.RMUri == 0)
                                //{
                                    try
                                    {
                                        Location userloc = new Location(db, LocationType.Person);
                                        //{
                                        userloc.Surname = u.Surname;
                                        userloc.GivenNames = u.FirstName;
                                        userloc.EmailAddress = u.Email;
                                        userloc.CanLogin = true;
                                        userloc.LogsInAs = u.FirstName[0].ToString() + u.Surname[0].ToString() + u.Employee_Number.ToString();
                                        userloc.AdditionalLogin = u.Email;
                                        string dlogon = lstDomain.Where(x=>x.email==u.Email).First().DomainLogon;
                                        if(dlogon.Length>1)
                                        {
                                        userloc.LogsInAs = "MSC\\"+dlogon;
                                        }
                                        else
                                        {
                                            userloc.LogsInAs = "MSC\\";
                                        }
                                        //userloc.UserType = UserTypes.Contributor;

                                        userloc.SetNotes("Entered by auto process: Source ID: " + u.Employee_Number, NotesUpdateType.AppendWithNewLine);
                                        userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                                        //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                                        userloc.IsWithin = true;
                                        userloc.UseProfileOf = locUseProfileof;
                                        userloc.Save();
                                        u.RMUri = userloc.Uri;
                                        BuildRMpersonalFolders(userloc);
                                        Console.WriteLine("Created " + u.Surname);
                                        //}
                                    }
                                    catch (Exception exp)
                                    {
                                        Console.WriteLine("Error creating person location: " + exp.Message.ToString());
                                    }
                                //}
                                //else
                                //{
                                //    Location userloc = new Location(db, u.RMUri);
                                //    userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                                //    userloc.Save();
                                //}
                            }
                        //}
                        //else
                        //{
                        //    Location posloc = new Location(db, p.RMUri);
                        //    //posloc.AddRelationship(bu2loc, LocRelationshipType.MemberOf, false);
                        //    posloc.Save();
                        //}
                    }




                    //
                    //Build SCC org
                    //Location lOrg = new Location(db, LocationType.Organization);
                    //lOrg.Surname = "Sunshine Coast Council";
                    //lOrg.IsWithin = true;
                    //lOrg.SetFieldValue(fd, new UserFieldValue("Organisation"));
                    //lOrg.UseProfileOf = locUseProfileof;
                    //lOrg.Save();
                    //Console.WriteLine("Created " + lOrg.Surname);
                    //
                    //foreach (BU1 bu1 in full)
                    //{
                    //    Location bu1loc = new Location(db, LocationType.Organization);
                    //    //{
                    //        bu1loc.Surname = bu1.Name;
                    //        bu1loc.SetNotes("Entered by auto process: Source ID: " + bu1.ID.ToString(), NotesUpdateType.AppendWithNewLine);
                    //        bu1loc.SetFieldValue(fd, new UserFieldValue("Branch"));
                    //        bu1loc.AddRelationship(lOrg, LocRelationshipType.MemberOf, false);
                    //        bu1loc.IsWithin = true;
                    //        bu1loc.UseProfileOf = locUseProfileof;
                    //        bu1loc.Save();
                    //        bu1.RMUri = bu1loc.Uri;
                    //        Console.WriteLine("Created " + bu1.Name);

                    //        foreach (BU2 bu2 in bu1.bu2)
                    //        {
                    //            Location bu2loc = new Location(db, LocationType.Organization);
                    //            //{
                    //                bu2loc.Surname = bu2.Name;
                    //                bu2loc.SetNotes("Entered by auto process: Source ID: " + bu2.ID.ToString(), NotesUpdateType.AppendWithNewLine);
                    //                bu2loc.AddRelationship(bu1loc, LocRelationshipType.MemberOf, false);
                    //                bu2loc.SetFieldValue(fd, new UserFieldValue("Department"));
                    //                bu2loc.IsWithin = true;
                    //                bu2loc.UseProfileOf = locUseProfileof;
                    //                bu2loc.Save();
                    //                bu2.RMUri = bu2loc.Uri;
                    //                Console.WriteLine("Created " + bu2.Name);
                    //                //}
                    //            //
                    //                foreach (Position p in bu2.position)
                    //                {
                    //                    if (p.RMUri == 0)
                    //                    {
                    //                        Location posloc = new Location(db, LocationType.Position);
                    //                        //{
                    //                        posloc.Surname = p.Title + " - " + p.Position_Number.ToString();
                    //                        posloc.SetNotes("Entered by auto process: Source ID: " + p.Position_Number.ToString(), NotesUpdateType.AppendWithNewLine);
                    //                        posloc.AddRelationship(bu2loc, LocRelationshipType.MemberOf, false);
                    //                        //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                    //                        posloc.IsWithin = true;
                    //                        posloc.UseProfileOf = locUseProfileof;
                    //                        posloc.Save();
                    //                        p.RMUri = posloc.Uri;
                    //                        Console.WriteLine("Created " + p.Title + " - " + p.Position_Number.ToString());
                    //                        //
                    //                        foreach (User u in p.user.Where(x => x.RMUri == 0))
                    //                        {
                    //                            if (u.RMUri == 0)
                    //                            {
                    //                                try
                    //                                {
                    //                                    Location userloc = new Location(db, LocationType.Person);
                    //                                    //{
                    //                                    userloc.Surname = u.Surname;
                    //                                    userloc.GivenNames = u.FirstName;
                    //                                    userloc.EmailAddress = u.Email;
                    //                                    userloc.CanLogin = true;
                    //                                    userloc.LogsInAs = u.FirstName[0].ToString() + u.Surname[0].ToString() + u.Employee_Number.ToString();
                    //                                    userloc.AdditionalLogin = u.FirstName + "." + u.Surname + "@sunshinecoast.qld.gov.au";
                    //                                    userloc.UserType = UserTypes.Contributor;
                    //                                    userloc.SetNotes("Entered by auto process: Source ID: " + u.Employee_Number, NotesUpdateType.AppendWithNewLine);
                    //                                    userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                    //                                    //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                    //                                    userloc.IsWithin = true;
                    //                                    userloc.Save();
                    //                                    u.RMUri = userloc.Uri;
                    //                                    BuildRMpersonalFolders(userloc);
                    //                                    Console.WriteLine("Created " + u.Surname);
                    //                                    //}
                    //                                }
                    //                                catch (Exception exp)
                    //                                {
                    //                                    Console.WriteLine("Error creating person location: " + exp.Message.ToString());
                    //                                }
                    //                            }
                    //                            else
                    //                            {
                    //                                Location userloc = new Location(db, u.RMUri);
                    //                                userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                    //                                userloc.Save();
                    //                            }
                    //                        }
                    //                    }
                    //                    else
                    //                    {
                    //                        Location posloc = new Location(db, p.RMUri);
                    //                        posloc.AddRelationship(bu2loc, LocRelationshipType.MemberOf, false);
                    //                        posloc.Save();
                    //                    }
                    //                }
                    //            //
                    //                foreach (BU3 bu3 in bu2.bu3)
                    //                {
                    //                    Location bu3loc = new Location(db, LocationType.Organization);
                    //                    //{
                    //                        bu3loc.Surname = bu3.Name;
                    //                        bu3loc.SetNotes("Entered by auto process: Source ID: " + bu3.ID.ToString(), NotesUpdateType.AppendWithNewLine);
                    //                        bu3loc.AddRelationship(bu2loc, LocRelationshipType.MemberOf, false);
                    //                        bu3loc.SetFieldValue(fd, new UserFieldValue("Team"));
                    //                        bu3loc.IsWithin = true;
                    //                        bu3loc.UseProfileOf = locUseProfileof;
                    //                        bu3loc.Save();
                    //                        bu3.RMUri = bu3loc.Uri;
                    //                        Console.WriteLine("Created " + bu3.Name);
                    //                        //
                    //                        foreach (Position p in bu3.position)
                    //                        {
                    //                            if (p.RMUri == 0)
                    //                            {
                    //                                Location posloc = new Location(db, LocationType.Position);
                    //                                //{
                    //                                posloc.Surname = p.Title + " - " + p.Position_Number.ToString();
                    //                                posloc.SetNotes("Entered by auto process: Source ID: " + p.Position_Number.ToString(), NotesUpdateType.AppendWithNewLine);
                    //                                posloc.AddRelationship(bu3loc, LocRelationshipType.MemberOf, false);
                    //                                //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                    //                                posloc.IsWithin = true;
                    //                                posloc.UseProfileOf = locUseProfileof;
                    //                                posloc.Save();
                    //                                p.RMUri = posloc.Uri;
                    //                                Console.WriteLine("Created " + p.Title + " - " + p.Position_Number.ToString());
                    //                                //
                    //                                foreach (User u in p.user.Where(x => x.RMUri == 0))
                    //                                {
                    //                                    if (u.RMUri == 0)
                    //                                    {
                    //                                        try
                    //                                        {
                    //                                            Location userloc = new Location(db, LocationType.Person);
                    //                                            //{
                    //                                            userloc.Surname = u.Surname;
                    //                                            userloc.GivenNames = u.FirstName;
                    //                                            userloc.EmailAddress = u.Email;
                    //                                            //
                    //                                            userloc.CanLogin = true;
                    //                                            userloc.LogsInAs = u.FirstName[0].ToString() + u.Surname[0].ToString() + u.Employee_Number.ToString();
                    //                                            userloc.AdditionalLogin = u.FirstName + "." + u.Surname + "@sunshinecoast.qld.gov.au";
                    //                                            userloc.UserType = UserTypes.Contributor;

                    //                                            userloc.SetNotes("Entered by auto process: Source ID: " + u.Employee_Number, NotesUpdateType.AppendWithNewLine);
                    //                                            userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                    //                                            //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                    //                                            userloc.IsWithin = true;
                    //                                            userloc.Save();
                    //                                            u.RMUri = userloc.Uri;
                    //                                            BuildRMpersonalFolders(userloc);
                    //                                            Console.WriteLine("Created " + u.Surname);
                    //                                            //}
                    //                                        }
                    //                                        catch (Exception exp)
                    //                                        {
                    //                                            Console.WriteLine("Error creating person location: " + exp.Message.ToString());
                    //                                        }
                    //                                    }
                    //                                    else
                    //                                    {
                    //                                        Location userloc = new Location(db, u.RMUri);
                    //                                        userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                    //                                        userloc.Save();
                    //                                    }
                    //                                }
                    //                            }
                    //                            else 
                    //                            {
                    //                                Location posloc = new Location(db, p.RMUri);
                    //                                posloc.AddRelationship(bu3loc, LocRelationshipType.MemberOf, false);
                    //                                posloc.Save();
                    //                            }
                    //                            //}
                    //                        }
                    //                        //
                    //                        foreach (BU4 bu4 in bu3.bu4)
                    //                        {
                    //                            Location bu4loc = new Location(db, LocationType.Organization);
                    //                            //{
                    //                                bu4loc.Surname = bu4.Name;
                    //                                bu4loc.SetNotes("Entered by auto process: Source ID: " + bu4.ID.ToString(), NotesUpdateType.AppendWithNewLine);
                    //                                bu4loc.AddRelationship(bu3loc, LocRelationshipType.MemberOf, false);
                    //                                bu4loc.SetFieldValue(fd, new UserFieldValue("Unit"));
                    //                                bu4loc.IsWithin = true;
                    //                                bu4loc.UseProfileOf = locUseProfileof;
                    //                                bu4loc.Save();
                    //                                bu4.RMUri = bu4loc.Uri;
                    //                                Console.WriteLine("Created " + bu4.Name);
                    //                                //
                    //                                foreach (Position p in bu4.position)
                    //                                {
                    //                                    if (p.RMUri == 0)
                    //                                    {
                    //                                        Location posloc = new Location(db, LocationType.Position);
                    //                                        posloc.Surname = p.Title + " - " + p.Position_Number.ToString();
                    //                                        posloc.SetNotes("Entered by auto process: Source ID: " + p.Position_Number.ToString(), NotesUpdateType.AppendWithNewLine);
                    //                                        posloc.AddRelationship(bu4loc, LocRelationshipType.MemberOf, false);
                    //                                        //bu3loc.SetFieldValue(fd, new UserFieldValue("Business Unit 3"));
                    //                                        posloc.IsWithin = true;
                    //                                        posloc.UseProfileOf = locUseProfileof;
                    //                                        posloc.Save();
                    //                                        p.RMUri = posloc.Uri;
                    //                                        Console.WriteLine("Created " + p.Title + " - " + p.Position_Number.ToString());
                    //                                        foreach (User u in p.user)
                    //                                        {
                    //                                            if (u.RMUri == 0)
                    //                                            {
                    //                                                try
                    //                                                {
                    //                                                    Location userloc = new Location(db, LocationType.Person);
                    //                                                    userloc.Surname = u.Surname;
                    //                                                    userloc.GivenNames = u.FirstName;
                    //                                                    userloc.EmailAddress = u.Email;
                    //                                                    userloc.CanLogin = true;
                    //                                                    userloc.LogsInAs = u.FirstName[0].ToString() + u.Surname[0].ToString() + u.Employee_Number.ToString();
                    //                                                    userloc.AdditionalLogin = u.FirstName + "." + u.Surname + "@sunshinecoast.qld.gov.au";
                    //                                                    userloc.UserType = UserTypes.Contributor;
                    //                                                    userloc.SetNotes("Entered by auto process: Source ID: " + u.Employee_Number, NotesUpdateType.AppendWithNewLine);
                    //                                                    userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                    //                                                    userloc.IsWithin = true;
                    //                                                    userloc.Save();
                    //                                                    u.RMUri = userloc.Uri;
                    //                                                    BuildRMpersonalFolders(userloc);
                    //                                                    Console.WriteLine("Created " + userloc.Surname);
                    //                                                }
                    //                                                catch (Exception exp)
                    //                                                {
                    //                                                    Console.WriteLine("Error creating person location: " + exp.Message.ToString());
                    //                                                }
                    //                                            }
                    //                                            else
                    //                                            {
                    //                                                Location userloc = new Location(db, u.RMUri);
                    //                                                userloc.AddRelationship(posloc, LocRelationshipType.MemberOf, false);
                    //                                                userloc.Save();
                    //                                            }
                    //                                        }
                    //                                    }
                    //                                    else
                    //                                    {
                    //                                        Location posloc = new Location(db, p.RMUri);
                    //                                        posloc.AddRelationship(bu4loc, LocRelationshipType.MemberOf, false);
                    //                                        posloc.Save();
                    //                                    }
                    //                                }
                    //                        }
                    //                }
                    //        }
                    //}
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message.ToString());
                Loggitt(exp.Message.ToString(), "Error creating location");
            }
        }
        static void BuildRMpersonalFolders(Location loc)
        {
            TrimApplication.Initialize();
            try
            {
                using (Database db = new Database())
                {
                    db.Id = "03";
                    db.Connect();
                    RecordType rt = new RecordType(db, 7);
                    Record r = new Record(rt);
                    r.Client = loc;
                    r.Title = "Personal Workspace";
                    //r.Assignee = loc;
                    //r.HomeLocation = loc;
                    TrimAccessControlList ackk = r.AccessControlList;
                    ackk.SetPrivate((int)RecordAccess.ViewRecord, loc);
                    r.AccessControlList = ackk;
                    //Classification cl = new Classification(db, 501);
                    //r.Classification = cl;
                    //


                    //r.SecurityProfile.
                    r.Save();
                    //r.AddToFavorites();
                }

            }
            catch(Exception exp) {
                Loggitt(exp.Message.ToString(), "Error creating Personal workspace - "+loc.Uri.ToString());
            }
                        
        }
        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\LocationsImport\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\LocationsImport\");
            }
            FileStream fs = new FileStream(@"C:\Temp\LocationsImport\scc_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs, new UTF8Encoding(true));
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\LocationsImport\scc_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            //s1.Write("Title: " + title + Environment.NewLine);
            s1.Write(msg + Environment.NewLine);
            //s1.Write("StackTrace: " + stkTrace + Environment.NewLine);
            //s1.Write("Date/Time: " + DateTime.Now.ToString() + Environment.NewLine);
            //s1.Write("============================================" + Environment.NewLine);
            s1.Close();
            fs1.Close();
        }
        private void deserialise()
        {

        }

    }
}
