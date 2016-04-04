using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace caCheckDates
{
    class Program
    {
        
        static void Main(string[] args)
        {
            int iCount = 0;
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                Console.WriteLine("Change date recieved: Y/N");
                string dr = Console.ReadLine();       
                TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                objS.SetSearchString("type:16");
                foreach(Record r in objS)
                {
                    if (r.DateRegistered <= r.DateReceived)
                    {
                        iCount++;
                        Console.WriteLine("Record with a date issue: " + r.Number);
                        if (dr == "Y")
                        {
                            //if (r.Number == "805639")
                            //{
                                r.DateReceived = r.DateRegistered;
                                //DateTime tdtDateReceived = reccheck.DateReceived;
                                //reccheck.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                try
                                {
                                    r.Save();
                                    Console.WriteLine("Changed ok");
                                }
                                catch (Exception exp)
                                {
                                    Console.WriteLine("Error: " + exp.Message.ToString());
                                }
                            //}
                        }
                    }
                }
            }
            Console.WriteLine("Date issue count: " + iCount.ToString());
        }
    }
}
