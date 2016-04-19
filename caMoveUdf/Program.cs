using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace caMoveUdf
{
    class Program
    {
        static void Main(string[] args)
        {
            using (Database db = new Database())
            {
                TrimMainObjectSearch objs = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                objs.SetSearchString("type:16 and not ECMSTDJobNo:\"\"");
                //type:16 and not ECMSTDJobNo:""
                //ECMInfringementInfringementID:
                FieldDefinition fdFrom = new FieldDefinition(db, 91);
                FieldDefinition fdTo = new FieldDefinition(db, 510);
                foreach (Record r in objs)
                {
                    Console.WriteLine("Record: {0}", r.Uri.ToString());
                    string iFrom = (string)r.GetFieldValue(fdFrom).ToDotNetObject();
                    string iTo = (string)r.GetFieldValue(fdTo).ToDotNetObject();
                    if(string.IsNullOrWhiteSpace(iTo))
                    {
                        try
                        {
                            if (r.DateRegistered <= r.DateReceived)
                            {
                                //If the date registered and date recieved are in the wrong order RM gets cranky. we check and save if needed, this is a data vallidation issue from the migration.
                                DateTime tdtDateReceived = r.DateReceived;
                                r.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                r.Save();
                                Console.WriteLine("Record date registered updated");
                            }
                            r.SetFieldValue(fdTo, new UserFieldValue(Convert.ToString(iFrom)));
                            r.Save();
                            Console.WriteLine("Record UDF field updated");
                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine("Error: {0}",exp.Message.ToString());
                        }
                    }
                }

            }
        }
    }
}
