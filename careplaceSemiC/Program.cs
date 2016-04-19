using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace careplaceSemiC
{
    class Program
    {
        static void Main(string[] args)
        {
            using (Database db = new Database())
            {
                TrimMainObjectSearch objs = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                objs.SetSearchString("not ApplicationNumber:\"\"");
                //not RequestNumber:""
                //PropertyNumber:88 to 2446699
                //not ApplicationNumber:""
                //
                //503 app number string
                //504 Prop number number
                //505 request number string
                    //
                //type:16 and not ECMSTDJobNo:\"\"
                //type:16 and not ECMSTDJobNo:""
                //ECMInfringementInfringementID:
                FieldDefinition fdApplicationList = new FieldDefinition(db, 1002);
                FieldDefinition fdPropertyList = new FieldDefinition(db, 1003);
                FieldDefinition fdRequestNumber = new FieldDefinition(db, 505);
                FieldDefinition fdPropertyNumber = new FieldDefinition(db, 504);
                FieldDefinition fdApplicationNumber = new FieldDefinition(db, 503);
                Console.WriteLine("Count: " + objs.FastCount.ToString());
                foreach (Record r in objs)
                {
                    Console.WriteLine("Record: " + r.Uri.ToString());
                    //if (r.Uri == 496206)
                    //{
                        try
                        {
                            string sApplicationList = (string)r.GetFieldValue(fdPropertyList).ToDotNetObject();
                            Console.WriteLine("Record field: App list {0} ", sApplicationList);
                            int sRequestNumber = (int)r.GetFieldValue(fdPropertyNumber).ToDotNetObject();
                            Console.WriteLine("Record field: App list {0}, request {1} ", sApplicationList, sRequestNumber);
                            string sNewValue;
                            if (!string.IsNullOrEmpty(sApplicationList))
                            {
                                sNewValue = sApplicationList + " " + sRequestNumber.ToString();
                            }
                            else
                            {
                                sNewValue = sRequestNumber.ToString();
                            }
                            Console.WriteLine("Record field: App list {0}, request {1}, new value {2} ", sApplicationList, sRequestNumber, sNewValue);
                            r.SetFieldValue(fdApplicationList, new UserFieldValue(sNewValue));
                            
                            r.Save();
                            Console.WriteLine("Record updated");
                            //Console.ReadLine();
                        }
                        catch (Exception exp)
                        {
                            Console.WriteLine("Error: " + exp.Message.ToString());
                            //Console.ReadLine();
                        }
                    //}

                    //string sApplicationNumbert = (string)r.GetFieldValue(fdApplicationNumber).ToDotNetObject();

                    //if (sApplicationList == ";")
                    //{
                    //    r.SetFieldValue(fdApplicationList, new UserFieldValue());
                    //}
                    //string sPropertyList = (string)r.GetFieldValue(fdPropertyList).ToDotNetObject();
                    //if (sPropertyList == ";")
                    //{
                    //    rk
                    //    r.SetFieldValue(fdPropertyList, new UserFieldValue());
                    //}
                }
            }
        }
    }
}
