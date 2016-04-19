using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.IO;

namespace caCheckFileLocation
{
    class Program
    {
        static void Main(string[] args)
        {
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                //updated:18/04/2016 07:30:00 to 18/04/2016 09:00:00 and electronic
                TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                objS.SetSearchString("updated:18/04/2016 07:30:00 to 18/04/2016 09:00:00 and electronic");
                foreach(Record rec in objS)
                {
                    string[] split = rec.DocumentDetails.Split(':');
                    string FilePath = split.Last().Replace("+", "\\");
                    //Console.WriteLine("Location: " + rec.EStore.Path + FilePath);

                    if (rec.EStore.Uri == 1)
                    {
                        if (File.Exists(rec.EStore.Path + FilePath))
                        {
                            //Console.WriteLine("doc exists in Repos1: " + rec.EStore.Path + FilePath);
                        }
                        else
                        {
                            Console.WriteLine("doc does not exist in Repos1: "+rec.Number + " - "  + rec.EStore.Path + FilePath);
                            ElectronicStore es3 = new ElectronicStore(db, 503);
                            if (File.Exists(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath))
                            {
                                Console.WriteLine("doc exists in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);

                                string[] strfilename = FilePath.Split('\\');

                                try
                                {
                                    //File.Copy(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath, rec.EStore.Path + FilePath);
                                    Console.WriteLine("Move doc here:");
                                }
                                catch (Exception exp)
                                {
                                    Console.WriteLine("Error: " + exp.Message.ToString());
                                }
                            }
                            else
                            {
                                Console.WriteLine("doc does not exist in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);
                            }


                            //Console.WriteLine("doc does not exist in Repos1: " + es1.Path + FilePath);
                        }
                    }
                }











                //Console.WriteLine("Enter record uri to check");
                //long lUri = Convert.ToInt64(Console.ReadLine());
                //Record rec = (Record)db.FindTrimObjectByUri(BaseObjectTypes.Record, lUri);
                //if(rec!=null)
                //{
                //    Console.WriteLine("Found Record: " + rec.Number);
                //    if(rec.IsElectronic)
                //    {
                //        Console.WriteLine("Record docstore location: " + rec.DocumentDetails);
                //        string[] split = rec.DocumentDetails.Split(':');
                //        string FilePath = split.Last().Replace("+", "\\");
                //        Console.WriteLine("Location: " + rec.EStore.Path + FilePath);

                //        if(rec.EStore.Uri==1)
                //        {
                //            if (File.Exists(rec.EStore.Path + FilePath))
                //            {
                //                Console.WriteLine("doc exists in Repos1: " + rec.EStore.Path + FilePath);
                //            }
                //            else
                //            {
                //                ElectronicStore es3 = new ElectronicStore(db, 503);
                //                if (File.Exists(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath))
                //                {
                //                    Console.WriteLine("doc exists in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);

                //                    string[] strfilename = FilePath.Split('\\');

                //                    try
                //                    {
                //                        File.Copy(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath, rec.EStore.Path + FilePath);
                //                    }
                //                    catch(Exception exp)
                //                    {
                //                        Console.WriteLine("Error: " + exp.Message.ToString());
                //                    }
                //                }
                //                else
                //                {
                //                    Console.WriteLine("doc does not exist in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);
                //                }


                //                //Console.WriteLine("doc does not exist in Repos1: " + es1.Path + FilePath);
                //            }
                //        }


                //        //ElectronicStore es3 = new ElectronicStore(db, 503);
                //        //if (File.Exists(@"\\C1\DocStore\REPOS3\SP\1\"+FilePath))
                //        //{
                //        //    Console.WriteLine("doc exists in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);
                //        //}
                //        //else
                //        //{
                //        //    Console.WriteLine("doc does not exist in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);
                //        //}
                //        //ElectronicStore es1 = new ElectronicStore(db, 1);
                //        //if (File.Exists(es1.Path + FilePath))
                //        //{
                //        //    Console.WriteLine("doc exists in Repos1: " + es1.Path + FilePath);
                //        //}
                //        //else
                //        //{
                //        //    Console.WriteLine("doc does not exist in Repos1: " + es1.Path + FilePath);
                //        //}
                //    }
                //    else
                //    {
                //        Console.WriteLine("Record has no elec doc");
                //    }
                //}
                //else
                //{
                //    Console.WriteLine("Record not found");
                //}
            }
            Console.ReadLine();
        }
    }
}
