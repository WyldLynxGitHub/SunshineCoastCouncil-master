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
            string strmessage = null;
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                //updated:18/04/2016 07:30:00 to 18/04/2016 09:00:00 and electronic
                //TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                //objS.SetSearchString("updated:17/04/2016 05:30:00 to 21/04/2016 05:00:00 and electronic");
                //foreach (Record rec in objS)
                //{
                //    string FilePath1 = null;
                //    var f = rec.DocumentDetails;
                //    if (f.Contains("Finalized"))
                //    {
                //        FilePath1 = f.Substring(0, f.Length - 37);
                //    }
                //    else
                //    {
                //        FilePath1 = f;
                //    }
                //    string[] k = FilePath1.Split(':');
                //    string FilePath = k.Last().Replace("+", "\\").Trim();

                //    string strFound = @"\\C1\DocStore\REPOS3\SP\1\" + FilePath;
                //    if (rec.EStore.Uri == 1)
                //    {
                //        if (!File.Exists(rec.EStore.Path + FilePath))
                //        {
                //            Console.WriteLine("doc does not exist in Repos1: " + rec.Number + " - " + rec.EStore.Path + FilePath);
                //            strmessage = strmessage + "Doc does not exist in Repos1: " + rec.Number + "(" + rec.Uri.ToString() + ") - " + rec.EStore.Path + FilePath;
                //            ElectronicStore es3 = new ElectronicStore(db, 503);
                //            if (File.Exists(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath))
                //            {

                //                Console.WriteLine("doc exists in Repos3: " + strFound);
                //                strmessage = strmessage + ", Found doc in " + strFound;
                //                string[] strfilename = FilePath.Split('\\');

                //                try
                //                {
                //                    File.Copy(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath, rec.EStore.Path + FilePath);
                //                    Console.WriteLine("Successfully moved document to right location");


                //                    //File.Copy(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath, rec.EStore.Path + FilePath);
                //                    //Console.WriteLine("Move doc here:");
                //                    strmessage = strmessage + " - Successfully moved doc";

                //                }
                //                catch (Exception exp)
                //                {
                //                    Console.WriteLine("Error: " + exp.Message.ToString());
                //                    strmessage = strmessage + " - Doc could not be moved, error: "+exp.Message.ToString();
                //                }
                //            }
                //            else
                //            {
                //                Console.WriteLine("doc does not exist in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);
                //                strmessage = strmessage + ", Did not find doc in " + strFound;
                //            }


                //            //Console.WriteLine("doc does not exist in Repos1: " + es1.Path + FilePath);
                //            strmessage = strmessage + Environment.NewLine;
                //        }

                //    }
                //}

                //Loggitt(strmessage, "Missing Docs review");



                Console.WriteLine("Enter record uri to check");
                long lUri = Convert.ToInt64(Console.ReadLine());
                Record rec = (Record)db.FindTrimObjectByUri(BaseObjectTypes.Record, lUri);
                if (rec != null)
                {
                    Console.WriteLine("Found Record: " + rec.Number);
                    if (rec.IsElectronic)
                    {
                        Console.WriteLine("Record docstore location: " + rec.DocumentDetails);
                        string[] split = rec.DocumentDetails.Split(':');
                        string FilePath = split.Last().Replace("+", "\\");
                        Console.WriteLine("Location: " + rec.EStore.Path + FilePath);

                        //if (rec.EStore.Uri == 1)
                        //{
                            if (File.Exists(rec.EStore.Path + FilePath))
                            {
                                Console.WriteLine("doc exists in Repos1: " + rec.EStore.Path + FilePath);
                            }
                            else
                            {
                                ElectronicStore es3 = new ElectronicStore(db, 503);
                                if (File.Exists(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath))
                                {
                                    Console.WriteLine("doc exists in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);

                                    string[] strfilename = FilePath.Split('\\');

                                    try
                                    {
                                        File.Copy(@"\\C1\DocStore\REPOS3\SP\1\" + FilePath, rec.EStore.Path + FilePath);
                                        Console.WriteLine("Successfully moved document to right location");

                                    }
                                    catch (Exception exp)
                                    {
                                        Console.WriteLine("Doc copy Error: " + exp.Message.ToString());
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("doc does not exist in Repos3: " + @"\\C1\DocStore\REPOS3\SP\1\" + FilePath);
                                }


                                //Console.WriteLine("doc does not exist in Repos1: " + es1.Path + FilePath);
                            }
                        //}


                    }
                    else
                    {
                        Console.WriteLine("Record has no elec doc");
                    }
                }
                else
                {
                    Console.WriteLine("Record not found");
                }
            }
            Console.WriteLine("Finished, press any key to close.");
            Console.ReadLine();
        }
        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\MissingDocs\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\MissingDocs\");
            }
            FileStream fs = new FileStream(@"C:\Temp\MissingDocs\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\MissingDocs\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            s1.Write("Title: " + title + Environment.NewLine);
            s1.Write("Message: " + msg + Environment.NewLine);
            //s1.Write("StackTrace: " + stkTrace + Environment.NewLine);
            s1.Write("Date/Time: " + DateTime.Now.ToString() + Environment.NewLine);
            s1.Write
("============================================" + Environment.NewLine);
            s1.Close();
            fs1.Close();
        }
    }
}
