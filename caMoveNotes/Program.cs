using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace caMoveNotes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (Database db = new Database())
            {
                TrimMainObjectSearch objs = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                //792378
                //First check to see if anything comes back

                //With more time this search should have been defined into manageable chunks
                objs.SetSearchString("type:16"); //Limit to the record type that I know is relevent

                //Initial attempts to manage the chunk runs

                //Console.WriteLine("Enter row limit");
                //objs.LimitOnRowsReturned = Convert.ToInt64(Console.ReadLine()); // limit to 1000 out of 3.6 mill
                //Console.WriteLine("Enter row skip");
                //objs.SkipCount = Convert.ToInt64(Console.ReadLine());
                //objs.SetSearchString("number:792378");

                //Changed to grab the entire set as an array of record uri's, so I can manage this in chunks by using a skip and take. Tried the RM Limitonrowsreturns and skipcount but I wasnt confident that it was working properly. Not enough time to invetigate and confirm. Each record type return is 3.6 mill records.
                long[] auri = objs.GetResultAsUriArray();
                
                //db.FindTrimObjectByUri(BaseObjectTypes.fi)
                FieldDefinition fdNotes = new FieldDefinition(db, 42);
                var iCount = 0;
                var iTotal = objs.FastCount;
                //Im changing this for each compile and moving it to a prod WG server to run, doing 4 at a time which takes about 3 hours. 1 mill over 3 hours.
                foreach(var lUri in auri.Skip(1826354).Take(23646))
                {
                    //find the record from the looped uri
                    try
                    {
                        Record rec = (Record)db.FindTrimObjectByUri(BaseObjectTypes.Record, lUri);
                        if (rec != null)
                        {
                            iCount++;
                            //Grab notes from the user defined field
                            var sNotes = rec.GetFieldValue(fdNotes).ToString();
                            //grab the normal record system notes
                            var sExistingNotes = rec.Notes;
                            //Check for and exclude the custom notes field if empty, it would be nice if I could do this search as part of the original setsearchstring as about 50% of the 3.6 are empty but have a related db row.
                            if (!string.IsNullOrWhiteSpace(sNotes))
                            {
                                //Display something that maks me feel like its actually working
                                Console.WriteLine("Processing {0} - {1} of 23646", rec.Number, iCount.ToString());
                                //Ive had a couple of runs at this using different methods and NO SORTING so there is the possibility that the notes have already been copied.
                                if (!sExistingNotes.Contains(sNotes))
                                {
                                    if (rec.DateRegistered <= rec.DateReceived)
                                    {
                                        //If the date registered and date recieved are in the wrong order RM gets cranky. we check and save if needed, this is a data vallidation issue from the migration.
                                        DateTime tdtDateReceived = rec.DateReceived;
                                        rec.DateRegistered = (TrimDateTime)tdtDateReceived.AddHours(1);
                                        rec.Save();
                                    }
                                    try
                                    {
                                        rec.Save();
                                        Console.WriteLine("Notes transfered from udf to record notes");
                                    }
                                    catch (Exception exp)
                                    {
                                        Console.WriteLine("Error saving notes. " + exp.Message.ToString());
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("User defined notes already in Record notes");
                                }

                            }
                            else
                            {
                                Console.WriteLine("User defined notes empty");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Uri {0} not found", lUri.ToString());
                        }
                    }
                    catch (Exception exp)
                    {
                        Console.WriteLine("Error: {0}", exp.Message.ToString());
                    }
                }
            }
        }
    }
}
