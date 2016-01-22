using HP.HPTRIM.SDK;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace caMigrateEcm
{
    public class bulkLoaderSample
    {

        Origin m_origin = null;
        BulkDataLoader m_loader = null;
        int m_recordCount = 0;
        String strErrorCol = null;
        //public void Dispose()
        //{
        //    if (m_loader != null)
        //    {
        //        m_loader.Dispose();
        //        m_loader = null;
        //    }
        //    if (m_origin != null)
        //    {
        //        m_origin.Dispose();
        //        m_origin = null;
        //    }
        //}
        //bool getNextRecord(PropertyOrFieldValue[] fields)
        //{
        //    // this function would normally be processing some input data source,
        //    // reading the next item, setting yp the fields array from this input.
        //    // For the example, we "hard code" some simple property values and
        //    // provide 5 records.
        //    m_recordCount++;
        //    if (m_recordCount > 25)
        //    {
        //        return false;
        //    }
        //    // a simple title
        //    String title = "Test Bulk Loaded Record Import #" + System.Convert.ToString(m_recordCount);
        //    title += ", imported from run number " + System.Convert.ToString(m_loader.RunHistoryUri);
        //    // some notes
        //    String notes = "Some notes for the bulk imported record #"
        //    + System.Convert.ToString(m_recordCount)
        //    + " data imported at: "
        //    + TrimDateTime.Now.ToLongDateTimeString();
        //    // a random date
        //    TrimDateTime dateCreated = new TrimDateTime(2007, 12, 14, 12, 0, 0);
        //    //// create a location "on the fly" for the Author property
        //    //PropertyOrFieldValue[] authorFields = new PropertyOrFieldValue[2];
        //    //authorFields[0] = new PropertyOrFieldValue(PropertyIds.LocationTypeOfLocation);
        //    //authorFields[1] = new PropertyOrFieldValue(PropertyIds.LocationSortName);
        //    //authorFields[0].SetValue(LocationType.Position);
        //    //authorFields[1].SetValue("Position #" + System.Convert.ToString(m_recordCount));
        //    ////12 12 HP TRIM Bulk Data Importing Programming Guide
        //    //// if we are rerunning, check to see if it already exists
        //    //PropertyValue positionName = new PropertyValue(PropertyIds.LocationSortName);
        //    //positionName.SetValue("Position #" + System.Convert.ToString(m_recordCount));
        //    //Int64 authorUri = m_loader.FindLocation(positionName);
        //    //if (authorUri == 0)
        //    //{
        //    //    // use the loader to setup the new location's properties
        //    //    Location authorLoc = m_loader.NewLocation();
                
        //    //        m_loader.SetProperties(authorLoc, authorFields);
        //    //        // now submit the location to the bulk loader queue
        //    //        authorUri = m_loader.SubmitLocation(authorLoc);
                
        //    //}
        //    // now set up the fields array based on these values
        //    fields[0].SetValue(title);
        //    fields[1].SetValue(dateCreated);
        //    fields[2].SetValue(notes);
        //    fields[3].SetValue(2001);
        //    fields[4].SetValue(DateTime.Now.ToShortDateString());
        //    //fields[3].SetValue(authorUri);
        //    return true;
        //}
        public bool run(Database db)
        {
            try
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                Console.WriteLine("Initialise BulkDataLoader ...");
                // create an origin to use for this sample. Look it up first just so you can rerun the code.
                m_origin = db.FindTrimObjectByName(BaseObjectTypes.Origin, "ECM bulk load") as Origin;

                //Console.WriteLine("origin check: " + m_origin.)
                //Console.ReadLine();
                if (m_origin == null)
                {
                    m_origin = new Origin(db, OriginType.Custom1);

                    m_origin.Name = "ECM bulk load";
                    m_origin.OriginLocation = "SQL Extracts";
                    // sample code assumes you have a record type defined called "Document"
                    m_origin.DefaultRecordType = db.FindTrimObjectByUri(BaseObjectTypes.RecordType, 2) as RecordType;
                    // don't bother with other origin defaults for the sample, just save it so we can use it
                    //m_origin.DefaultContainer = db.FindTrimObjectByUri(BaseObjectTypes.Record, 148020) as Record;

                    m_origin.Save();
                    Console.WriteLine("origin check, new: " + m_origin.Name);
                }
                else
                {
                    Console.WriteLine("origin check: " + m_origin.Name);
                }
                //Console.ReadLine();
                // construct a BulkDataLoader for this origin
                m_loader = new BulkDataLoader(m_origin);
                // initialise it as per instructions
                if (!m_loader.Initialise())
                {
                    // this sample has no way of dealing with the error.
                    Console.WriteLine(m_loader.ErrorMessage);
                    return false;
                }
                Console.WriteLine("Starting migration run ...");
                //Console.ReadLine();
                m_loader.StartRun("Input Data", @"C:\Temp");
                //13 13 HP TRIM Bulk Data Importing Programming Guide
                // setup the property array that will be used to transfer record metadata
                PropertyOrFieldValue[] recordFields = new PropertyOrFieldValue[11];
                recordFields[0] = new PropertyOrFieldValue(PropertyIds.RecordTitle);
                recordFields[1] = new PropertyOrFieldValue(PropertyIds.RecordDateCreated);
                recordFields[2] = new PropertyOrFieldValue(PropertyIds.RecordNotes);
                recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordContainer);
                recordFields[4] = new PropertyOrFieldValue(PropertyIds.RecordNumber);
                FieldDefinition fdDocumentType = new FieldDefinition(db, 517);
                recordFields[5] = new PropertyOrFieldValue(fdDocumentType);
                FieldDefinition fdInfringementNo = new FieldDefinition(db, 516);
                recordFields[6] = new PropertyOrFieldValue(fdInfringementNo);
                FieldDefinition fdInternalReference = new FieldDefinition(db, 515);
                recordFields[7] = new PropertyOrFieldValue(fdInternalReference);
                FieldDefinition fdJobNo = new FieldDefinition(db, 514);
                recordFields[8] = new PropertyOrFieldValue(fdJobNo);
                FieldDefinition fdPropertyNo = new FieldDefinition(db, 504);
                recordFields[9] = new PropertyOrFieldValue(fdPropertyNo);
                FieldDefinition fdSummaryText = new FieldDefinition(db, 513);
                recordFields[10] = new PropertyOrFieldValue(fdSummaryText);
                //SummaryText
                //PropertyNo
                //recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordAuthor);
                // now lets add some recordsfdInternalReference
                //

                //string connectionString = "Data Source=yfdev.cloudapp.net;Initial Catalog=SCC_ECM;Persist Security Info=True;User ID=EPL;Password=Password1!";
                //string connectionString = "Data Source=MSI-GS60;Initial Catalog=ECMTOne;Integrated Security=True";
                //Console.ReadLine();
                Console.WriteLine("Enter Connection string");
                using (SqlConnection con = new SqlConnection(Console.ReadLine()))
                {
                    con.Open();
                    Console.WriteLine("Enter SQL string");
                    string strloc = null;
                    using (SqlCommand command = new SqlCommand(Console.ReadLine(), con))
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        string strIndexLoc = null;
                        Console.WriteLine("Enter the network location of the index");
                        string strInloc = Console.ReadLine();
                        Console.WriteLine("Copy documents or remove");
                        BulkLoaderCopyMode windowsmode = BulkLoaderCopyMode.WindowsCopy;
                        switch (Console.ReadLine())
                        {
                            case "C":
                                windowsmode = BulkLoaderCopyMode.WindowsCopy;
                                break;
                            case "R":
                                windowsmode = BulkLoaderCopyMode.WindowsMove;
                                break;
                            default:
                                windowsmode = BulkLoaderCopyMode.WindowsCopy;
                                break;
                        }
                        //Console.ReadLine();
                        if (!Directory.Exists(strInloc))
                        {
                            strInloc = null;
                        }
                        //Console.WriteLine("read count: " + reader.FieldCount.ToString());


                        //int count = reader.FieldCount;
                        //while (reader.Read())
                        //{
                        //    for (int i = 0; i < count; i++)
                        //    {
                        //        Console.WriteLine(reader.GetValue(i));
                        //    }
                        //}

                        //Console.ReadLine();
                        while (reader.Read())
                        {
                            try
                            {
                                string loc = null;
                                if (strInloc != null)
                                {
                                    loc = strInloc + reader.GetString(3);
                                }

                                Record reccont = GetFolderUri(reader.GetString(6), reader.GetString(7), db);
                                if (reccont != null)
                                {
                                    Console.WriteLine("Importing ECM T1 record " + reader.GetInt32(0).ToString() + " ...");
                                    //recordFields[0]
                                    if (!reader.IsDBNull(0))
                                    {
                                        recordFields[0].SetValue(reader.GetString(1));
                                    }
                                    else
                                    {
                                        recordFields[0].SetValue("No description for title");
                                    }
                                    recordFields[1].SetValue((TrimDateTime)reader.GetDateTime(2));
                                    recordFields[2].SetValue("Migrated from EMS T1");
                                    recordFields[3].SetValue(reccont.Uri);
                                    recordFields[4].SetValue(reader.GetInt32(0).ToString()); //Record number
                                    if (reader.GetString(16) != null) { recordFields[5].SetValue(reader.GetString(16)); } else { recordFields[4].SetValue("No Document type"); }
                                    recordFields[6].SetValue(reader[14].ToString()); //InfringementNo
                                    recordFields[7].SetValue(reader[10].ToString()); //InternalReference
                                    recordFields[8].SetValue(reader[15].ToString()); //JobNo
                                    recordFields[9].SetValue(reader[13].ToString()); //PropertyNo
                                    recordFields[10].SetValue(reader[11].ToString()); //SummaryText

                                    Record importRec = m_loader.NewRecord();
                                    if (File.Exists(loc))
                                    {
                                        InputDocument doc = new InputDocument();
                                        doc.SetAsFile(loc);

                                        m_loader.SetDocument(importRec, doc, windowsmode);
                                    }
                                    m_loader.SetProperties(importRec, recordFields);
                                    m_loader.SubmitRecord(importRec);
                                    if (m_loader.Error.Bad == true)
                                    {
                                        Console.WriteLine("Error -  " + m_loader.Error.Message + " " + m_loader.Error.Source);
                                        strErrorCol = strErrorCol + "Loader Error: " + reader.GetInt32(0).ToString()+", " + " reader.GetInt32(0).ToString(): " + reader.GetInt32(0).ToString() + " reader.GetString(1): " + reader.GetString(1) + " reader.GetString(6): " + reader.GetString(6) + " reader.GetString(7)" + reader.GetString(7) + " reccont.Uri:"+ reccont.Uri.ToString()+   " - " + m_loader.Error.Message + Environment.NewLine;
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Could no find folder: " + reader.GetString(6), reader.GetString(7));
                                    strErrorCol = strErrorCol + "Find folder Error: " + reader.GetInt32(0).ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                                }
                            }
                            catch (Exception exp)
                            {
                                Loggitt("ReadLin Error: " + exp.Message.ToString(), "Error"+ " reader.GetInt32(0).ToString(): "+ reader.GetInt32(0).ToString()+ " reader.GetString(1): "+ reader.GetString(1)+ " reader.GetString(6): " + reader.GetString(6) + " reader.GetString(7)"+ reader.GetString(7));
                            }
                        }
                        //}
                    }
                }

                Console.WriteLine("Processing import batch ...");
                m_loader.ProcessAccumulatedData();
                Int64 runHistoryUri = m_loader.RunHistoryUri;
                // we're done, lets end the run now
                Console.WriteLine("Processing complete ...");
                m_loader.EndRun();
                stopwatch.Stop();

                TimeSpan ts = stopwatch.Elapsed;

                // Format and display the TimeSpan value.
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                Console.WriteLine("Process Time " + elapsedTime);
                var sdsd = m_loader.ErrorMessage;
                // just for interest, lets look at the origin history object and output what it did
                OriginHistory hist = new OriginHistory(db, runHistoryUri);
                Console.WriteLine("Number of records created ..... " + System.Convert.ToString(hist.RecordsCreated));
                //Console.WriteLine("Number of locations created ... "
                //+ System.Convert.ToString(hist.LocationsCreated));
                Loggitt("Bulk loading run:" + Environment.NewLine + "Number of records created: " + hist.RecordsCreated.ToString() + Environment.NewLine + "Time to run: " + elapsedTime + Environment.NewLine + "Loader errors: " + sdsd + Environment.NewLine + "Records in error: : " + hist.RecordsInError.ToString() + Environment.NewLine + Environment.NewLine + strErrorCol, "BulkLoader Log");

            }
            catch (Exception exp)
            {
                Loggitt("Main Error: " + exp.Message.ToString(), "Error");
            }
            return true;
        }
        private static Record GetFolderUri(string p1, string p2, Database db)
        {
            Record r = null;
            //zz legacy T1 ECM - Business Folders - Information Management - Projects
            Classification c = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, "zz legacy T1 ECM - Business Folders - " + p1);
            if (c != null)
            {
                TrimMainObjectSearch mso = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                //string ss = "title:" + reader.GetString(7) + " and classification:zz Legacy T1 ECM - Business Folders - " + reader.GetString(6);
                mso.SetSearchString("classification:" + c.Uri);
                foreach (Record rec in mso)
                {
                    if (rec.Title == p2)
                    {
                        r = rec;
                    }
                }
            }
            return r;
        }
        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\BulkLoader\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\BulkLoader\");
            }
            FileStream fs = new FileStream(@"C:\Temp\BulkLoader\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\BulkLoader\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
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
