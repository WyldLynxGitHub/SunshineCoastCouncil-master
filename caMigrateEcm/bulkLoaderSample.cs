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
                m_loader = new BulkDataLoader(m_origin);
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
                PropertyOrFieldValue[] recordFields = new PropertyOrFieldValue[20];
                recordFields[0] = new PropertyOrFieldValue(PropertyIds.RecordTitle);
                recordFields[1] = new PropertyOrFieldValue(PropertyIds.RecordDateCreated);
                recordFields[2] = new PropertyOrFieldValue(PropertyIds.RecordNotes);
                recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordContainer);
                recordFields[4] = new PropertyOrFieldValue(PropertyIds.RecordNumber);
                //
                FieldDefinition fdDocumentType = new FieldDefinition(db, 13);
                recordFields[5] = new PropertyOrFieldValue(fdDocumentType);
                //
                FieldDefinition fdInfringementNo = new FieldDefinition(db, 12);
                recordFields[6] = new PropertyOrFieldValue(fdInfringementNo);
                //
                FieldDefinition fdInternalReference = new FieldDefinition(db, 11);
                recordFields[7] = new PropertyOrFieldValue(fdInternalReference);
                //
                FieldDefinition fdJobNo = new FieldDefinition(db, 10);
                recordFields[8] = new PropertyOrFieldValue(fdJobNo);
                //
                FieldDefinition fdPropertyNo = new FieldDefinition(db, 504);
                recordFields[9] = new PropertyOrFieldValue(fdPropertyNo);
                //
                FieldDefinition fdSummaryText = new FieldDefinition(db, 9);
                recordFields[10] = new PropertyOrFieldValue(fdSummaryText);
                //New set
                FieldDefinition fdBusinessCode = new FieldDefinition(db, 20);
                recordFields[11] = new PropertyOrFieldValue(fdBusinessCode);
                //
                FieldDefinition fdClassName = new FieldDefinition(db, 14);
                recordFields[12] = new PropertyOrFieldValue(fdClassName);
                //
                FieldDefinition fdCorrespondant = new FieldDefinition(db, 21);
                recordFields[13] = new PropertyOrFieldValue(fdCorrespondant);
                //
                FieldDefinition fdDateReceived = new FieldDefinition(db, 19);
                recordFields[14] = new PropertyOrFieldValue(fdDateReceived);
                //
                FieldDefinition fdDateRegistered = new FieldDefinition(db, 18);
                recordFields[15] = new PropertyOrFieldValue(fdDateRegistered);
                //
                FieldDefinition fdDeclaredDate = new FieldDefinition(db, 17);
                recordFields[16] = new PropertyOrFieldValue(fdDeclaredDate);
                //
                FieldDefinition fdDocumentDate = new FieldDefinition(db, 16);
                recordFields[17] = new PropertyOrFieldValue(fdDocumentDate);
                //
                FieldDefinition fdExternalReference = new FieldDefinition(db, 15);
                recordFields[18] = new PropertyOrFieldValue(fdExternalReference);
                //
                FieldDefinition fdFileName = new FieldDefinition(db, 22);
                recordFields[19] = new PropertyOrFieldValue(fdFileName);

                //string connectionString = "Data Source=yfdev.cloudapp.net;Initial Catalog=SCC_ECM;Persist Security Info=True;User ID=EPL;Password=Password1!";
                //string connectionString = "Data Source=MSI-GS60;Initial Catalog=ECMNew;Integrated Security=True";
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
                        //Console.WriteLine("Query count: " + reader.FieldCount.ToString());
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
                        while (reader.Read())
                        {
                            try
                            {
                                //if (reader[0] != null)
                                //{
                                    string loc = null;
                                    if (strInloc != null)
                                    {
                                        loc = strInloc + reader.GetString(3);
                                    }

                                    Record reccont = GetFolderUri(reader.GetString(6), reader.GetString(7), db);
                                    if (reccont != null)
                                    {
                                        Console.WriteLine("Importing ECM T1 record " + reader.GetInt32(0).ToString() + " ...");
                                    string fTitle;
                                    if (reader.IsDBNull(1) == true)
                                    {
                                        fTitle = "Record description from ECM Export is Null";
                                        strErrorCol = strErrorCol + "Ecm description to Title issue: Description Null for record no " + reader.GetInt32(0).ToString() + Environment.NewLine;
                                    }
                                    else
                                    {
                                        if(reader.GetString(1).Length>253)
                                        {
                                            fTitle = reader.GetString(1).Substring(0, 253);
                                            strErrorCol = strErrorCol + "Ecm description to Title issue: Description greater then 254 ch, description truncated for record no " + reader.GetInt32(0).ToString() + Environment.NewLine;
                                        }
                                        else if(reader.GetString(1).Length<1)
                                        {
                                            fTitle = "Record description from ECM Export is Empty";
                                            strErrorCol = strErrorCol + "Ecm description to Title issue: Description blank for record no " + reader.GetInt32(0).ToString() + Environment.NewLine;
                                        }
                                        else
                                        {
                                            fTitle = reader.GetString(1);
                                        }
                                    }
                                        recordFields[0].SetValue(fTitle); //Title
                                        recordFields[1].SetValue((TrimDateTime)reader.GetDateTime(2));
                                        recordFields[2].SetValue("Migrated from EMS T1 using " + m_origin.Name);
                                        recordFields[3].SetValue(reccont.Uri); //Container
                                        recordFields[4].SetValue(reader.GetInt32(0).ToString()); //Record number

                                        //Potential null fileds
                                        if (reader.IsDBNull(16) == false) { recordFields[5].SetValue(reader.GetString(16)); } else { recordFields[5].SetValue("No Document type"); } //Document type
                                        if (reader.IsDBNull(14) == false) { recordFields[6].SetValue(reader[14].ToString()); } else { recordFields[6].ClearValue(); }//InfringementNo
                                    if (reader.IsDBNull(10) == false) { recordFields[7].SetValue(reader[10].ToString()); } else { recordFields[7].ClearValue(); } //InternalReference
                                    if (reader.IsDBNull(16) == false) { recordFields[8].SetValue(reader[15].ToString()); } else { recordFields[8].ClearValue(); }//JobNo
                                        if (reader.IsDBNull(13) == false) { recordFields[9].SetValue(reader[13].ToString()); } else { recordFields[9].ClearValue(); }//PropertyNo
                                    if (reader.IsDBNull(11) == false) { recordFields[10].SetValue(reader[11].ToString()); } else { recordFields[10].ClearValue(); }//SummaryText
                                                                                                                                                                  //
                                                                                                                                                                  //
                                    if (reader.IsDBNull(18) == false) { recordFields[11].SetValue(reader[18].ToString()); } else { recordFields[11].ClearValue(); } //BusinessCode
                                    if (reader.IsDBNull(17) == false) { recordFields[12].SetValue(reader[17].ToString()); } else { recordFields[12].ClearValue(); }//ClassName
                                    if (reader.IsDBNull(8) == false) { recordFields[13].SetValue(reader[8].ToString()); } else { recordFields[13].ClearValue(); }//Correspondant
                                    if (reader.IsDBNull(12) == false) { recordFields[14].SetValue((TrimDateTime)reader.GetDateTime(12)); } //DateReceived
                                        if (reader.IsDBNull(5) == false) { recordFields[15].SetValue((TrimDateTime)reader.GetDateTime(5)); } //DateRegistered
                                        if (reader.IsDBNull(4) == false) { recordFields[16].SetValue((TrimDateTime)reader.GetDateTime(4)); } //DeclaredDate
                                        if (reader.IsDBNull(2) == false) { recordFields[17].SetValue((TrimDateTime)reader.GetDateTime(2)); } //DocumentDate
                                        if (reader.IsDBNull(9) == false) { recordFields[18].SetValue(reader[9].ToString()); } //ExternalReference
                                    if (reader.IsDBNull(3) == false) { recordFields[19].SetValue(reader[3].ToString()); } else { recordFields[19].ClearValue(); }//FileName




                                    //
                                    //

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
                                            strErrorCol = strErrorCol + "Loader Error: " + reader.GetInt32(0).ToString() + ", " + " reader.GetInt32(0).ToString(): " + reader.GetInt32(0).ToString() + " reader.GetString(1): " + reader.GetString(1) + " reader.GetString(6): " + reader.GetString(6) + " reader.GetString(7)" + reader.GetString(7) + " reccont.Uri:" + reccont.Uri.ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Could no find folder: " + reader.GetString(6), reader.GetString(7));
                                        strErrorCol = strErrorCol + "Find folder Error: " + reader.GetInt32(0).ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                                    }
                                    //strErrorCol = strErrorCol + "No DocSet  for record number" + Environment.NewLine;
                                //}

                            }
                            catch (Exception exp)
                            {
                                Loggitt("ReadLine Error: " + exp.Message.ToString(), "GetInt32(0): "+ reader.GetInt32(0).ToString()+ ", GetString(1): "+ reader.GetString(1)+ ", GetString(2): " + reader.GetString(2) + ", GetString(3): " + reader.GetString(3) + ", GetString(4): " + reader.GetString(4) + ", GetString(5): " + reader.GetString(5) + ", GetString(6): " + reader.GetString(6) + ", GetString(7)"+ reader.GetString(7) + ", GetString(8)" + reader.GetString(8) + ", GetString(9)" + reader.GetString(9) + ", GetString(10)" + reader.GetString(10));
                            }
                        }
                    }
                }

                Console.WriteLine("Processing import batch ...");
                m_loader.ProcessAccumulatedData();
                Int64 runHistoryUri = m_loader.RunHistoryUri;
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
            //For tfr to SCC
            //zz legacy T1 ECM - Business Folders - Information Management - Projects
            //Local
            //zz Legacy T1 ECM - Business Folders
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
