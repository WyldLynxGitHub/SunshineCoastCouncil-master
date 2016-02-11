using HP.HPTRIM.SDK;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data;
using System.Reflection;


namespace caMigrateEcm
{
    public class bulkLoaderSample
    {      //}
        class ECMMigration
        {
            #region Properties  
            //0
            public string DocSetID { get; set; }
            //1
            public string ECMDescription { get; set; }
            //2
            public Nullable<DateTime> DocumentDate { get; set; }
            //3
            public string FileLocation { get; set; }
            //4
            public Nullable<DateTime> DeclaredDate { get; set; }
            //5
            public Nullable<DateTime> DateRegistered { get; set; }
            //6
            public string Bcs { get; set; }
            //7
            public string Folder { get; set; }
            //8
            public string Correspondant { get; set; }
            //9
            public string ExternalReference { get; set; }
            //10
            public int InternalReference { get; set; }
            //11
            public string SummaryText { get; set; }
            //12
            public Nullable<DateTime> DateReceived { get; set; }
            //13
            public string PropertyNo { get; set; }
            //14
            public string InfringementNo { get; set; }
            //15
            public string JobNo { get; set; }
            //16
            public string DocumentType { get; set; }
            //17
            public string ClassName { get; set; }
            //18
            public string BusinessCode { get; set; }

            #endregion
        }

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
                FieldDefinition fdDocumentType = new FieldDefinition(db, "ECM STD:DocumentType");
                recordFields[5] = new PropertyOrFieldValue(fdDocumentType);
                //
                FieldDefinition fdInfringementNo = new FieldDefinition(db, "ECM STD:InfringementNo");
                recordFields[6] = new PropertyOrFieldValue(fdInfringementNo);
                //
                FieldDefinition fdInternalReference = new FieldDefinition(db, "ECM STD:InternalReference");
                recordFields[7] = new PropertyOrFieldValue(fdInternalReference);
                //
                FieldDefinition fdJobNo = new FieldDefinition(db, "ECM STD:JobNo");
                recordFields[8] = new PropertyOrFieldValue(fdJobNo);
                //
                FieldDefinition fdPropertyNo = new FieldDefinition(db, "ECM STD:PropertyNo");
                recordFields[9] = new PropertyOrFieldValue(fdPropertyNo);
                //
                FieldDefinition fdSummaryText = new FieldDefinition(db, "ECM STD:SummaryText");
                recordFields[10] = new PropertyOrFieldValue(fdSummaryText);
                //New set
                FieldDefinition fdBusinessCode = new FieldDefinition(db, "ECM STD:BusinessCode");
                recordFields[11] = new PropertyOrFieldValue(fdBusinessCode);
                //
                FieldDefinition fdClassName = new FieldDefinition(db, "ECM STD:ClassName");
                recordFields[12] = new PropertyOrFieldValue(fdClassName);
                //
                FieldDefinition fdCorrespondant = new FieldDefinition(db, "ECM STD:Correspondant");
                recordFields[13] = new PropertyOrFieldValue(fdCorrespondant);
                //
                FieldDefinition fdDateReceived = new FieldDefinition(db, "ECM STD:DateReceived");
                recordFields[14] = new PropertyOrFieldValue(fdDateReceived);
                //
                FieldDefinition fdDateRegistered = new FieldDefinition(db, "ECM STD:DateRegistered");
                recordFields[15] = new PropertyOrFieldValue(fdDateRegistered);
                //
                FieldDefinition fdDeclaredDate = new FieldDefinition(db, "ECM STD:DeclaredDate");
                recordFields[16] = new PropertyOrFieldValue(fdDeclaredDate);
                //
                FieldDefinition fdDocumentDate = new FieldDefinition(db, "ECM STD:DocumentDate");
                recordFields[17] = new PropertyOrFieldValue(fdDocumentDate);
                //
                FieldDefinition fdExternalReference = new FieldDefinition(db, "ECM STD:ExternalReference");
                recordFields[18] = new PropertyOrFieldValue(fdExternalReference);
                //
                FieldDefinition fdFileName = new FieldDefinition(db, "ECM FileName");
                recordFields[19] = new PropertyOrFieldValue(fdFileName);
                //
                //string connectionString = "Data Source=yfdev.cloudapp.net;Initial Catalog=SCC_ECM;Persist Security Info=True;User ID=EPL;Password=Password1!";
                //string connectionString = "Data Source=MSI-GS60;Initial Catalog=ECMNew;Integrated Security=True";
                //Console.ReadLine();
                //
                Console.WriteLine("Enter Connection string");

                SqlConnection con = new SqlConnection(Console.ReadLine());
                //{
                con.Open();
                Console.WriteLine("Enter SQL string");
                using (SqlCommand command = new SqlCommand(Console.ReadLine(), con))
                {
                    DataSet dataSet = new DataSet();

                    DataTable dt = new DataTable();
                    //DataTable dtt = extension.

                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = command;
                    da.Fill(dt);
                    //List<ECMMigration> mytbl = Helper.DataTableToList<ECMMigration>(dt);
                    //dataSet.Tables[0].SerializeToJSon();
                    List<ECMMigration> lstecm = new List<ECMMigration>();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        //Console.WriteLine(reader.GetString(1));
                        ECMMigration ecm = new ECMMigration();
                        ecm.DocSetID = reader.GetInt32(0).ToString();
                        ecm.ECMDescription= reader.GetString(1);
                        lstecm.Add(ecm);

                    }

                    con.Close();
                    con.Dispose();

                    foreach(ECMMigration l in lstecm)
                    {
                        Console.WriteLine(l.DocSetID + " - " + l.ECMDescription);
                    }


                    //SqlDataReader dr = command.ExecuteReader();
                    //while (dr.Read())
                    //{
                    //    Console.WriteLine(dr.GetString(1));
                    //}
                    //dr.Close();
                    //cnn.Close();

                    //DataTable dta = extension.ToList<ECMMigration>(dt);
                    //List<ECMMigration> newStudents = dt.ToList<ECMMigration>();

                    //Stopwatch stopwatch1 = new Stopwatch();
                    //stopwatch1.Start();
                    var list = (from tr in dt.AsEnumerable()
                                select new ECMMigration()
                                {
                                    DocSetID = tr.Field<string>("STD:DocumentSetID"),
                                    ECMDescription = tr.Field<string>("STD:DocumentDescription"),
                                    DocumentDate=tr.Field<DateTime>("STD:DocumentDate"),
                                    FileLocation=tr.Field<string>("Expr1")??null,
                                    DeclaredDate = tr.Field<DateTime>("STD:DeclaredDate"),
                                    DateRegistered = tr.Field<DateTime>("STD:DateRegistered"),
                                    Bcs = tr.Field<string>("Expr2"),
                                    Folder = tr.Field<string>("Expr3"),
                                    Correspondant = tr.Field<string>("STD:Correspondant"),
                                    ExternalReference = tr.Field<string>("STD:ExternalReference"),
                                    InternalReference = tr.Field<int>("STD:InternalReference"),
                                    SummaryText = tr.Field<string>("STD:SummaryText"),
                                    DateReceived = tr.Field<DateTime>("STD:DateReceived"),
                                    PropertyNo = tr.Field<string>("STD:PropertyNo"),
                                    InfringementNo = tr.Field<string>("STD:InfringementNo"),
                                    JobNo = tr.Field<string>("STD:JobNo"),
                                    DocumentType = tr.Field<string>("STD:DocumentType"),
                                    ClassName = tr.Field<string>("STD:ClassName"),
                                    BusinessCode = tr.Field<string>("STD:BusinessCode")

                                }).ToList();

                    //Console.WriteLine("Result: " + h.DocSetID);
                    //Console.WriteLine("{0}\t{1}", h.DocSetID, h.ECMDescription);
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

                    string strTotal = list.Count().ToString();
                    int intCount = 0;

                    foreach (ECMMigration h in list)
                    {
                        intCount = intCount + 1;
                        try
                        {
                            //if (reader[0] != null)
                            //{
                            string loc = null;
                            if (strInloc != null)
                            {
                                loc = strInloc + h.FileLocation;
                            }

                            Record reccont = GetFolderUri(h.Bcs, h.Folder, db);
                            if (reccont != null)
                            {
                                Console.WriteLine("Importing ECM T1 record " + h.DocSetID.ToString() + " Row "+ intCount.ToString() +" of "+strTotal);
                                string fTitle;
                                if (h.ECMDescription == null)
                                {
                                    fTitle = "Record description from ECM Export is Null";
                                    strErrorCol = strErrorCol + "Ecm description to Title issue: Description Null for record no " + h.DocSetID.ToString() + Environment.NewLine;
                                }
                                else
                                {
                                    if (h.ECMDescription.Length > 253)
                                    {
                                        fTitle = h.ECMDescription.Substring(0, 253);
                                        strErrorCol = strErrorCol + "Ecm description to Title issue: Description greater then 254 ch, description truncated for record no " + h.DocSetID.ToString() + Environment.NewLine;
                                    }
                                    else if (h.ECMDescription.Length < 1)
                                    {
                                        fTitle = "Record description from ECM Export is Empty";
                                        strErrorCol = strErrorCol + "Ecm description to Title issue: Description blank for record no " + h.DocSetID.ToString() + Environment.NewLine;
                                    }
                                    else
                                    {
                                        fTitle = h.ECMDescription;
                                    }
                                }
                                recordFields[0].SetValue(fTitle); //Title
                                recordFields[1].SetValue((TrimDateTime)h.DocumentDate);
                                recordFields[2].SetValue("Migrated from EMS T1 using " + m_origin.Name);
                                recordFields[3].SetValue(reccont.Uri); //Container
                                recordFields[4].SetValue(h.DocSetID.ToString()); //Record number
                                if (h.DocumentType != null) { recordFields[5].SetValue(h.DocumentType); } else { recordFields[5].SetValue("No Document type"); } //Document type
                                if (h.InfringementNo != null) { recordFields[6].SetValue(h.InfringementNo); } else { recordFields[6].ClearValue(); }//InfringementNo
                                if (h.InternalReference != 0) { recordFields[7].SetValue(h.InternalReference); } else { recordFields[7].ClearValue(); } //InternalReference
                                if (h.JobNo != null) { recordFields[8].SetValue(h.JobNo); } else { recordFields[8].ClearValue(); }//JobNo
                                if (h.PropertyNo != null) { recordFields[9].SetValue(h.PropertyNo); } else { recordFields[9].ClearValue(); }//PropertyNo
                                if (h.SummaryText != null) { recordFields[10].SetValue(h.SummaryText); } else { recordFields[10].ClearValue(); }//SummaryText
                                if (h.BusinessCode != null) { recordFields[11].SetValue(h.BusinessCode); } else { recordFields[11].ClearValue(); } //BusinessCode
                                if (h.ClassName != null) { recordFields[12].SetValue(h.ClassName); } else { recordFields[12].ClearValue(); }//ClassName
                                if (h.Correspondant != null) { recordFields[13].SetValue(h.Correspondant); } else { recordFields[13].ClearValue(); }//Correspondant
                                if (h.DateReceived != null) { recordFields[14].SetValue((TrimDateTime)h.DateReceived); } else { recordFields[14].ClearValue(); }//DateReceived
                                if (h.DateRegistered != null) { recordFields[15].SetValue((TrimDateTime)h.DateRegistered); } else { recordFields[15].ClearValue(); }//DateRegistered
                                if (h.DeclaredDate != null) { recordFields[16].SetValue((TrimDateTime)h.DeclaredDate); } else { recordFields[16].ClearValue(); }//DeclaredDate
                                if (h.DocumentDate != null) { recordFields[17].SetValue((TrimDateTime)h.DocumentDate); } else { recordFields[17].ClearValue(); }//DocumentDate
                                if (h.ExternalReference != null) { recordFields[18].SetValue(h.ExternalReference); } else { recordFields[18].ClearValue(); }//ExternalReference
                                if (h.FileLocation != null) { recordFields[19].SetValue(h.FileLocation); } else { recordFields[19].ClearValue(); }//FileName

                                Record importRec = m_loader.NewRecord();
                                //SecurityLevel sl = new SecurityLevel(db, 2);
                                //importRec.SecurityProfile.SecurityLevel = sl;


                                //if(var recUri = m_loader.FindRecord(h.DocSetID.ToString())==0)
                                //    {

                                //}

                                //Record UpdateRec = m_loader.UpdateDocument()

                                //importRec.DateFinalized
                                //importRec.SetAsFinal

                                
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
                                    strErrorCol = strErrorCol + "Loader Error: " + h.DocSetID.ToString() + ", " + " ECMDescription: " + h.ECMDescription + " Bcs: " + h.Bcs + " Folder" + h.Folder + " reccont.Uri:" + reccont.Uri.ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                                }
                            }
                            else
                            {
                                Console.WriteLine("Could no find folder: " + " Bcs: " + h.Bcs + " Folder" + h.Folder);
                                strErrorCol = strErrorCol + "Find folder Error: " + h.DocSetID.ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                            }
                        }
                        catch (Exception exp)
                        {
                            Loggitt("ReadLine Error: " + exp.Message.ToString(), " DocSetID: " + h.DocSetID + ", ECMDescription: " + h.ECMDescription + ", GetString(2): " + h.DocumentDate.ToString()+ ", FileLocation: " + h.FileLocation + ", DeclaredDate: " + h.DeclaredDate.ToString() + ", DateRegistered: " + h.DateRegistered.ToString() + ", Bcs: " + h.Bcs + ", Folder:" + h.Folder);
                        }
                    }
                    //stopwatch1.Stop();
                    //TimeSpan ts1 = stopwatch1.Elapsed;
                    //string elapsedTime1 = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts1.Hours, ts1.Minutes, ts1.Seconds, ts1.Milliseconds / 10);
                    //Console.WriteLine("Process Time " + elapsedTime1);
                }




                //
                //Console.WriteLine("Enter Connection string");
                //using (SqlConnection con = new SqlConnection(Console.ReadLine()))
                //{
                //    con.Open();
                //    Console.WriteLine("Enter SQL string");
                //    string strloc = null;
                //    using (SqlCommand command = new SqlCommand(Console.ReadLine(), con))

                //    using (SqlDataReader reader = command.ExecuteReader())
                //    {
                //        //Console.WriteLine("Query count: " + reader.FieldCount.ToString());
                //        string strIndexLoc = null;
                //        Console.WriteLine("Enter the network location of the index");
                //        string strInloc = Console.ReadLine();
                //        Console.WriteLine("Copy documents or remove");
                //        BulkLoaderCopyMode windowsmode = BulkLoaderCopyMode.WindowsCopy;
                //        switch (Console.ReadLine())
                //        {
                //            case "C":
                //                windowsmode = BulkLoaderCopyMode.WindowsCopy;
                //                break;
                //            case "R":
                //                windowsmode = BulkLoaderCopyMode.WindowsMove;
                //                break;
                //            default:
                //                windowsmode = BulkLoaderCopyMode.WindowsCopy;
                //                break;
                //        }
                //        //Console.ReadLine();
                //        if (!Directory.Exists(strInloc))
                //        {
                //            strInloc = null;
                //        }





                //        while (reader.Read())
                //        {
                //            //try
                //            //{
                //            //    //if (reader[0] != null)
                //            //    //{
                //            //        string loc = null;
                //            //        if (strInloc != null)
                //            //        {
                //            //            loc = strInloc + reader.GetString(3);
                //            //        }

                //            //        Record reccont = GetFolderUri(reader.GetString(6), reader.GetString(7), db);
                //            //        if (reccont != null)
                //            //        {
                //            //            Console.WriteLine("Importing ECM T1 record " + reader.GetInt32(0).ToString() + " ...");
                //            //        string fTitle;
                //            //        if (reader.IsDBNull(1) == true)
                //            //        {
                //            //            fTitle = "Record description from ECM Export is Null";
                //            //            strErrorCol = strErrorCol + "Ecm description to Title issue: Description Null for record no " + reader.GetInt32(0).ToString() + Environment.NewLine;
                //            //        }
                //            //        else
                //            //        {
                //            //            if(reader.GetString(1).Length>253)
                //            //            {
                //            //                fTitle = reader.GetString(1).Substring(0, 253);
                //            //                strErrorCol = strErrorCol + "Ecm description to Title issue: Description greater then 254 ch, description truncated for record no " + reader.GetInt32(0).ToString() + Environment.NewLine;
                //            //            }
                //            //            else if(reader.GetString(1).Length<1)
                //            //            {
                //            //                fTitle = "Record description from ECM Export is Empty";
                //            //                strErrorCol = strErrorCol + "Ecm description to Title issue: Description blank for record no " + reader.GetInt32(0).ToString() + Environment.NewLine;
                //            //            }
                //            //            else
                //            //            {
                //            //                fTitle = reader.GetString(1);
                //            //            }
                //            //        }
                //            //            recordFields[0].SetValue(fTitle); //Title
                //            //            recordFields[1].SetValue((TrimDateTime)reader.GetDateTime(2));
                //            //            recordFields[2].SetValue("Migrated from EMS T1 using " + m_origin.Name);
                //            //            recordFields[3].SetValue(reccont.Uri); //Container
                //            //            recordFields[4].SetValue(reader.GetInt32(0).ToString()); //Record number

                //            //            //Potential null fileds
                //            //            if (reader.IsDBNull(16) == false) { recordFields[5].SetValue(reader.GetString(16)); } else { recordFields[5].SetValue("No Document type"); } //Document type
                //            //            if (reader.IsDBNull(14) == false) { recordFields[6].SetValue(reader[14].ToString()); } else { recordFields[6].ClearValue(); }//InfringementNo
                //            //        if (reader.IsDBNull(10) == false) { recordFields[7].SetValue(reader[10].ToString()); } else { recordFields[7].ClearValue(); } //InternalReference
                //            //        if (reader.IsDBNull(16) == false) { recordFields[8].SetValue(reader[15].ToString()); } else { recordFields[8].ClearValue(); }//JobNo
                //            //            if (reader.IsDBNull(13) == false) { recordFields[9].SetValue(reader[13].ToString()); } else { recordFields[9].ClearValue(); }//PropertyNo
                //            //        if (reader.IsDBNull(11) == false) { recordFields[10].SetValue(reader[11].ToString()); } else { recordFields[10].ClearValue(); }//SummaryText
                //            //                                                                                                                                      //
                //            //                                                                                                                                      //
                //            //        if (reader.IsDBNull(18) == false) { recordFields[11].SetValue(reader[18].ToString()); } else { recordFields[11].ClearValue(); } //BusinessCode
                //            //        if (reader.IsDBNull(17) == false) { recordFields[12].SetValue(reader[17].ToString()); } else { recordFields[12].ClearValue(); }//ClassName
                //            //        if (reader.IsDBNull(8) == false) { recordFields[13].SetValue(reader[8].ToString()); } else { recordFields[13].ClearValue(); }//Correspondant
                //            //        if (reader.IsDBNull(12) == false) { recordFields[14].SetValue((TrimDateTime)reader.GetDateTime(12)); } //DateReceived
                //            //            if (reader.IsDBNull(5) == false) { recordFields[15].SetValue((TrimDateTime)reader.GetDateTime(5)); } //DateRegistered
                //            //            if (reader.IsDBNull(4) == false) { recordFields[16].SetValue((TrimDateTime)reader.GetDateTime(4)); } //DeclaredDate
                //            //            if (reader.IsDBNull(2) == false) { recordFields[17].SetValue((TrimDateTime)reader.GetDateTime(2)); } //DocumentDate
                //            //            if (reader.IsDBNull(9) == false) { recordFields[18].SetValue(reader[9].ToString()); } //ExternalReference
                //            //        if (reader.IsDBNull(3) == false) { recordFields[19].SetValue(reader[3].ToString()); } else { recordFields[19].ClearValue(); }//FileName

                //            //        Record importRec = m_loader.NewRecord();


                //            //            if (File.Exists(loc))
                //            //            {
                //            //                InputDocument doc = new InputDocument();
                //            //                doc.SetAsFile(loc);

                //            //                m_loader.SetDocument(importRec, doc, windowsmode);
                //            //            }

                //            //            m_loader.SetProperties(importRec, recordFields);
                //            //            m_loader.SubmitRecord(importRec);

                //            //            if (m_loader.Error.Bad == true)
                //            //            {
                //            //                Console.WriteLine("Error -  " + m_loader.Error.Message + " " + m_loader.Error.Source);
                //            //                strErrorCol = strErrorCol + "Loader Error: " + reader.GetInt32(0).ToString() + ", " + " reader.GetInt32(0).ToString(): " + reader.GetInt32(0).ToString() + " reader.GetString(1): " + reader.GetString(1) + " reader.GetString(6): " + reader.GetString(6) + " reader.GetString(7)" + reader.GetString(7) + " reccont.Uri:" + reccont.Uri.ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                //            //            }
                //            //        }
                //            //        else
                //            //        {
                //            //            Console.WriteLine("Could no find folder: " + reader.GetString(6), reader.GetString(7));
                //            //            strErrorCol = strErrorCol + "Find folder Error: " + reader.GetInt32(0).ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                //            //        }
                //            //}
                //            //catch (Exception exp)
                //            //{
                //            //    Loggitt("ReadLine Error: " + exp.Message.ToString(), "GetInt32(0): "+ reader.GetInt32(0).ToString()+ ", GetString(1): "+ reader.GetString(1)+ ", GetString(2): " + reader.GetString(2) + ", GetString(3): " + reader.GetString(3) + ", GetString(4): " + reader.GetString(4) + ", GetString(5): " + reader.GetString(5) + ", GetString(6): " + reader.GetString(6) + ", GetString(7)"+ reader.GetString(7) + ", GetString(8)" + reader.GetString(8) + ", GetString(9)" + reader.GetString(9) + ", GetString(10)" + reader.GetString(10));
                //            //}
                //        }
                //    }
                //}
                stopwatch.Stop();
                Console.WriteLine("Process import batch? Enter to continue.");
                Console.ReadLine();
                m_loader.ProcessAccumulatedData();
                Int64 runHistoryUri = m_loader.RunHistoryUri;
                

                Console.WriteLine("Processing complete ...");
                m_loader.EndRun();


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
        /*Converts DataTable To List*/
        //public static List<TSource> ToList<TSource>(this DataTable dataTable) where TSource : new()
        //{
        //    var dataList = new List<TSource>();

        //    const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic;
        //    var objFieldNames = (from PropertyInfo aProp in typeof(TSource).GetProperties(flags)
        //                         select new
        //                         {
        //                             Name = aProp.Name,
        //                             Type = Nullable.GetUnderlyingType(aProp.PropertyType) ??
        //                 aProp.PropertyType
        //                         }).ToList();
        //    var dataTblFieldNames = (from DataColumn aHeader in dataTable.Columns
        //                             select new
        //                             {
        //                                 Name = aHeader.ColumnName,
        //                                 Type = aHeader.DataType
        //                             }).ToList();
        //    var commonFields = objFieldNames.Intersect(dataTblFieldNames).ToList();

        //    foreach (DataRow dataRow in dataTable.AsEnumerable().ToList())
        //    {
        //        var aTSource = new TSource();
        //        foreach (var aField in commonFields)
        //        {
        //            PropertyInfo propertyInfos = aTSource.GetType().GetProperty(aField.Name);
        //            var value = (dataRow[aField.Name] == DBNull.Value) ?
        //            null : dataRow[aField.Name]; //if database field is nullable
        //            propertyInfos.SetValue(aTSource, value, null);
        //        }
        //        dataList.Add(aTSource);
        //    }
        //    return dataList;
        //}
    }
}
