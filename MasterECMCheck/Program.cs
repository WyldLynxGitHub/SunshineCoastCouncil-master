using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.IO;

namespace MasterECMCheck
{
    class Program
    {
        class ECM
        {
            public int ECMDocumentSetID { get; set; }
            public string ECMOrigbatch { get; set; }
            public string RecId { get; set; }
            public string docLoc { get; set; }
        }
        public static string strCon;
        static void Main(string[] args)
        {
            TrimApplication.Initialize();

            int iEcmCreated;
            int iEcmNoFile;
            int iEcmNotCreated;
            strCon = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            //strCon = "Data Source=MSI-GS60;Initial Catalog=ECMStage;Integrated Security=True";
            string sds = "Select DocumentSetID, Origbatch, ClassName, maxversion from ECM_MasterCheck Where OrigBatch = 'Extract Remainder' and rmuri is null order by DocumentSetID";
            //Console.WriteLine("Enter to proceed");
            //Console.ReadLine();
            List<ECM> lstEcm = new List<ECM>();
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(strCon))
            {
                con.Open();
                //Console.WriteLine("Enter to proceed - conn open: "+con.Database);
                //Console.ReadLine();
                Console.WriteLine("Enter SQL statement:");


                //
                using (SqlCommand command = new SqlCommand(sds, con))
                {

                    DataSet dataSet = new DataSet();
                    
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = command;
                    da.Fill(dt);
                    Console.WriteLine("Datatable complete: Rows "+dt.Rows.Count.ToString());
                    //foreach (DataRow dr in dt.Rows)
                    //{
                    //    foreach (DataColumn dc in dt.Columns)
                    //    {
                    //        var field1 = dr[dc].ToString();
                    //        Console.WriteLine("Row no: " + field1);

                    //    }
                    //    string sdsd s= dr.Field<string>(0);
                    //}
                    //Console.ReadLine();
                    //
                    //using (SqlDataReader reader = command.ExecuteReader())
                    //{
                    //    int ECMDocumentSetID = reader.GetOrdinal("DocumentSetID");
                    //    int ECMOrigbatch = reader.GetOrdinal("Origbatch");
                    //    int ECMClassName = reader.GetOrdinal("ClassName");
                    //    int ECMmaxversion = reader.GetOrdinal("maxversion");
                    //    //int ECMVolumeUpdatable = reader.GetOrdinal("maxversion");
                    iEcmCreated = 0;
                    iEcmNoFile = 0;
                    iEcmNotCreated = 0;
                    //    Console.WriteLine("Start reading SQL");
                    //    while (reader.Read())
                    //    {
                    //        ECM ecm = new ECM();
                    //        ecm.ECMDocumentSetID= reader.GetInt32(ECMDocumentSetID);
                    //        ecm.ECMOrigbatch= reader.GetString(ECMOrigbatch);
                    //        ecm.RecId= reader.GetInt32(ECMDocumentSetID).ToString();
                    //        ecm.docLoc= CheckAttachment(reader.GetInt32(ECMDocumentSetID).ToString(), reader.GetString(ECMOrigbatch));
                    //        lstEcm.Add(ecm);
                    //        Console.WriteLine("Adding slq to list: " + reader.GetInt32(ECMDocumentSetID).ToString());
                    //    }
                    //    Console.WriteLine("Finished reading Sql - count: " + lstEcm.Count().ToString());
                    //}
                }
            }
            

            using (Database db = new Database())
            {
                //foreach (ECM ecm in lstEcm)
                foreach (DataRow dr in dt.Rows)
                {
                    Console.WriteLine("Check ecm record: " + dr.Field<Int32>(0).ToString());
                    var sDocLoc = CheckAttachment(dr.Field<Int32>(0).ToString(), dr.Field<string>(1));
                    //}
                    Record r = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, dr.Field<Int32>(0).ToString());
                    //new Record(db, );
                    if (r != null)
                    {
                        if (r.IsElectronic)
                        {
                            Console.WriteLine("Confirm record created: " + r.Number + " and file attached.");
                            iEcmCreated++;
                            UpdateMasterinRM(r);
                        }
                        else
                        {
                            if (File.Exists(sDocLoc))
                            {
                                Console.WriteLine("File found and available: " + r.Number);
                                UpdateMasterinRMnoDoc(r, 0);
                            }
                            else
                            {
                                Console.WriteLine("File could not be found: " + r.Number + " - " + sDocLoc);
                                UpdateMasterinRMnoDoc(r, 1);
                            }

                            iEcmNoFile++;
                            //UpdateMasterinRMnoDoc(intECMDocumentSetID.ToString(), r);
                        }

                    }
                    else
                    {
                        iEcmNotCreated++;
                        if (File.Exists(sDocLoc))
                        {
                            Console.WriteLine("ECM record not created but file exists");
                            UpdateMasterNoRm(dr.Field<Int32>(0).ToString(), 0);
                        }
                        else
                        {
                            Console.WriteLine("ECM record not created and no file exists/");
                            UpdateMasterNoRm(dr.Field<Int32>(0).ToString(), 1);
                        }
                    }
                    //}
                }
            }
            //    }
            //}
            //}
            Console.WriteLine("Summary:" + Environment.NewLine + "ECM complete migrations: " + iEcmCreated.ToString() + Environment.NewLine + "ECM Created in Eddie, no attachment: " + iEcmNoFile.ToString() + Environment.NewLine + "ECM Not Created: " + iEcmNotCreated.ToString());
            Console.ReadLine();
            //
        }


        private static void UpdateMasterinRM(Record r)
        {
                long ruri = r.Uri;
                int rRev = r.RevisionNumber;
                //"Select DocumentSetID, Origbatch, ClassName, maxversion from ECM_MasterCheck"
                string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, RMUri=@rmuri, RmRevisions=@ver, Issue=@issue WHERE DocumentSetID=@docsetid";
            try
            {
                using (SqlConnection con = new SqlConnection(strCon))
                {
                    using (SqlCommand cmd = new SqlCommand(SQL, con))
                    {
                        cmd.Parameters.AddWithValue("@docsetid", r.Number);
                        cmd.Parameters.AddWithValue("@now", DateTime.Now);
                        cmd.Parameters.AddWithValue("@rmuri", ruri);
                        cmd.Parameters.AddWithValue("@ver", rRev);
                        cmd.Parameters.AddWithValue("@issue", 0);
                        con.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("SQL error: " + exp.Message.ToString());
            }

        }
        private static void UpdateMasterNoRm(string recnum, int prob)
        {
                string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, Issue=@issue WHERE DocumentSetID=@docsetid";
            try
            {
                using (SqlConnection con = new SqlConnection(strCon))
                {
                    using (SqlCommand cmd = new SqlCommand(SQL, con))
                    {
                        cmd.Parameters.AddWithValue("@docsetid", recnum);
                        cmd.Parameters.AddWithValue("@now", DateTime.Now);
                        cmd.Parameters.AddWithValue("@issue", prob);
                        //cmd.Parameters.AddWithValue("@rmuri", ruri);
                        //cmd.Parameters.AddWithValue("@ver", rRev);
                        con.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("SQL error: " + exp.Message.ToString());
            }
                //con.Close();
        }
        private static void UpdateMasterinRMnoDoc(Record r, int prob)
        {
            long ruri = r.Uri;
            int rRev = 0;
            //"Select DocumentSetID, Origbatch, ClassName, maxversion from ECM_MasterCheck"
            string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, RMUri=@rmuri, RmRevisions=@ver, Issue=@issue WHERE DocumentSetID=@docsetid";
            try
            {
                using (SqlConnection con = new SqlConnection(strCon))
                {
                    using (SqlCommand cmd = new SqlCommand(SQL, con))
                    {
                        cmd.Parameters.AddWithValue("@docsetid", r.Number);
                        cmd.Parameters.AddWithValue("@now", DateTime.Now);
                        cmd.Parameters.AddWithValue("@rmuri", ruri);
                        cmd.Parameters.AddWithValue("@ver", rRev);
                        cmd.Parameters.AddWithValue("@issue", prob);
                        con.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
                //con.Close();
            }
            catch (Exception exp)
            {
                Console.WriteLine("SQL error: " + exp.Message.ToString());
            }

        }
        private static string CheckAttachment(string filelocUri, string Batch)
        {
            //Console.WriteLine("CheckAttachment: filelocUri " + filelocUri+ " Batch "+ Batch);
            //Console.ReadLine();
            //Select * from [Elai] Where ID in(SELECT MIN(ID) AS ID FROM [Elai] GROUP BY [STD:DocumentID]) and [STD:Version] = 1 Order by [ID] OFFSET 1457802 ROWS FETCH NEXT 5 ROWS ONLY
            //Select [Volume:StorageLocation], [Volume:Filename] from [Extract Public Indexed] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Public Indexed] GROUP BY [STD:DocumentID]) and [STD: Version] = 1 and [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri)
            string sqlBatch = null;
            sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from eri Where ID in (SELECT MIN(ID) AS ID FROM eri GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            string strStorageLoc = null;
            //switch (Batch)
            //{
            //    case "Extract Public":
            //        //sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from [Extract Public Indexed] where [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            //        sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from [Extract Public Indexed] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Public Indexed] GROUP BY [STD:DocumentID]) and [STD: Version] = 1 and [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            //        break;
            //    case "Extract Public Zip":
            //        //sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from [Extract Public Zip Indexed] where [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            //        sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from [Extract Public Zip Indexed] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Public Zip Indexed] GROUP BY [STD:DocumentID]) and [STD: Version] = 1 and [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            //        break;
            //    case "Linked To Application":
            //        //sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from Elai where [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            //        sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from Elai Where ID in(SELECT MIN(ID) AS ID FROM Elai GROUP BY [STD:DocumentID]) and [STD: Version] = 1 and [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);

            //        break;
            //    case "Linked To Property":
            //        //sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from Elpi where [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);
            //        sqlBatch = "Select [Volume:StorageLocation], [Volume:Filename] from Elpi Where ID in(SELECT MIN(ID) AS ID FROM Elpi GROUP BY [STD:DocumentID]) and [STD: Version] = 1 and [STD:DocumentSetID] = " + Convert.ToInt32(filelocUri);

            //        break;
            //}
            //[Volume:StorageLocation] ,[Volume: Filename]
            //Console.WriteLine("After switch: sqlBatch " + sqlBatch);
            //Console.ReadLine();
            using (SqlConnection con1 = new SqlConnection(strCon))
            {
                con1.Open();
                //Console.WriteLine("Find doc loc string, Conn " + con1.Database);
                //Console.ReadLine();
                using (SqlCommand command = new SqlCommand(sqlBatch, con1))
                {
                    //DataSet dataSet = new DataSet();
                    //DataTable dt = new DataTable();
                    //SqlDataAdapter da = new SqlDataAdapter();
                    //da.SelectCommand = command;
                    //da.Fill(dt);
                    SqlDataReader reader = command.ExecuteReader();
                    //int ECMStorageLocation = reader.GetOrdinal("[Volume:StorageLocation]");
                    //int ECMFilename = reader.GetOrdinal("[Volume:Filename]");

                    while (reader.Read())
                    {
                        strStorageLoc = FindMappedStorage(reader.GetString(0)) + "\\" + reader.GetString(1);
                    }

                }
            }
            return strStorageLoc;
        }
        public static void CreateNewEddieRecord(Database db, long uri)
        {
            Record rec = new Record(db, uri);


            //recordFields[0] = new PropertyOrFieldValue(PropertyIds.RecordTitle);
            //recordFields[1] = new PropertyOrFieldValue(PropertyIds.RecordDateCreated);
            //recordFields[2] = new PropertyOrFieldValue(PropertyIds.RecordNotes);
            //recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordContainer);
            //recordFields[4] = new PropertyOrFieldValue(PropertyIds.RecordNumber);
            //recordFields[110] = new PropertyOrFieldValue(PropertyIds.RecordBarcode);

            //
            FieldDefinition fd_ECMApplicationRAM_Application_Number = new FieldDefinition(db, "Application Number List");
            FieldDefinition fd_ECMBarCodeBarCodeString = new FieldDefinition(db, "ECM BarCode:BarCodeString");
            FieldDefinition fd_ECMBCSRetention_Period = new FieldDefinition(db, "ECM BCS:Retention_Period");
            FieldDefinition fd_ECMCaseCaseDescription = new FieldDefinition(db, "ECM Case:CaseDescription");
            FieldDefinition fd_ECMCaseCaseName = new FieldDefinition(db, "ECM Case:CaseName");
            FieldDefinition fd_ECMCaseCaseNumber = new FieldDefinition(db, "ECM Case:CaseNumber");
            FieldDefinition fd_ECMCaseCaseOfficer = new FieldDefinition(db, "ECM Case:CaseOfficer");
            FieldDefinition fd_ECMCaseCaseStatus = new FieldDefinition(db, "ECM Case:CaseStatus");
            FieldDefinition fd_ECMCaseCaseType = new FieldDefinition(db, "ECM Case:CaseType");
            FieldDefinition fd_ECMCaseEndDate = new FieldDefinition(db, "ECM Case:EndDate");
            FieldDefinition fd_ECMCaseStartDate = new FieldDefinition(db, "ECM Case:StartDate");
            FieldDefinition fd_ECMCorrespondentDescription = new FieldDefinition(db, "ECM Correspondent:Description");
            FieldDefinition fd_ECMECMDriveID = new FieldDefinition(db, "ECM ECMDriveID");
            FieldDefinition fd_ECMEmployeeCommencement_Date = new FieldDefinition(db, "ECM Employee:Commencement_Date");
            FieldDefinition fd_ECMEmployeeEmployee_Given_Name = new FieldDefinition(db, "ECM Employee:Employee_Given_Name");
            FieldDefinition fd_ECMEmployeeEmployee_Surname = new FieldDefinition(db, "ECM Employee:Employee_Surname");
            FieldDefinition fd_ECMEmployeeEmployeeCode = new FieldDefinition(db, "ECM Employee:EmployeeCode");
            FieldDefinition fd_ECmEmployeeHR_Case_Description = new FieldDefinition(db, "ECm Employee:HR_Case_Description");
            FieldDefinition fd_ECMEmployeePreferred_Name = new FieldDefinition(db, "ECM Employee:Preferred_Name");
            FieldDefinition fd_ECMEmployeeStatus = new FieldDefinition(db, "ECM Employee:Status");
            FieldDefinition fd_ECMEmployeeTermination_Date = new FieldDefinition(db, "ECM Employee:Termination_Date");
            FieldDefinition fd_ECMFileName = new FieldDefinition(db, "ECM FileName");
            FieldDefinition fd_ECMFileNetFoldername = new FieldDefinition(db, "ECM FileNet:Foldername");
            FieldDefinition fd_ECMHR_Case_Code = new FieldDefinition(db, "ECM HR_Case_Code");
            FieldDefinition fd_ECMInfoExpertLevel1_Name = new FieldDefinition(db, "ECM InfoExpert:Level1_Name");
            FieldDefinition fd_ECMInfoExpertLevel2_Name = new FieldDefinition(db, "ECM InfoExpert:Level2_Name");
            FieldDefinition fd_ECMInfoExpertLevel3_Description = new FieldDefinition(db, "ECM InfoExpert:Level3_Description");
            FieldDefinition fd_ECMInfoExpertLevel3_Name = new FieldDefinition(db, "ECM InfoExpert:Level3_Name");
            FieldDefinition fd_ECMInfoExpertLevel4_Description = new FieldDefinition(db, "ECM InfoExpert:Level4_Description");
            FieldDefinition fd_ECMInfoExpertLevel4_Name = new FieldDefinition(db, "ECM InfoExpert:Level4_Name");
            FieldDefinition fd_ECMInfoExpertLevel5_Description = new FieldDefinition(db, "ECM InfoExpert:Level5_Description");
            FieldDefinition fd_ECMInfoExpertLevel5_Name = new FieldDefinition(db, "ECM InfoExpert:Level5_Name");
            FieldDefinition fd_ECMInfringementInfringementID = new FieldDefinition(db, "ECM Infringement:InfringementID");
            FieldDefinition fd_ECMNotes = new FieldDefinition(db, "ECM Notes");
            FieldDefinition fd_ECMPositionClose_Date = new FieldDefinition(db, "ECM Position:Close_Date");
            FieldDefinition fd_ECMPositionOpen_Date = new FieldDefinition(db, "ECM Position:Open_Date");
            FieldDefinition fd_ECMPositionPosition_ID = new FieldDefinition(db, "ECM Position:Position_ID");
            FieldDefinition fd_ECMPositionPosition_Title = new FieldDefinition(db, "ECM Position:Position_Title");
            FieldDefinition fd_ECMPositionVacancy_ID = new FieldDefinition(db, "ECM Position:Vacancy_ID");
            FieldDefinition fd_ECMPositionVacancy_Status = new FieldDefinition(db, "ECM Position:Vacancy_Status");
            FieldDefinition fd_ECMProjectAndContractProject_Name = new FieldDefinition(db, "ECM ProjectAndContract:[Project_Name");
            FieldDefinition fd_ECMProjectAndContractContract_or_Activity_Desc = new FieldDefinition(db, "ECM ProjectAndContract:Contract_or_Activity_Desc");
            FieldDefinition fd_ECMProjectAndContractContract_or_Activity_Name = new FieldDefinition(db, "ECM ProjectAndContract:Contract_or_Activity_Name");
            FieldDefinition fd_ECMProjectAndContractContract_or_Activity_No = new FieldDefinition(db, "ECM ProjectAndContract:Contract_or_Activity_No");
            FieldDefinition fd_ECMProjectAndContractEnd_Date = new FieldDefinition(db, "ECM ProjectAndContract:End_Date");
            FieldDefinition fd_ECMProjectAndContractProject_Description = new FieldDefinition(db, "ECM ProjectAndContract:Project_Description");
            FieldDefinition fd_ECMProjectAndContractProject_End_Date = new FieldDefinition(db, "ECM ProjectAndContract:Project_End_Date");
            FieldDefinition fd_ECMProjectAndContractProject_Officer = new FieldDefinition(db, "ECM ProjectAndContract:Project_Officer");
            FieldDefinition fd_ECMProjectAndContractProject_Start_Date = new FieldDefinition(db, "ECM ProjectAndContract:Project_Start_Date");
            FieldDefinition fd_ECMProjectAndContractProjectNo = new FieldDefinition(db, "ECM ProjectAndContract:ProjectNo");
            FieldDefinition fd_ECMProjectAndContractResponsible_Officer = new FieldDefinition(db, "ECM ProjectAndContract:Responsible_Officer");
            FieldDefinition fd_ECMProjectAndContractStart_Date = new FieldDefinition(db, "ECM ProjectAndContract:Start_Date");
            FieldDefinition fd_ECMProjectAndContractStatus = new FieldDefinition(db, "ECM ProjectAndContract:Status");
            FieldDefinition fd_ECMPropertyHouse_No = new FieldDefinition(db, "ECM Property:House_No");
            FieldDefinition fd_ECMPropertyHouse_No_Suffix = new FieldDefinition(db, "ECM Property:House_No_Suffix");
            FieldDefinition fd_ECMPropertyHouse_No_To = new FieldDefinition(db, "ECM Property:House_No_To");
            FieldDefinition fd_ECMPropertyHouse_No_To_Suffix = new FieldDefinition(db, "ECM Property:House_No_To_Suffix");
            FieldDefinition fd_ECMPropertyLocality_Name = new FieldDefinition(db, "ECM Property:Locality_Name");
            FieldDefinition fd_ECMPropertyPostcode = new FieldDefinition(db, "ECM Property:Postcode");
            FieldDefinition fd_ECMPropertyProperty_Name = new FieldDefinition(db, "ECM Property:Property_Name");
            FieldDefinition fd_ECMPropertyProperty_No = new FieldDefinition(db, "Property Number List");
            FieldDefinition fd_ECMPropertyStreet_Name = new FieldDefinition(db, "ECM Property:Street_Name");
            FieldDefinition fd_ECMPropertyUnit_No = new FieldDefinition(db, "ECM Property:Unit_No");
            FieldDefinition fd_ECMPropertyUnit_No_Suffix = new FieldDefinition(db, "ECM Property:Unit_No_Suffix");
            FieldDefinition fd_ECMreference = new FieldDefinition(db, "ECM reference");
            FieldDefinition fd_ECMRelatedDocDocSetID = new FieldDefinition(db, "ECM RelatedDoc:DocSetID");
            FieldDefinition fd_ECMSTDAllRevTitle = new FieldDefinition(db, "ECM STD:AllRevTitle");
            FieldDefinition fd_ECMSTDApplicationNo = new FieldDefinition(db, "ECM STD:ApplicationNo");
            FieldDefinition fd_ECMSTDBusinessCode = new FieldDefinition(db, "ECM STD:BusinessCode");
            FieldDefinition fd_ECMSTDClassName = new FieldDefinition(db, "ECM STD:ClassName");
            FieldDefinition fd_ECMSTDCorrespondant = new FieldDefinition(db, "ECM STD:Correspondant");
            FieldDefinition fd_ECMSTDCustomerRequest = new FieldDefinition(db, "ECM STD:CustomerRequest");
            FieldDefinition fd_ECMSTDDateLastAccessed = new FieldDefinition(db, "ECM STD:DateLastAccessed");
            FieldDefinition fd_ECMSTDDateReceived = new FieldDefinition(db, "ECM STD:DateReceived");
            FieldDefinition fd_ECMSTDDateRegistered = new FieldDefinition(db, "ECM STD:DateRegistered");
            FieldDefinition fd_ECMSTDDeclaredDate = new FieldDefinition(db, "ECM STD:DeclaredDate");
            FieldDefinition fd_ECMSTDDestructionDue = new FieldDefinition(db, "ECM STD:DestructionDue");
            FieldDefinition fd_ECMSTDDocumentDate = new FieldDefinition(db, "ECM STD:DocumentDate");
            FieldDefinition fd_ECMSTDDocumentType = new FieldDefinition(db, "ECM STD:DocumentType");
            FieldDefinition fd_ECMSTDExternalReference = new FieldDefinition(db, "ECM STD:ExternalReference");
            FieldDefinition fd_ECMSTDInfringementNo = new FieldDefinition(db, "Infringement Number");
            FieldDefinition fd_ECMSTDInternalReference = new FieldDefinition(db, "ECM STD:InternalReference");
            FieldDefinition fd_ECMSTDInternalScanRequest = new FieldDefinition(db, "ECM STD:InternalScanRequest");
            FieldDefinition fd_ECMSTDJobNo = new FieldDefinition(db, "ECM STD:JobNo");
            FieldDefinition fd_ECMSTDOtherReferences = new FieldDefinition(db, "ECM STD:OtherReferences");
            FieldDefinition fd_ECMSTDPropertyNo = new FieldDefinition(db, "ECM STD:PropertyNo");
            FieldDefinition fd_ECMSTDSummaryText = new FieldDefinition(db, "ECM STD:SummaryText");
            FieldDefinition fd_ECMStreetsLocality = new FieldDefinition(db, "ECM Streets:Locality");
            FieldDefinition fd_ECMStreetsPostcode = new FieldDefinition(db, "ECM Streets:Postcode");
            FieldDefinition fd_ECMStreetsStreet_Name = new FieldDefinition(db, "ECM Streets:Street_Name");
            FieldDefinition fd_ECMUserGroupsDescription = new FieldDefinition(db, "ECM UserGroups:Description");
            FieldDefinition fd_ECMUserGroupsExtNo = new FieldDefinition(db, "ECM UserGroups:ExtNo");
            FieldDefinition fd_ECMUserGroupsGivenName = new FieldDefinition(db, "ECM UserGroups:GivenName");
            FieldDefinition fd_ECMUserGroupsOrgEmail = new FieldDefinition(db, "ECM UserGroups:OrgEmail");
            FieldDefinition fd_ECMUserGroupsSurname = new FieldDefinition(db, "ECM UserGroups:Surname");
            FieldDefinition fd_ECMVolumeFilename = new FieldDefinition(db, "ECM Volume:Filename");
            FieldDefinition fd_ECMVolumeLastModified = new FieldDefinition(db, "ECM Volume:LastModified");
            FieldDefinition fd_ECMVolumeMedia = new FieldDefinition(db, "ECM Volume:Media");
            FieldDefinition fd_ECMVolumePrimaryCacheName = new FieldDefinition(db, "ECM Volume:PrimaryCacheName");
            FieldDefinition fd_ECMVolumeRenditionTypeName = new FieldDefinition(db, "ECM Volume:RenditionTypeName");
            FieldDefinition fd_ECMVolumeShareDrive = new FieldDefinition(db, "ECM Volume:ShareDrive");
            FieldDefinition fd_ECMVolumeSize = new FieldDefinition(db, "ECM Volume:Size");
            FieldDefinition fd_ECMVolumeStatus = new FieldDefinition(db, "ECM Volume:Status");
            FieldDefinition fd_ECMVolumeStorageLocation = new FieldDefinition(db, "ECM Volume:StorageLocation");
            FieldDefinition fd_ECMVolumeUpdatable = new FieldDefinition(db, "ECM Volume:Updatable");
        }
        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\ECMMigration\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\ECMMigration\");
            }
            FileStream fs = new FileStream(@"C:\Temp\ECMMigration\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\ECMMigration\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
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
        private static string FindMappedStorage(string eCMVolumeStorageLocation)
        {
            string skjfdh = null;
            if (eCMVolumeStorageLocation.Length == 21)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 11);
            }
            else if (eCMVolumeStorageLocation.Length == 22)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 12);
            }

            switch (skjfdh)
            {
                case "ECMIncache2":
                    skjfdh = @"T:\ECMIncache2\INCACHE";
                    break;
                case "ECMIncache3":
                    skjfdh = @"Y:\ECMIncache3\INCACHE";
                    break;
                case "ECMIncache4":
                    skjfdh = @"Y:\ECMIncache4\INCACHE";
                    break;
                case "ECMIncache5":
                    skjfdh = @"Y:\ECMIncache5\INCACHE";
                    break;
                case "ECMIncache6":
                    skjfdh = @"Y:\ECMIncache6\INCACHE";
                    break;
                case "ECMIncache7":
                    skjfdh = @"Y:\ECMIncache7\INCACHE";
                    break;
                case "ECMIncache8":
                    skjfdh = @"Y:\ECMIncache8\INCACHE";
                    break;
                case "ECMIncache10":
                    skjfdh = @"W:\ECMIncache10\INCACHE";
                    break;
                case "ECMIncache11":
                    skjfdh = @"W:\ECMIncache11\INCACHE";
                    break;
                case "ECMIncache12":
                    skjfdh = @"V:\ECMIncache12\INCACHE";
                    break;
                case "ECMIncache13":
                    skjfdh = @"V:\ECMIncache13\INCACHE";
                    break;
                case "ECMIncache14":
                    skjfdh = @"U:\ECMIncache14\INCACHE";
                    break;
                case "ECMIncache15":
                    skjfdh = @"U:\ECMIncache15\INCACHE";
                    break;
                case "ECMIncache17":
                    skjfdh = @"T:\ECMIncache17\INCACHE";
                    break;
                case "ECMIncache18":
                    skjfdh = @"T:\ECMIncache18\INCACHE";
                    break;

            }


            return skjfdh;


        }
        private static Record GetFolderUri(string function, string activity, Database db)
        {
            Record r = null;
            //For tfr to SCC
            //zz legacy T1 ECM - Business Folders - Information Management - Projects
            //Local
            //zz Legacy T1 ECM - Business Folders
            string ClassSearch = "zz legacy T1 ECM - Business Folders - " + function;
            Classification c = (Classification)db.FindTrimObjectByName(BaseObjectTypes.Classification, "zz legacy T1 ECM - Business Folders - " + function);
            if (c != null)
            {
                TrimMainObjectSearch mso = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                //string ss = "title:" + reader.GetString(7) + " and classification:zz Legacy T1 ECM - Business Folders - " + reader.GetString(6);
                mso.SetSearchString("classification:" + c.Uri);
                foreach (Record rec in mso)
                {
                    if (rec.Title == activity)
                    {
                        r = rec;
                    }
                }
            }
            return r;
        }
    }
}
