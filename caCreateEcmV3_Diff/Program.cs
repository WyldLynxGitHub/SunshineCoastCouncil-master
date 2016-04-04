using HP.HPTRIM.SDK;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace caCreateEcmV3_Diff
{
    class Program
    {
        public static string strCon;
        public static bool bDetailMsg = false;
        public static string strOffset = null;
        public static string strFetch = null;
        class ECMMigration
        {
            public string DocSetID { get; set; }
            //1
            public string ECMDescription { get; set; }
            public string ECMBCSFunction_Name { get; set; }
            public string ECMBCSActivity_Name { get; set; }
            public string ECMBCSSubject_Name { get; set; }
            public int ECMSTDVersion { get; set; }
            //Auto Create from here
            public String ECMApplicationRAM_Application_Number { get; set; }
            public String ECMBarCodeBarCodeString { get; set; }
            public String ECMBCSRetention_Period { get; set; }
            public String ECMCaseCaseDescription { get; set; }
            public String ECMCaseCaseName { get; set; }
            public String ECMCaseCaseNumber { get; set; }
            public String ECMCaseCaseOfficer { get; set; }
            public String ECMCaseCaseStatus { get; set; }
            public String ECMCaseCaseType { get; set; }
            public Nullable<DateTime> ECMCaseEndDate { get; set; }
            public Nullable<DateTime> ECMCaseStartDate { get; set; }
            public String ECMCorrespondentDescription { get; set; }
            public int ECMECMDriveID { get; set; }
            public Nullable<DateTime> ECMEmployeeCommencement_Date { get; set; }
            public String ECMEmployeeEmployee_Given_Name { get; set; }
            public String ECMEmployeeEmployee_Surname { get; set; }
            public String ECMEmployeeEmployeeCode { get; set; }
            public String ECmEmployeeHR_Case_Description { get; set; }
            public String ECMEmployeePreferred_Name { get; set; }
            public String ECMEmployeeStatus { get; set; }
            public Nullable<DateTime> ECMEmployeeTermination_Date { get; set; }
            public String ECMFileName { get; set; }
            public String ECMFileNetFoldername { get; set; }
            public String ECMHR_Case_Code { get; set; }
            public String ECMInfoExpertLevel1_Name { get; set; }
            public String ECMInfoExpertLevel2_Name { get; set; }
            public String ECMInfoExpertLevel3_Description { get; set; }
            public String ECMInfoExpertLevel3_Name { get; set; }
            public String ECMInfoExpertLevel4_Description { get; set; }
            public String ECMInfoExpertLevel4_Name { get; set; }
            public String ECMInfoExpertLevel5_Description { get; set; }
            public String ECMInfoExpertLevel5_Name { get; set; }
            public String ECMInfringementInfringementID { get; set; }
            public String ECMNotes { get; set; }
            public Nullable<DateTime> ECMPositionClose_Date { get; set; }
            public Nullable<DateTime> ECMPositionOpen_Date { get; set; }
            public String ECMPositionPosition_ID { get; set; }
            public String ECMPositionPosition_Title { get; set; }
            public String ECMPositionVacancy_ID { get; set; }
            public String ECMPositionVacancy_Status { get; set; }
            public String ECMProjectAndContractProject_Name { get; set; }
            public String ECMProjectAndContractContract_or_Activity_Desc { get; set; }
            public String ECMProjectAndContractContract_or_Activity_Name { get; set; }
            public String ECMProjectAndContractContract_or_Activity_No { get; set; }
            public Nullable<DateTime> ECMProjectAndContractEnd_Date { get; set; }
            public String ECMProjectAndContractProject_Description { get; set; }
            public Nullable<DateTime> ECMProjectAndContractProject_End_Date { get; set; }
            public String ECMProjectAndContractProject_Officer { get; set; }
            public Nullable<DateTime> ECMProjectAndContractProject_Start_Date { get; set; }
            public String ECMProjectAndContractProjectNo { get; set; }
            public String ECMProjectAndContractResponsible_Officer { get; set; }
            public Nullable<DateTime> ECMProjectAndContractStart_Date { get; set; }
            public String ECMProjectAndContractStatus { get; set; }
            public String ECMPropertyHouse_No { get; set; }
            public String ECMPropertyHouse_No_Suffix { get; set; }
            public String ECMPropertyHouse_No_To { get; set; }
            public String ECMPropertyHouse_No_To_Suffix { get; set; }
            public String ECMPropertyLocality_Name { get; set; }
            public String ECMPropertyPostcode { get; set; }
            public String ECMPropertyProperty_Name { get; set; }
            public String ECMPropertyProperty_No { get; set; }
            public String ECMPropertyStreet_Name { get; set; }
            public String ECMPropertyUnit_No { get; set; }
            public String ECMPropertyUnit_No_Suffix { get; set; }
            public int ECMreference { get; set; }
            public int ECMRelatedDocDocSetID { get; set; }
            public String ECMSTDAllRevTitle { get; set; }
            public String ECMSTDApplicationNo { get; set; }
            public String ECMSTDBusinessCode { get; set; }
            public String ECMSTDClassName { get; set; }
            public String ECMSTDCorrespondant { get; set; }
            public String ECMSTDCustomerRequest { get; set; }
            public Nullable<DateTime> ECMSTDDateLastAccessed { get; set; }
            public Nullable<DateTime> ECMSTDDateReceived { get; set; }
            public Nullable<DateTime> ECMSTDDateRegistered { get; set; }
            public Nullable<DateTime> ECMSTDDeclaredDate { get; set; }
            public Nullable<DateTime> ECMSTDDestructionDue { get; set; }
            public Nullable<DateTime> ECMSTDDocumentDate { get; set; }
            public String ECMSTDDocumentType { get; set; }
            public String ECMSTDExternalReference { get; set; }
            public String ECMSTDInfringementNo { get; set; }
            public String ECMSTDInternalReference { get; set; }
            public String ECMSTDInternalScanRequest { get; set; }
            public String ECMSTDJobNo { get; set; }
            public String ECMSTDOtherReferences { get; set; }
            public String ECMSTDPropertyNo { get; set; }
            public int ECMSTDSummaryText { get; set; }
            public String ECMStreetsLocality { get; set; }
            public String ECMStreetsPostcode { get; set; }
            public String ECMStreetsStreet_Name { get; set; }
            public String ECMUserGroupsDescription { get; set; }
            public String ECMUserGroupsExtNo { get; set; }
            public String ECMUserGroupsGivenName { get; set; }
            public String ECMUserGroupsOrgEmail { get; set; }
            public String ECMUserGroupsSurname { get; set; }
            public String ECMVolumeFilename { get; set; }
            public Nullable<DateTime> ECMVolumeLastModified { get; set; }
            public String ECMVolumeMedia { get; set; }
            public String ECMVolumePrimaryCacheName { get; set; }
            public String ECMVolumeRenditionTypeName { get; set; }
            public String ECMVolumeShareDrive { get; set; }
            public int ECMVolumeSize { get; set; }
            public String ECMVolumeStatus { get; set; }
            public String ECMVolumeStorageLocation { get; set; }
            public bool ECMVolumeUpdatable { get; set; }
        }
        static void Main(string[] args)
        {
            strCon = "Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True";
            CreateNewrecords();
        }
        private static void CreateNewrecords()
        {
            List<ECMMigration> lstecm = new List<ECMMigration>();
            //Console.WriteLine("Enter Connection string");
            using (SqlConnection con = new SqlConnection("Data Source=Cl1025;Initial Catalog=StagingMove;Integrated Security=True"))
            {
                con.Open();
                Console.WriteLine("Connected to SQL 1025");
                Console.WriteLine("Export Batch?");
                string strBatch = Console.ReadLine();
                Console.WriteLine("Offset?");
                strOffset = Console.ReadLine();
                Console.WriteLine("Fetch?");
                strFetch = Console.ReadLine();
                Console.WriteLine("Show detailed messages:");
                switch (Console.ReadLine())
                {
                    case "y":
                        bDetailMsg = true;
                        break;
                    case "n":
                        bDetailMsg = false;
                        break;
                }
                string strsql = null;
                switch (strBatch)
                {
                    case "ep":
                        strsql = "Select * from [Extract Public DIFF] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Public DIFF] GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] in (Select DocumentSetID from ECM_MasterCheck Where OrigBatch = 'Extract Public' and rmuri is null and Pass = 2) Order by [ID] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                        break;
                    case "epz":
                        strsql = "Select * from [Extract Public Zip DIFF] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Public Zip DIFF] GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] in (Select DocumentSetID from ECM_MasterCheck Where OrigBatch = 'Extract Public Zip' and rmuri is null and Pass = 2) Order by [ID] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                        break;
                    case "elai":
                        strsql = "Select * from [Extract Linked To Application DIFF] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Linked To Application DIFF] GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] in (Select DocumentSetID from ECM_MasterCheck Where OrigBatch = 'Linked To Application' and rmuri is null and Pass = 2) Order by [ID] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                        break;
                    case "elpi":
                        strsql = "Select * from [Extract Linked To Property DIFF] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Linked To Property DIFF] GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] in (Select DocumentSetID from ECM_MasterCheck Where OrigBatch = 'Linked To Property' and rmuri is null and Pass = 2) Order by [ID] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                        break;
                    case "eri":
                        strsql = "Select * from [Extract Remainder DIFF] Where ID in(SELECT MIN(ID) AS ID FROM [Extract Remainder DIFF] GROUP BY [STD:DocumentID]) and [STD:Version] = 1 and [STD:DocumentSetID] in (Select DocumentSetID from ECM_MasterCheck Where OrigBatch = 'Extract Remainder' and rmuri is null and Pass = 2) Order by [ID] OFFSET " + strOffset + " ROWS FETCH NEXT " + strFetch + " ROWS ONLY";
                        break;
                }
                if (strsql == null)
                {
                    Console.WriteLine("SQL is blank");
                    Console.ReadLine();
                    return;
                }
                using (SqlCommand command = new SqlCommand(strsql, con))
                {
                    Console.WriteLine("Query connected");
                    DataSet dataSet = new DataSet();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = command;
                    da.Fill(dt);
                    
                    SqlDataReader reader = command.ExecuteReader();

                    int ECMSTDDocumentSetID = reader.GetOrdinal("STD:DocumentSetID");
                    int ECMSTDDocumentDescription = reader.GetOrdinal("STD:DocumentDescription");
                    int ECMBCSFunction_Name = reader.GetOrdinal("BCS:Function_Name");
                    int ECMBCSActivity_Name = reader.GetOrdinal("BCS:Activity_Name");
                    int ECMBCSSubject_Name = reader.GetOrdinal("BCS:Subject_Name");
                    int ECMSTDVersion = reader.GetOrdinal("STD:Version");
                    //BCS: Activity_Name

                    //Auto Create
                    int ECMApplicationRAM_Application_Number = reader.GetOrdinal("Application:RAM_Application_Number");
                    int ECMBarCodeBarCodeString = reader.GetOrdinal("BarCode:BarCodeString");
                    int ECMBCSRetention_Period = reader.GetOrdinal("BCS:Retention_Period");
                    int ECMCaseCaseDescription = reader.GetOrdinal("Case:CaseDescription");
                    int ECMCaseCaseName = reader.GetOrdinal("Case:CaseName");
                    int ECMCaseCaseNumber = reader.GetOrdinal("Case:CaseNumber");
                    int ECMCaseCaseOfficer = reader.GetOrdinal("Case:CaseOfficer");
                    int ECMCaseCaseStatus = reader.GetOrdinal("Case:CaseStatus");
                    int ECMCaseCaseType = reader.GetOrdinal("Case:CaseType");
                    int ECMCaseEndDate = reader.GetOrdinal("Case:EndDate");
                    int ECMCaseStartDate = reader.GetOrdinal("Case:StartDate");
                    //int ECMCorrespondentDescription = reader.GetOrdinal("Correspondent:Description");
                    int ECMECMDriveID = reader.GetOrdinal("ECMDriveID");
                    int ECMEmployeeCommencement_Date = reader.GetOrdinal("Employee:Commencement_Date");
                    int ECMEmployeeEmployee_Given_Name = reader.GetOrdinal("Employee:Employee_Given_Name");
                    int ECMEmployeeEmployee_Surname = reader.GetOrdinal("Employee:Employee_Surname");
                    int ECMEmployeeEmployeeCode = reader.GetOrdinal("Employee:EmployeeCode");
                    int ECmEmployeeHR_Case_Description = reader.GetOrdinal("Employee:HR_Case_Description");
                    int ECMEmployeePreferred_Name = reader.GetOrdinal("Employee:Preferred_Name");
                    int ECMEmployeeStatus = reader.GetOrdinal("Employee:Status");
                    int ECMEmployeeTermination_Date = reader.GetOrdinal("Employee:Termination_Date");
                    //int ECMFileName = reader.GetOrdinal("FileName");
                    int ECMFileNetFoldername = reader.GetOrdinal("FileNet:Foldername");
                    int ECMHR_Case_Code = reader.GetOrdinal("HR_Case_Code");
                    int ECMInfoExpertLevel1_Name = reader.GetOrdinal("InfoExpert:Level1_Name");
                    int ECMInfoExpertLevel2_Name = reader.GetOrdinal("InfoExpert:Level2_Name");
                    int ECMInfoExpertLevel3_Description = reader.GetOrdinal("InfoExpert:Level3_Description");
                    int ECMInfoExpertLevel3_Name = reader.GetOrdinal("InfoExpert:Level3_Name");
                    int ECMInfoExpertLevel4_Description = reader.GetOrdinal("InfoExpert:Level4_Description");
                    int ECMInfoExpertLevel4_Name = reader.GetOrdinal("InfoExpert:Level4_Name");
                    int ECMInfoExpertLevel5_Description = reader.GetOrdinal("InfoExpert:Level5_Description");
                    int ECMInfoExpertLevel5_Name = reader.GetOrdinal("InfoExpert:Level5_Name");
                    int ECMInfringementInfringementID = reader.GetOrdinal("Infringement:InfringementID");
                    int ECMNotes = reader.GetOrdinal("Notes");
                    int ECMPositionClose_Date = reader.GetOrdinal("Position:Close_Date");
                    int ECMPositionOpen_Date = reader.GetOrdinal("Position:Open_Date");
                    int ECMPositionPosition_ID = reader.GetOrdinal("Position:Position_ID");
                    int ECMPositionPosition_Title = reader.GetOrdinal("Position:Position_Title");
                    int ECMPositionVacancy_ID = reader.GetOrdinal("Position:Vacancy_ID");
                    int ECMPositionVacancy_Status = reader.GetOrdinal("Position:Vacancy_Status");
                    int ECMProjectAndContractProject_Name = reader.GetOrdinal("ProjectAndContract:[Project_Name");
                    int ECMProjectAndContractContract_or_Activity_Desc = reader.GetOrdinal("ProjectAndContract:Contract_or_Activity_Desc");
                    int ECMProjectAndContractContract_or_Activity_Name = reader.GetOrdinal("ProjectAndContract:Contract_or_Activity_Name");
                    int ECMProjectAndContractContract_or_Activity_No = reader.GetOrdinal("ProjectAndContract:Contract_or_Activity_No");
                    int ECMProjectAndContractEnd_Date = reader.GetOrdinal("ProjectAndContract:End_Date");
                    int ECMProjectAndContractProject_Description = reader.GetOrdinal("ProjectAndContract:Project_Description");
                    int ECMProjectAndContractProject_End_Date = reader.GetOrdinal("ProjectAndContract:Project_End_Date");
                    int ECMProjectAndContractProject_Officer = reader.GetOrdinal("ProjectAndContract:Project_Officer");
                    int ECMProjectAndContractProject_Start_Date = reader.GetOrdinal("ProjectAndContract:Project_Start_Date");
                    int ECMProjectAndContractProjectNo = reader.GetOrdinal("ProjectAndContract:ProjectNo");
                    int ECMProjectAndContractResponsible_Officer = reader.GetOrdinal("ProjectAndContract:Responsible_Officer");
                    int ECMProjectAndContractStart_Date = reader.GetOrdinal("ProjectAndContract:Start_Date");
                    int ECMProjectAndContractStatus = reader.GetOrdinal("ProjectAndContract:Status");
                    int ECMPropertyHouse_No = reader.GetOrdinal("Property:House_No");
                    int ECMPropertyHouse_No_Suffix = reader.GetOrdinal("Property:House_No_Suffix");
                    int ECMPropertyHouse_No_To = reader.GetOrdinal("Property:House_No_To");
                    int ECMPropertyHouse_No_To_Suffix = reader.GetOrdinal("Property:House_No_To_Suffix");
                    int ECMPropertyLocality_Name = reader.GetOrdinal("Property:Locality_Name");
                    int ECMPropertyPostcode = reader.GetOrdinal("Property:Postcode");
                    int ECMPropertyProperty_Name = reader.GetOrdinal("Property:Property_Name");
                    int ECMPropertyProperty_No = reader.GetOrdinal("Property:Property_No");
                    int ECMPropertyStreet_Name = reader.GetOrdinal("Property:Street_Name");
                    int ECMPropertyUnit_No = reader.GetOrdinal("Property:Unit_No");
                    int ECMPropertyUnit_No_Suffix = reader.GetOrdinal("Property:Unit_No_Suffix");
                    //int ECMreference = reader.GetOrdinal("reference");
                    int ECMRelatedDocDocSetID = reader.GetOrdinal("RelatedDoc:DocSetID");
                    int ECMSTDAllRevTitle = reader.GetOrdinal("STD:AllRevTitle");
                    int ECMSTDApplicationNo = reader.GetOrdinal("STD:ApplicationNo");
                    int ECMSTDBusinessCode = reader.GetOrdinal("STD:BusinessCode");
                    int ECMSTDClassName = reader.GetOrdinal("STD:ClassName");
                    int ECMSTDCorrespondant = reader.GetOrdinal("STD:Correspondant");
                    int ECMSTDCustomerRequest = reader.GetOrdinal("STD:CustomerRequest");
                    int ECMSTDDateLastAccessed = reader.GetOrdinal("STD:DateLastAccessed");
                    int ECMSTDDateReceived = reader.GetOrdinal("STD:DateReceived");
                    int ECMSTDDateRegistered = reader.GetOrdinal("STD:DateRegistered");
                    int ECMSTDDeclaredDate = reader.GetOrdinal("STD:DeclaredDate");
                    int ECMSTDDestructionDue = reader.GetOrdinal("STD:DestructionDue");
                    int ECMSTDDocumentDate = reader.GetOrdinal("STD:DocumentDate");
                    int ECMSTDDocumentType = reader.GetOrdinal("STD:DocumentType");
                    int ECMSTDExternalReference = reader.GetOrdinal("STD:ExternalReference");
                    int ECMSTDInfringementNo = reader.GetOrdinal("STD:InfringementNo");
                    int ECMSTDInternalReference = reader.GetOrdinal("STD:InternalReference");
                    int ECMSTDInternalScanRequest = reader.GetOrdinal("STD:InternalScanRequest");
                    int ECMSTDJobNo = reader.GetOrdinal("STD:JobNo");
                    int ECMSTDOtherReferences = reader.GetOrdinal("STD:OtherReferences");
                    int ECMSTDPropertyNo = reader.GetOrdinal("STD:PropertyNo");
                    int ECMSTDSummaryText = reader.GetOrdinal("STD:SummaryText");
                    int ECMStreetsLocality = reader.GetOrdinal("Streets:Locality");
                    int ECMStreetsPostcode = reader.GetOrdinal("Streets:Postcode");
                    int ECMStreetsStreet_Name = reader.GetOrdinal("Streets:Street_Name");
                    int ECMUserGroupsDescription = reader.GetOrdinal("UserGroups:Description");
                    int ECMUserGroupsExtNo = reader.GetOrdinal("UserGroups:ExtNo");
                    int ECMUserGroupsGivenName = reader.GetOrdinal("UserGroups:GivenName");
                    int ECMUserGroupsOrgEmail = reader.GetOrdinal("UserGroups:OrgEmail");
                    int ECMUserGroupsSurname = reader.GetOrdinal("UserGroups:Surname");
                    int ECMVolumeFilename = reader.GetOrdinal("Volume:Filename");
                    int ECMVolumeLastModified = reader.GetOrdinal("Volume:LastModified");
                    int ECMVolumeMedia = reader.GetOrdinal("Volume:Media");
                    int ECMVolumePrimaryCacheName = reader.GetOrdinal("Volume:PrimaryCacheName");
                    int ECMVolumeRenditionTypeName = reader.GetOrdinal("Volume:RenditionTypeName");
                    int ECMVolumeShareDrive = reader.GetOrdinal("Volume:ShareDrive");
                    int ECMVolumeSize = reader.GetOrdinal("Volume:Size");
                    int ECMVolumeStatus = reader.GetOrdinal("Volume:Status");
                    int ECMVolumeStorageLocation = reader.GetOrdinal("Volume:StorageLocation");
                    int ECMVolumeUpdatable = reader.GetOrdinal("Volume:Updatable");
                    //
                    while (reader.Read())
                    {
                        //Console.WriteLine(reader.GetString(8));
                        //Console.WriteLine("ECMVolumeFilename={0}", reader.GetString(ECMVolumeFilename));
                        ECMMigration ecm = new ECMMigration();
                        ecm.DocSetID = reader.GetInt32(ECMSTDDocumentSetID).ToString();
                        ecm.ECMDescription = reader.GetString(ECMSTDDocumentDescription);
                        ecm.ECMBCSFunction_Name = reader.GetString(ECMBCSFunction_Name);
                        ecm.ECMBCSActivity_Name = reader.GetString(ECMBCSActivity_Name);
                        ecm.ECMBCSSubject_Name = reader.GetString(ECMBCSSubject_Name);
                        ecm.ECMSTDVersion = reader.GetInt16(ECMSTDVersion);
                        //AutoCreate
                        try { if (reader.IsDBNull(ECMApplicationRAM_Application_Number) == false) { ecm.ECMApplicationRAM_Application_Number = reader.GetString(ECMApplicationRAM_Application_Number); } } catch { Console.WriteLine("Data convert problem with ECMApplicationRAM_Application_Number"); }
                        try { if (reader.IsDBNull(ECMBarCodeBarCodeString) == false) { ecm.ECMBarCodeBarCodeString = reader.GetString(ECMBarCodeBarCodeString); } } catch { Console.WriteLine("Data convert problem with ECMBarCodeBarCodeString"); }
                        try { if (reader.IsDBNull(ECMBCSRetention_Period) == false) { ecm.ECMBCSRetention_Period = reader.GetString(ECMBCSRetention_Period); } } catch { Console.WriteLine("Data convert problem with ECMBCSRetention_Period"); }
                        try { if (reader.IsDBNull(ECMCaseCaseDescription) == false) { ecm.ECMCaseCaseDescription = reader.GetString(ECMCaseCaseDescription); } } catch { Console.WriteLine("Data convert problem with ECMCaseCaseDescription"); }
                        try { if (reader.IsDBNull(ECMCaseCaseName) == false) { ecm.ECMCaseCaseName = reader.GetString(ECMCaseCaseName); } } catch { Console.WriteLine("Data convert problem with ECMCaseCaseName"); }
                        try { if (reader.IsDBNull(ECMCaseCaseNumber) == false) { ecm.ECMCaseCaseNumber = reader.GetString(ECMCaseCaseNumber); } } catch { Console.WriteLine("Data convert problem with ECMCaseCaseNumber"); }
                        try { if (reader.IsDBNull(ECMCaseCaseOfficer) == false) { ecm.ECMCaseCaseOfficer = reader.GetString(ECMCaseCaseOfficer); } } catch { Console.WriteLine("Data convert problem with ECMCaseCaseOfficer"); }
                        try { if (reader.IsDBNull(ECMCaseCaseStatus) == false) { ecm.ECMCaseCaseStatus = reader.GetString(ECMCaseCaseStatus); } } catch { Console.WriteLine("Data convert problem with ECMCaseCaseStatus"); }
                        try { if (reader.IsDBNull(ECMCaseCaseType) == false) { ecm.ECMCaseCaseType = reader.GetString(ECMCaseCaseType); } } catch { Console.WriteLine("Data convert problem with ECMCaseCaseType"); }
                        try { if (reader.IsDBNull(ECMCaseEndDate) == false) { ecm.ECMCaseEndDate = reader.GetDateTime(ECMCaseEndDate); } } catch { Console.WriteLine("Data convert problem with ECMCaseEndDate"); }
                        try { if (reader.IsDBNull(ECMCaseStartDate) == false) { ecm.ECMCaseStartDate = reader.GetDateTime(ECMCaseStartDate); } } catch { Console.WriteLine("Data convert problem with ECMCaseStartDate"); }
                        //try { if (reader.IsDBNull(ECMCorrespondentDescription) == false) { ecm.ECMCorrespondentDescription = reader.GetString(ECMCorrespondentDescription); } } catch { Console.WriteLine("Data convert problem with ECMCorrespondentDescription"); }
                        try { if (reader.IsDBNull(ECMECMDriveID) == false) { ecm.ECMECMDriveID = reader.GetInt32(ECMECMDriveID); } } catch { Console.WriteLine("Data convert problem with ECMECMDriveID"); }
                        try { if (reader.IsDBNull(ECMEmployeeCommencement_Date) == false) { ecm.ECMEmployeeCommencement_Date = reader.GetDateTime(ECMEmployeeCommencement_Date); } } catch { Console.WriteLine("Data convert problem with ECMEmployeeCommencement_Date"); }
                        try { if (reader.IsDBNull(ECMEmployeeEmployee_Given_Name) == false) { ecm.ECMEmployeeEmployee_Given_Name = reader.GetString(ECMEmployeeEmployee_Given_Name); } } catch { Console.WriteLine("Data convert problem with ECMEmployeeEmployee_Given_Name"); }
                        try { if (reader.IsDBNull(ECMEmployeeEmployee_Surname) == false) { ecm.ECMEmployeeEmployee_Surname = reader.GetString(ECMEmployeeEmployee_Surname); } } catch { Console.WriteLine("Data convert problem with ECMEmployeeEmployee_Surname"); }
                        try { if (reader.IsDBNull(ECMEmployeeEmployeeCode) == false) { ecm.ECMEmployeeEmployeeCode = reader.GetString(ECMEmployeeEmployeeCode); } } catch { Console.WriteLine("Data convert problem with ECMEmployeeEmployeeCode"); }
                        try { if (reader.IsDBNull(ECmEmployeeHR_Case_Description) == false) { ecm.ECmEmployeeHR_Case_Description = reader.GetString(ECmEmployeeHR_Case_Description); } } catch { Console.WriteLine("Data convert problem with ECmEmployeeHR_Case_Description"); }
                        try { if (reader.IsDBNull(ECMEmployeePreferred_Name) == false) { ecm.ECMEmployeePreferred_Name = reader.GetString(ECMEmployeePreferred_Name); } } catch { Console.WriteLine("Data convert problem with ECMEmployeePreferred_Name"); }
                        try { if (reader.IsDBNull(ECMEmployeeStatus) == false) { ecm.ECMEmployeeStatus = reader.GetString(ECMEmployeeStatus); } } catch { Console.WriteLine("Data convert problem with ECMEmployeeStatus"); }
                        try { if (reader.IsDBNull(ECMEmployeeTermination_Date) == false) { ecm.ECMEmployeeTermination_Date = reader.GetDateTime(ECMEmployeeTermination_Date); } } catch { Console.WriteLine("Data convert problem with ECMEmployeeTermination_Date"); }
                        //try { if (reader.IsDBNull(ECMFileName) == false) { ecm.ECMFileName = reader.GetString(ECMFileName); } } catch { Console.WriteLine("Data convert problem with ECMFileName"); }
                        try { if (reader.IsDBNull(ECMFileNetFoldername) == false) { ecm.ECMFileNetFoldername = reader.GetString(ECMFileNetFoldername); } } catch { Console.WriteLine("Data convert problem with ECMFileNetFoldername"); }
                        try { if (reader.IsDBNull(ECMHR_Case_Code) == false) { ecm.ECMHR_Case_Code = reader.GetString(ECMHR_Case_Code); } } catch { Console.WriteLine("Data convert problem with ECMHR_Case_Code"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel1_Name) == false) { ecm.ECMInfoExpertLevel1_Name = reader.GetString(ECMInfoExpertLevel1_Name); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel1_Name"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel2_Name) == false) { ecm.ECMInfoExpertLevel2_Name = reader.GetString(ECMInfoExpertLevel2_Name); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel2_Name"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel3_Description) == false) { ecm.ECMInfoExpertLevel3_Description = reader.GetString(ECMInfoExpertLevel3_Description); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel3_Description"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel3_Name) == false) { ecm.ECMInfoExpertLevel3_Name = reader.GetString(ECMInfoExpertLevel3_Name); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel3_Name"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel4_Description) == false) { ecm.ECMInfoExpertLevel4_Description = reader.GetString(ECMInfoExpertLevel4_Description); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel4_Description"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel4_Name) == false) { ecm.ECMInfoExpertLevel4_Name = reader.GetString(ECMInfoExpertLevel4_Name); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel4_Name"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel5_Description) == false) { ecm.ECMInfoExpertLevel5_Description = reader.GetString(ECMInfoExpertLevel5_Description); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel5_Description"); }
                        try { if (reader.IsDBNull(ECMInfoExpertLevel5_Name) == false) { ecm.ECMInfoExpertLevel5_Name = reader.GetString(ECMInfoExpertLevel5_Name); } } catch { Console.WriteLine("Data convert problem with ECMInfoExpertLevel5_Name"); }
                        try { if (reader.IsDBNull(ECMInfringementInfringementID) == false) { ecm.ECMInfringementInfringementID = reader.GetString(ECMInfringementInfringementID); } } catch { Console.WriteLine("Data convert problem with ECMInfringementInfringementID"); }
                        try { if (reader.IsDBNull(ECMNotes) == false) { ecm.ECMNotes = reader.GetString(ECMNotes); } } catch { Console.WriteLine("Data convert problem with ECMNotes"); }
                        try { if (reader.IsDBNull(ECMPositionClose_Date) == false) { ecm.ECMPositionClose_Date = reader.GetDateTime(ECMPositionClose_Date); } } catch { Console.WriteLine("Data convert problem with ECMPositionClose_Date"); }
                        try { if (reader.IsDBNull(ECMPositionOpen_Date) == false) { ecm.ECMPositionOpen_Date = reader.GetDateTime(ECMPositionOpen_Date); } } catch { Console.WriteLine("Data convert problem with ECMPositionOpen_Date"); }
                        try { if (reader.IsDBNull(ECMPositionPosition_ID) == false) { ecm.ECMPositionPosition_ID = reader.GetString(ECMPositionPosition_ID); } } catch { Console.WriteLine("Data convert problem with ECMPositionPosition_ID"); }
                        try { if (reader.IsDBNull(ECMPositionPosition_Title) == false) { ecm.ECMPositionPosition_Title = reader.GetString(ECMPositionPosition_Title); } } catch { Console.WriteLine("Data convert problem with ECMPositionPosition_Title"); }
                        try { if (reader.IsDBNull(ECMPositionVacancy_ID) == false) { ecm.ECMPositionVacancy_ID = reader.GetString(ECMPositionVacancy_ID); } } catch { Console.WriteLine("Data convert problem with ECMPositionVacancy_ID"); }
                        try { if (reader.IsDBNull(ECMPositionVacancy_Status) == false) { ecm.ECMPositionVacancy_Status = reader.GetString(ECMPositionVacancy_Status); } } catch { Console.WriteLine("Data convert problem with ECMPositionVacancy_Status"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractProject_Name) == false) { ecm.ECMProjectAndContractProject_Name = reader.GetString(ECMProjectAndContractProject_Name); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractProject_Name"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractContract_or_Activity_Desc) == false) { ecm.ECMProjectAndContractContract_or_Activity_Desc = reader.GetString(ECMProjectAndContractContract_or_Activity_Desc); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractContract_or_Activity_Desc"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractContract_or_Activity_Name) == false) { ecm.ECMProjectAndContractContract_or_Activity_Name = reader.GetString(ECMProjectAndContractContract_or_Activity_Name); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractContract_or_Activity_Name"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractContract_or_Activity_No) == false) { ecm.ECMProjectAndContractContract_or_Activity_No = reader.GetString(ECMProjectAndContractContract_or_Activity_No); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractContract_or_Activity_No"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractEnd_Date) == false) { ecm.ECMProjectAndContractEnd_Date = reader.GetDateTime(ECMProjectAndContractEnd_Date); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractEnd_Date"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractProject_Description) == false) { ecm.ECMProjectAndContractProject_Description = reader.GetString(ECMProjectAndContractProject_Description); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractProject_Description"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractProject_End_Date) == false) { ecm.ECMProjectAndContractProject_End_Date = reader.GetDateTime(ECMProjectAndContractProject_End_Date); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractProject_End_Date"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractProject_Officer) == false) { ecm.ECMProjectAndContractProject_Officer = reader.GetString(ECMProjectAndContractProject_Officer); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractProject_Officer"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractProject_Start_Date) == false) { ecm.ECMProjectAndContractProject_Start_Date = reader.GetDateTime(ECMProjectAndContractProject_Start_Date); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractProject_Start_Date"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractProjectNo) == false) { ecm.ECMProjectAndContractProjectNo = reader.GetString(ECMProjectAndContractProjectNo); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractProjectNo"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractResponsible_Officer) == false) { ecm.ECMProjectAndContractResponsible_Officer = reader.GetString(ECMProjectAndContractResponsible_Officer); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractResponsible_Officer"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractStart_Date) == false) { ecm.ECMProjectAndContractStart_Date = reader.GetDateTime(ECMProjectAndContractStart_Date); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractStart_Date"); }
                        try { if (reader.IsDBNull(ECMProjectAndContractStatus) == false) { ecm.ECMProjectAndContractStatus = reader.GetString(ECMProjectAndContractStatus); } } catch { Console.WriteLine("Data convert problem with ECMProjectAndContractStatus"); }
                        try { if (reader.IsDBNull(ECMPropertyHouse_No) == false) { ecm.ECMPropertyHouse_No = reader.GetString(ECMPropertyHouse_No); } } catch { Console.WriteLine("Data convert problem with ECMPropertyHouse_No"); }
                        try { if (reader.IsDBNull(ECMPropertyHouse_No_Suffix) == false) { ecm.ECMPropertyHouse_No_Suffix = reader.GetString(ECMPropertyHouse_No_Suffix); } } catch { Console.WriteLine("Data convert problem with ECMPropertyHouse_No_Suffix"); }
                        try { if (reader.IsDBNull(ECMPropertyHouse_No_To) == false) { ecm.ECMPropertyHouse_No_To = reader.GetString(ECMPropertyHouse_No_To); } } catch { Console.WriteLine("Data convert problem with ECMPropertyHouse_No_To"); }
                        try { if (reader.IsDBNull(ECMPropertyHouse_No_To_Suffix) == false) { ecm.ECMPropertyHouse_No_To_Suffix = reader.GetString(ECMPropertyHouse_No_To_Suffix); } } catch { Console.WriteLine("Data convert problem with ECMPropertyHouse_No_To_Suffix"); }
                        try { if (reader.IsDBNull(ECMPropertyLocality_Name) == false) { ecm.ECMPropertyLocality_Name = reader.GetString(ECMPropertyLocality_Name); } } catch { Console.WriteLine("Data convert problem with ECMPropertyLocality_Name"); }
                        try { if (reader.IsDBNull(ECMPropertyPostcode) == false) { ecm.ECMPropertyPostcode = reader.GetString(ECMPropertyPostcode); } } catch { Console.WriteLine("Data convert problem with ECMPropertyPostcode"); }
                        try { if (reader.IsDBNull(ECMPropertyProperty_Name) == false) { ecm.ECMPropertyProperty_Name = reader.GetString(ECMPropertyProperty_Name); } } catch { Console.WriteLine("Data convert problem with ECMPropertyProperty_Name"); }
                        try { if (reader.IsDBNull(ECMPropertyProperty_No) == false) { ecm.ECMPropertyProperty_No = reader.GetString(ECMPropertyProperty_No); } } catch { Console.WriteLine("Data convert problem with ECMPropertyProperty_No"); }
                        try { if (reader.IsDBNull(ECMPropertyStreet_Name) == false) { ecm.ECMPropertyStreet_Name = reader.GetString(ECMPropertyStreet_Name); } } catch { Console.WriteLine("Data convert problem with ECMPropertyStreet_Name"); }
                        try { if (reader.IsDBNull(ECMPropertyUnit_No) == false) { ecm.ECMPropertyUnit_No = reader.GetString(ECMPropertyUnit_No); } } catch { Console.WriteLine("Data convert problem with ECMPropertyUnit_No"); }
                        try { if (reader.IsDBNull(ECMPropertyUnit_No_Suffix) == false) { ecm.ECMPropertyUnit_No_Suffix = reader.GetString(ECMPropertyUnit_No_Suffix); } } catch { Console.WriteLine("Data convert problem with ECMPropertyUnit_No_Suffix"); }
                        //try { if (reader.IsDBNull(ECMreference) == false) { ecm.ECMreference = reader.GetInt32(ECMreference); } } catch { Console.WriteLine("Data convert problem with ECMreference"); }
                        try { if (reader.IsDBNull(ECMRelatedDocDocSetID) == false) { ecm.ECMRelatedDocDocSetID = reader.GetInt32(ECMRelatedDocDocSetID); } } catch { Console.WriteLine("Data convert problem with ECMRelatedDocDocSetID"); }
                        try { if (reader.IsDBNull(ECMSTDAllRevTitle) == false) { ecm.ECMSTDAllRevTitle = reader.GetString(ECMSTDAllRevTitle); } } catch { Console.WriteLine("Data convert problem with ECMSTDAllRevTitle"); }
                        try { if (reader.IsDBNull(ECMSTDApplicationNo) == false) { ecm.ECMSTDApplicationNo = reader.GetString(ECMSTDApplicationNo); } } catch { Console.WriteLine("Data convert problem with ECMSTDApplicationNo"); }
                        try { if (reader.IsDBNull(ECMSTDBusinessCode) == false) { ecm.ECMSTDBusinessCode = reader.GetString(ECMSTDBusinessCode); } } catch { Console.WriteLine("Data convert problem with ECMSTDBusinessCode"); }
                        try { if (reader.IsDBNull(ECMSTDClassName) == false) { ecm.ECMSTDClassName = reader.GetString(ECMSTDClassName); } } catch { Console.WriteLine("Data convert problem with ECMSTDClassName"); }
                        try { if (reader.IsDBNull(ECMSTDCorrespondant) == false) { ecm.ECMSTDCorrespondant = reader.GetString(ECMSTDCorrespondant); } } catch { Console.WriteLine("Data convert problem with ECMSTDCorrespondant"); }
                        try { if (reader.IsDBNull(ECMSTDCustomerRequest) == false) { ecm.ECMSTDCustomerRequest = reader.GetString(ECMSTDCustomerRequest); } } catch { Console.WriteLine("Data convert problem with ECMSTDCustomerRequest"); }
                        try { if (reader.IsDBNull(ECMSTDDateLastAccessed) == false) { ecm.ECMSTDDateLastAccessed = reader.GetDateTime(ECMSTDDateLastAccessed); } } catch { Console.WriteLine("Data convert problem with ECMSTDDateLastAccessed"); }
                        try { if (reader.IsDBNull(ECMSTDDateReceived) == false) { ecm.ECMSTDDateReceived = reader.GetDateTime(ECMSTDDateReceived); } } catch { Console.WriteLine("Data convert problem with ECMSTDDateReceived"); }
                        try { if (reader.IsDBNull(ECMSTDDateRegistered) == false) { ecm.ECMSTDDateRegistered = reader.GetDateTime(ECMSTDDateRegistered); } } catch { Console.WriteLine("Data convert problem with ECMSTDDateRegistered"); }
                        try { if (reader.IsDBNull(ECMSTDDeclaredDate) == false) { ecm.ECMSTDDeclaredDate = reader.GetDateTime(ECMSTDDeclaredDate); } } catch { Console.WriteLine("Data convert problem with ECMSTDDeclaredDate"); }
                        try { if (reader.IsDBNull(ECMSTDDestructionDue) == false) { ecm.ECMSTDDestructionDue = reader.GetDateTime(ECMSTDDestructionDue); } } catch { Console.WriteLine("Data convert problem with ECMSTDDestructionDue"); }
                        try { if (reader.IsDBNull(ECMSTDDocumentDate) == false) { ecm.ECMSTDDocumentDate = reader.GetDateTime(ECMSTDDocumentDate); } } catch { Console.WriteLine("Data convert problem with ECMSTDDocumentDate"); }
                        try { if (reader.IsDBNull(ECMSTDDocumentType) == false) { ecm.ECMSTDDocumentType = reader.GetString(ECMSTDDocumentType); } } catch { Console.WriteLine("Data convert problem with ECMSTDDocumentType"); }
                        try { if (reader.IsDBNull(ECMSTDExternalReference) == false) { ecm.ECMSTDExternalReference = reader.GetString(ECMSTDExternalReference); } } catch { Console.WriteLine("Data convert problem with ECMSTDExternalReference"); }
                        try { if (reader.IsDBNull(ECMSTDInfringementNo) == false) { ecm.ECMSTDInfringementNo = reader.GetString(ECMSTDInfringementNo); } } catch { Console.WriteLine("Data convert problem with ECMSTDInfringementNo"); }
                        try { if (reader.IsDBNull(ECMSTDInternalReference) == false) { ecm.ECMSTDInternalReference = reader.GetString(ECMSTDInternalReference); } } catch { Console.WriteLine("Data convert problem with ECMSTDInternalReference"); }
                        try { if (reader.IsDBNull(ECMSTDInternalScanRequest) == false) { ecm.ECMSTDInternalScanRequest = reader.GetString(ECMSTDInternalScanRequest); } } catch { Console.WriteLine("Data convert problem with ECMSTDInternalScanRequest"); }
                        try { if (reader.IsDBNull(ECMSTDJobNo) == false) { ecm.ECMSTDJobNo = reader.GetString(ECMSTDJobNo); } } catch { Console.WriteLine("Data convert problem with ECMSTDJobNo"); }
                        try { if (reader.IsDBNull(ECMSTDOtherReferences) == false) { ecm.ECMSTDOtherReferences = reader.GetString(ECMSTDOtherReferences); } } catch { Console.WriteLine("Data convert problem with ECMSTDOtherReferences"); }
                        try { if (reader.IsDBNull(ECMSTDPropertyNo) == false) { ecm.ECMSTDPropertyNo = reader.GetString(ECMSTDPropertyNo); } } catch { Console.WriteLine("Data convert problem with ECMSTDPropertyNo"); }
                        try { if (reader.IsDBNull(ECMSTDSummaryText) == false) { ecm.ECMSTDSummaryText = reader.GetInt32(ECMSTDSummaryText); } } catch { Console.WriteLine("Data convert problem with ECMSTDSummaryText"); }
                        try { if (reader.IsDBNull(ECMStreetsLocality) == false) { ecm.ECMStreetsLocality = reader.GetString(ECMStreetsLocality); } } catch { Console.WriteLine("Data convert problem with ECMStreetsLocality"); }
                        try { if (reader.IsDBNull(ECMStreetsPostcode) == false) { ecm.ECMStreetsPostcode = reader.GetString(ECMStreetsPostcode); } } catch { Console.WriteLine("Data convert problem with ECMStreetsPostcode"); }
                        try { if (reader.IsDBNull(ECMStreetsStreet_Name) == false) { ecm.ECMStreetsStreet_Name = reader.GetString(ECMStreetsStreet_Name); } } catch { Console.WriteLine("Data convert problem with ECMStreetsStreet_Name"); }
                        try { if (reader.IsDBNull(ECMUserGroupsDescription) == false) { ecm.ECMUserGroupsDescription = reader.GetString(ECMUserGroupsDescription); } } catch { Console.WriteLine("Data convert problem with ECMUserGroupsDescription"); }
                        try { if (reader.IsDBNull(ECMUserGroupsExtNo) == false) { ecm.ECMUserGroupsExtNo = reader.GetString(ECMUserGroupsExtNo); } } catch { Console.WriteLine("Data convert problem with ECMUserGroupsExtNo"); }
                        try { if (reader.IsDBNull(ECMUserGroupsGivenName) == false) { ecm.ECMUserGroupsGivenName = reader.GetString(ECMUserGroupsGivenName); } } catch { Console.WriteLine("Data convert problem with ECMUserGroupsGivenName"); }
                        try { if (reader.IsDBNull(ECMUserGroupsOrgEmail) == false) { ecm.ECMUserGroupsOrgEmail = reader.GetString(ECMUserGroupsOrgEmail); } } catch { Console.WriteLine("Data convert problem with ECMUserGroupsOrgEmail"); }
                        try { if (reader.IsDBNull(ECMUserGroupsSurname) == false) { ecm.ECMUserGroupsSurname = reader.GetString(ECMUserGroupsSurname); } } catch { Console.WriteLine("Data convert problem with ECMUserGroupsSurname"); }
                        try { if (reader.IsDBNull(ECMVolumeFilename) == false) { ecm.ECMVolumeFilename = reader.GetString(ECMVolumeFilename); } } catch { Console.WriteLine("Data convert problem with ECMVolumeFilename"); }
                        try { if (reader.IsDBNull(ECMVolumeLastModified) == false) { ecm.ECMVolumeLastModified = reader.GetDateTime(ECMVolumeLastModified); } } catch { Console.WriteLine("Data convert problem with ECMVolumeLastModified"); }
                        try { if (reader.IsDBNull(ECMVolumeMedia) == false) { ecm.ECMVolumeMedia = reader.GetString(ECMVolumeMedia); } } catch { Console.WriteLine("Data convert problem with ECMVolumeMedia"); }
                        try { if (reader.IsDBNull(ECMVolumePrimaryCacheName) == false) { ecm.ECMVolumePrimaryCacheName = reader.GetString(ECMVolumePrimaryCacheName); } } catch { Console.WriteLine("Data convert problem with ECMVolumePrimaryCacheName"); }
                        try { if (reader.IsDBNull(ECMVolumeRenditionTypeName) == false) { ecm.ECMVolumeRenditionTypeName = reader.GetString(ECMVolumeRenditionTypeName); } } catch { Console.WriteLine("Data convert problem with ECMVolumeRenditionTypeName"); }
                        try { if (reader.IsDBNull(ECMVolumeShareDrive) == false) { ecm.ECMVolumeShareDrive = reader.GetString(ECMVolumeShareDrive); } } catch { Console.WriteLine("Data convert problem with ECMVolumeShareDrive"); }
                        try { if (reader.IsDBNull(ECMVolumeSize) == false) { ecm.ECMVolumeSize = reader.GetInt32(ECMVolumeSize); } } catch { Console.WriteLine("Data convert problem with ECMVolumeSize"); }
                        try { if (reader.IsDBNull(ECMVolumeStatus) == false) { ecm.ECMVolumeStatus = reader.GetString(ECMVolumeStatus); } } catch { Console.WriteLine("Data convert problem with ECMVolumeStatus"); }
                        try { if (reader.IsDBNull(ECMVolumeStorageLocation) == false) { ecm.ECMVolumeStorageLocation = reader.GetString(ECMVolumeStorageLocation); } } catch { Console.WriteLine("Data convert problem with ECMVolumeStorageLocation"); }

                        lstecm.Add(ecm);

                    }
                }
            }
            //SQL closed
            processtoEddie(lstecm);

        }
        private static void processtoEddie(List<ECMMigration> lstecm)
        {
            Console.WriteLine("Moved into RM processing");
            string strTotal = lstecm.Count.ToString();
            int iCount = 0;
            TrimApplication.Initialize();
            //Parallel.ForEach(lstecm, new ParallelOptions { MaxDegreeOfParallelism = 6 }, (ecm) =>

            DateTime dt1 = DateTime.Now;

            using (Database db = new Database())
            {
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
                //FieldDefinition fd_ECMCorrespondentDescription = new FieldDefinition(db, "ECM Correspondent:Description");
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
                //FieldDefinition fd_ECMreference = new FieldDefinition(db, "ECM reference");
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
                //FieldDefinition fd_ECMVolumeUpdatable = new FieldDefinition(db, "ECM Volume:Updatable");
                //
                RecordType rt = new RecordType(db, 16);
                //Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId);



                //});

                foreach (ECMMigration ecm in lstecm)
                {
                    iCount++;
                    DateTime dt2 = DateTime.Now;
                    TimeSpan result = dt2 - dt1;
                    Console.WriteLine("Processing {0} - Number {1} of {2}, time since start {3} - batch offset {4}, Fetch {5} ({6})", ecm.DocSetID, iCount.ToString(), strTotal, result.TotalMinutes.ToString(), strOffset, strFetch, (Convert.ToInt32(strOffset) + iCount).ToString());
                    try
                    {
                        Record reccheck = (Record)db.FindTrimObjectByName(BaseObjectTypes.Record, ecm.DocSetID);
                        //if (reccheck == null)
                        //{
                        string strStorageLoc = FindMappedStorage(ecm.ECMVolumeStorageLocation) + "\\" + ecm.ECMVolumeFilename;
                        if (File.Exists(strStorageLoc))
                        {
                            Record reccont = GetFolderUri(ecm.ECMBCSFunction_Name + " - " + ecm.ECMBCSActivity_Name.Replace("-", "~"), ecm.ECMBCSSubject_Name, db);
                            if (reccont != null)
                            {
                                string fTitle;
                                if (ecm.ECMDescription == null)
                                {
                                    fTitle = "Record description from ECM Export is Null";
                                    //strErrorCol = strErrorCol + "Ecm description to Title issue: Description Null for record no " + h.DocSetID.ToString() + Environment.NewLine;
                                }
                                else
                                {
                                    if (ecm.ECMDescription.Length > 253)
                                    {
                                        fTitle = ecm.ECMDescription.Substring(0, 253);
                                        //strErrorCol = strErrorCol + "Ecm description to Title issue: Description greater then 254 ch, description truncated for record no " + h.DocSetID.ToString() + Environment.NewLine;
                                    }
                                    else if (ecm.ECMDescription.Length < 1)
                                    {
                                        fTitle = "Record description from ECM Export is Empty";
                                        //strErrorCol = strErrorCol + "Ecm description to Title issue: Description blank for record no " + h.DocSetID.ToString() + Environment.NewLine;
                                    }
                                    else
                                    {
                                        fTitle = ecm.ECMDescription;
                                    }
                                }

                                Record rec = null;
                                if (reccheck == null)
                                {
                                    rec = new Record(rt);
                                    rec.Container = reccont;
                                }
                                else
                                {
                                    rec = reccheck;
                                }



                                rec.Title = fTitle;
                                rec.DateCreated = (TrimDateTime)ecm.ECMSTDDocumentDate;
                                rec.SetNotes("Migrated from EMS T1", NotesUpdateType.AppendOnly);
                                rec.LongNumber = ecm.DocSetID;
                                //
                                rec.SetFieldValue(fd_ECMApplicationRAM_Application_Number, UserFieldValue.FromDotNetObject(ecm.ECMApplicationRAM_Application_Number));
                                rec.SetFieldValue(fd_ECMBarCodeBarCodeString, UserFieldValue.FromDotNetObject(ecm.ECMBarCodeBarCodeString));
                                rec.SetFieldValue(fd_ECMBCSRetention_Period, UserFieldValue.FromDotNetObject(ecm.ECMBCSRetention_Period));
                                rec.SetFieldValue(fd_ECMCaseCaseDescription, UserFieldValue.FromDotNetObject(ecm.ECMCaseCaseDescription));
                                rec.SetFieldValue(fd_ECMCaseCaseName, UserFieldValue.FromDotNetObject(ecm.ECMCaseCaseName));
                                rec.SetFieldValue(fd_ECMCaseCaseNumber, UserFieldValue.FromDotNetObject(ecm.ECMCaseCaseNumber));
                                rec.SetFieldValue(fd_ECMCaseCaseOfficer, UserFieldValue.FromDotNetObject(ecm.ECMCaseCaseOfficer));
                                rec.SetFieldValue(fd_ECMCaseCaseStatus, UserFieldValue.FromDotNetObject(ecm.ECMCaseCaseStatus));
                                rec.SetFieldValue(fd_ECMCaseCaseType, UserFieldValue.FromDotNetObject(ecm.ECMCaseCaseType));
                                rec.SetFieldValue(fd_ECMCaseEndDate, UserFieldValue.FromDotNetObject(ecm.ECMCaseEndDate));
                                rec.SetFieldValue(fd_ECMCaseStartDate, UserFieldValue.FromDotNetObject(ecm.ECMCaseStartDate));
                                //rec.SetFieldValue(fd_ECMCorrespondentDescription, UserFieldValue.FromDotNetObject(ecm.ECMCorrespondentDescription));
                                //rec.SetFieldValue(fd_ECMECMDriveID, UserFieldValue.FromDotNetObject(ecm.ECMECMDriveID));
                                rec.SetFieldValue(fd_ECMEmployeeCommencement_Date, UserFieldValue.FromDotNetObject(ecm.ECMEmployeeCommencement_Date));
                                rec.SetFieldValue(fd_ECMEmployeeEmployee_Given_Name, UserFieldValue.FromDotNetObject(ecm.ECMEmployeeEmployee_Given_Name));
                                rec.SetFieldValue(fd_ECMEmployeeEmployee_Surname, UserFieldValue.FromDotNetObject(ecm.ECMEmployeeEmployee_Surname));
                                rec.SetFieldValue(fd_ECMEmployeeEmployeeCode, UserFieldValue.FromDotNetObject(ecm.ECMEmployeeEmployeeCode));
                                rec.SetFieldValue(fd_ECmEmployeeHR_Case_Description, UserFieldValue.FromDotNetObject(ecm.ECmEmployeeHR_Case_Description));
                                rec.SetFieldValue(fd_ECMEmployeePreferred_Name, UserFieldValue.FromDotNetObject(ecm.ECMEmployeePreferred_Name));
                                rec.SetFieldValue(fd_ECMEmployeeStatus, UserFieldValue.FromDotNetObject(ecm.ECMEmployeeStatus));
                                rec.SetFieldValue(fd_ECMEmployeeTermination_Date, UserFieldValue.FromDotNetObject(ecm.ECMEmployeeTermination_Date));
                                rec.SetFieldValue(fd_ECMFileName, UserFieldValue.FromDotNetObject(ecm.ECMFileName));
                                rec.SetFieldValue(fd_ECMFileNetFoldername, UserFieldValue.FromDotNetObject(ecm.ECMFileNetFoldername));
                                rec.SetFieldValue(fd_ECMHR_Case_Code, UserFieldValue.FromDotNetObject(ecm.ECMHR_Case_Code));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel1_Name, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel1_Name));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel2_Name, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel2_Name));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel3_Description, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel3_Description));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel3_Name, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel3_Name));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel4_Description, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel4_Description));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel4_Name, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel4_Name));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel5_Description, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel5_Description));
                                rec.SetFieldValue(fd_ECMInfoExpertLevel5_Name, UserFieldValue.FromDotNetObject(ecm.ECMInfoExpertLevel5_Name));
                                rec.SetFieldValue(fd_ECMInfringementInfringementID, UserFieldValue.FromDotNetObject(ecm.ECMInfringementInfringementID));
                                rec.SetFieldValue(fd_ECMNotes, UserFieldValue.FromDotNetObject(ecm.ECMNotes));
                                rec.SetFieldValue(fd_ECMPositionClose_Date, UserFieldValue.FromDotNetObject(ecm.ECMPositionClose_Date));
                                rec.SetFieldValue(fd_ECMPositionOpen_Date, UserFieldValue.FromDotNetObject(ecm.ECMPositionOpen_Date));
                                rec.SetFieldValue(fd_ECMPositionPosition_ID, UserFieldValue.FromDotNetObject(ecm.ECMPositionPosition_ID));
                                rec.SetFieldValue(fd_ECMPositionPosition_Title, UserFieldValue.FromDotNetObject(ecm.ECMPositionPosition_Title));
                                rec.SetFieldValue(fd_ECMPositionVacancy_ID, UserFieldValue.FromDotNetObject(ecm.ECMPositionVacancy_ID));
                                rec.SetFieldValue(fd_ECMPositionVacancy_Status, UserFieldValue.FromDotNetObject(ecm.ECMPositionVacancy_Status));
                                rec.SetFieldValue(fd_ECMProjectAndContractProject_Name, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractProject_Name));
                                rec.SetFieldValue(fd_ECMProjectAndContractContract_or_Activity_Desc, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractContract_or_Activity_Desc));
                                rec.SetFieldValue(fd_ECMProjectAndContractContract_or_Activity_Name, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractContract_or_Activity_Name));
                                rec.SetFieldValue(fd_ECMProjectAndContractContract_or_Activity_No, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractContract_or_Activity_No));
                                rec.SetFieldValue(fd_ECMProjectAndContractEnd_Date, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractEnd_Date));
                                rec.SetFieldValue(fd_ECMProjectAndContractProject_Description, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractProject_Description));
                                rec.SetFieldValue(fd_ECMProjectAndContractProject_End_Date, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractProject_End_Date));
                                rec.SetFieldValue(fd_ECMProjectAndContractProject_Officer, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractProject_Officer));
                                rec.SetFieldValue(fd_ECMProjectAndContractProject_Start_Date, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractProject_Start_Date));
                                rec.SetFieldValue(fd_ECMProjectAndContractProjectNo, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractProjectNo));
                                rec.SetFieldValue(fd_ECMProjectAndContractResponsible_Officer, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractResponsible_Officer));
                                rec.SetFieldValue(fd_ECMProjectAndContractStart_Date, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractStart_Date));
                                rec.SetFieldValue(fd_ECMProjectAndContractStatus, UserFieldValue.FromDotNetObject(ecm.ECMProjectAndContractStatus));
                                rec.SetFieldValue(fd_ECMPropertyHouse_No, UserFieldValue.FromDotNetObject(ecm.ECMPropertyHouse_No));
                                rec.SetFieldValue(fd_ECMPropertyHouse_No_Suffix, UserFieldValue.FromDotNetObject(ecm.ECMPropertyHouse_No_Suffix));
                                //rec.SetFieldValue(fd_ECMPropertyHouse_No_To, UserFieldValue.FromDotNetObject(ecm.ECMPropertyHouse_No_To));
                                rec.SetFieldValue(fd_ECMPropertyHouse_No_To_Suffix, UserFieldValue.FromDotNetObject(ecm.ECMPropertyHouse_No_To_Suffix));
                                rec.SetFieldValue(fd_ECMPropertyLocality_Name, UserFieldValue.FromDotNetObject(ecm.ECMPropertyLocality_Name));
                                rec.SetFieldValue(fd_ECMPropertyPostcode, UserFieldValue.FromDotNetObject(ecm.ECMPropertyPostcode));
                                rec.SetFieldValue(fd_ECMPropertyProperty_Name, UserFieldValue.FromDotNetObject(ecm.ECMPropertyProperty_Name));
                                rec.SetFieldValue(fd_ECMPropertyProperty_No, UserFieldValue.FromDotNetObject(ecm.ECMPropertyProperty_No));
                                rec.SetFieldValue(fd_ECMPropertyStreet_Name, UserFieldValue.FromDotNetObject(ecm.ECMPropertyStreet_Name));
                                rec.SetFieldValue(fd_ECMPropertyUnit_No, UserFieldValue.FromDotNetObject(ecm.ECMPropertyUnit_No));
                                rec.SetFieldValue(fd_ECMPropertyUnit_No_Suffix, UserFieldValue.FromDotNetObject(ecm.ECMPropertyUnit_No_Suffix));
                                //rec.SetFieldValue(fd_ECMreference, UserFieldValue.FromDotNetObject(ecm.ECMreference));
                                rec.SetFieldValue(fd_ECMRelatedDocDocSetID, UserFieldValue.FromDotNetObject(ecm.ECMRelatedDocDocSetID));
                                rec.SetFieldValue(fd_ECMSTDAllRevTitle, UserFieldValue.FromDotNetObject(ecm.ECMSTDAllRevTitle));
                                rec.SetFieldValue(fd_ECMSTDApplicationNo, UserFieldValue.FromDotNetObject(ecm.ECMSTDApplicationNo));
                                rec.SetFieldValue(fd_ECMSTDBusinessCode, UserFieldValue.FromDotNetObject(ecm.ECMSTDBusinessCode));
                                rec.SetFieldValue(fd_ECMSTDClassName, UserFieldValue.FromDotNetObject(ecm.ECMSTDClassName));
                                rec.SetFieldValue(fd_ECMSTDCorrespondant, UserFieldValue.FromDotNetObject(ecm.ECMSTDCorrespondant));
                                rec.SetFieldValue(fd_ECMSTDCustomerRequest, UserFieldValue.FromDotNetObject(ecm.ECMSTDCustomerRequest));
                                rec.SetFieldValue(fd_ECMSTDDateLastAccessed, UserFieldValue.FromDotNetObject(ecm.ECMSTDDateLastAccessed));
                                rec.SetFieldValue(fd_ECMSTDDateReceived, UserFieldValue.FromDotNetObject(ecm.ECMSTDDateReceived));
                                rec.SetFieldValue(fd_ECMSTDDateRegistered, UserFieldValue.FromDotNetObject(ecm.ECMSTDDateRegistered));
                                rec.SetFieldValue(fd_ECMSTDDeclaredDate, UserFieldValue.FromDotNetObject(ecm.ECMSTDDeclaredDate));
                                rec.SetFieldValue(fd_ECMSTDDestructionDue, UserFieldValue.FromDotNetObject(ecm.ECMSTDDestructionDue));
                                rec.SetFieldValue(fd_ECMSTDDocumentDate, UserFieldValue.FromDotNetObject(ecm.ECMSTDDocumentDate));
                                rec.SetFieldValue(fd_ECMSTDDocumentType, UserFieldValue.FromDotNetObject(ecm.ECMSTDDocumentType));
                                rec.SetFieldValue(fd_ECMSTDExternalReference, UserFieldValue.FromDotNetObject(ecm.ECMSTDExternalReference));
                                rec.SetFieldValue(fd_ECMSTDInfringementNo, UserFieldValue.FromDotNetObject(ecm.ECMSTDInfringementNo));
                                rec.SetFieldValue(fd_ECMSTDInternalReference, UserFieldValue.FromDotNetObject(ecm.ECMSTDInternalReference));
                                rec.SetFieldValue(fd_ECMSTDInternalScanRequest, UserFieldValue.FromDotNetObject(ecm.ECMSTDInternalScanRequest));
                                rec.SetFieldValue(fd_ECMSTDJobNo, UserFieldValue.FromDotNetObject(ecm.ECMSTDJobNo));
                                rec.SetFieldValue(fd_ECMSTDOtherReferences, UserFieldValue.FromDotNetObject(ecm.ECMSTDOtherReferences));
                                rec.SetFieldValue(fd_ECMSTDPropertyNo, UserFieldValue.FromDotNetObject(ecm.ECMSTDPropertyNo));
                                rec.SetFieldValue(fd_ECMSTDSummaryText, UserFieldValue.FromDotNetObject(ecm.ECMSTDSummaryText));
                                rec.SetFieldValue(fd_ECMStreetsLocality, UserFieldValue.FromDotNetObject(ecm.ECMStreetsLocality));
                                rec.SetFieldValue(fd_ECMStreetsPostcode, UserFieldValue.FromDotNetObject(ecm.ECMStreetsPostcode));
                                rec.SetFieldValue(fd_ECMStreetsStreet_Name, UserFieldValue.FromDotNetObject(ecm.ECMStreetsStreet_Name));
                                rec.SetFieldValue(fd_ECMUserGroupsDescription, UserFieldValue.FromDotNetObject(ecm.ECMUserGroupsDescription));
                                rec.SetFieldValue(fd_ECMUserGroupsExtNo, UserFieldValue.FromDotNetObject(ecm.ECMUserGroupsExtNo));
                                rec.SetFieldValue(fd_ECMUserGroupsGivenName, UserFieldValue.FromDotNetObject(ecm.ECMUserGroupsGivenName));
                                rec.SetFieldValue(fd_ECMUserGroupsOrgEmail, UserFieldValue.FromDotNetObject(ecm.ECMUserGroupsOrgEmail));
                                rec.SetFieldValue(fd_ECMUserGroupsSurname, UserFieldValue.FromDotNetObject(ecm.ECMUserGroupsSurname));
                                rec.SetFieldValue(fd_ECMVolumeFilename, UserFieldValue.FromDotNetObject(ecm.ECMVolumeFilename));
                                rec.SetFieldValue(fd_ECMVolumeLastModified, UserFieldValue.FromDotNetObject(ecm.ECMVolumeLastModified));
                                rec.SetFieldValue(fd_ECMVolumeMedia, UserFieldValue.FromDotNetObject(ecm.ECMVolumeMedia));
                                rec.SetFieldValue(fd_ECMVolumePrimaryCacheName, UserFieldValue.FromDotNetObject(ecm.ECMVolumePrimaryCacheName));
                                rec.SetFieldValue(fd_ECMVolumeRenditionTypeName, UserFieldValue.FromDotNetObject(ecm.ECMVolumeRenditionTypeName));
                                rec.SetFieldValue(fd_ECMVolumeShareDrive, UserFieldValue.FromDotNetObject(ecm.ECMVolumeShareDrive));
                                rec.SetFieldValue(fd_ECMVolumeSize, UserFieldValue.FromDotNetObject(ecm.ECMVolumeSize));
                                rec.SetFieldValue(fd_ECMVolumeStatus, UserFieldValue.FromDotNetObject(ecm.ECMVolumeStatus));
                                rec.SetFieldValue(fd_ECMVolumeStorageLocation, UserFieldValue.FromDotNetObject(ecm.ECMVolumeStorageLocation));
                                //rec.SetFieldValue(fd_ECMVolumeUpdatable, UserFieldValue.FromDotNetObject(ecm.ECMVolumeUpdatable));
                                try
                                {
                                    InputDocument doc = new InputDocument();
                                    doc.SetAsFile(strStorageLoc);
                                    rec.SetDocument(doc, false, false, "Imported from ECM");
                                    rec.Save();
                                    deleteEcmSource(strStorageLoc);
                                    Console.WriteLine("New record created: " + rec.Number);
                                    UpdateMasterinRM(rec);
                                }
                                catch (Exception exp)
                                {
                                    Console.WriteLine("Save new record error: " + exp.Message.ToString());
                                    UpdateMasterNoRm(ecm.DocSetID, 0);
                                    Loggitt("Record number " + ecm.DocSetID + "  could not be added - record save: " + exp.Message.ToString() + Environment.NewLine + iCount.ToString() + " " + strTotal + "" + result.TotalMinutes.ToString() + " " + (Convert.ToInt32(strOffset) + iCount).ToString(), "ECM Migrate issue");

                                }
                            }
                            else
                            {
                                Console.WriteLine("Cant find ECM container in Eddie: " + ecm.ECMBCSFunction_Name + " - " + ecm.ECMBCSActivity_Name.Replace("-", "~"));
                            }

                        }
                        else
                        {
                            Console.WriteLine("ECM electronic document cant be found: " + strStorageLoc);
                            UpdateMasterNoRm(ecm.DocSetID, 1);
                            // UpdateMasterinRMnoDoc()
                        }

                        //}
                        //else
                        //{
                        //    Console.WriteLine("ECM record already exists: " + ecm.DocSetID);
                        //    UpdateMasterinRM(reccheck);
                        //}

                    }
                    catch (Exception exp)
                    {
                        Console.WriteLine("Outer look error: " + exp.Message.ToString());
                        Loggitt("Record number " + ecm.DocSetID + "  could not be added -  Outer Error: " + exp.Message.ToString() + Environment.NewLine + iCount.ToString() + " " + strTotal + "" + result.TotalMinutes.ToString() + " " + (Convert.ToInt32(strOffset) + iCount).ToString(), "ECM Migrate issue");
                        //Console.ReadLine();
                    }
                }
            }
            //});

        }
        private static void deleteEcmSource(string strStorageLoc)
        {
            try
            {
                if (File.Exists(strStorageLoc))
                {
                    File.Delete(strStorageLoc);
                    Console.WriteLine("ECM File deleted ok");
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error deleting ECm file: " + strStorageLoc);
            }
        }
        private static string FindMappedStorage(string eCMVolumeStorageLocation)
        {
            //string dfd = @"\\scas57\ECMIncache18\";
            string skjfdh = null;
            if (eCMVolumeStorageLocation.Length == 21)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 11);
            }
            else if (eCMVolumeStorageLocation.Length == 22)
            {
                skjfdh = eCMVolumeStorageLocation.Substring(9, 12);
            }
            //var skjfdh = eCMVolumeStorageLocation.Substring(9, 11);
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
        private static void UpdateMasterinRM(Record r)
        {
            try
            {
                if (bDetailMsg)
                {
                    Console.WriteLine("Starting UpdateMasterinRM");
                    Console.ReadLine();
                }
                long ruri = r.Uri;
                int rRev = r.RevisionNumber;
                //"Select DocumentSetID, Origbatch, ClassName, maxversion from ECM_MasterCheck"
                string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, RMUri=@rmuri, RmRevisions=@ver, Issue=@issue WHERE DocumentSetID=@docsetid";
                if (bDetailMsg)
                {
                    Console.WriteLine("UpdateMasterinRM query: " + SQL);
                    Console.ReadLine();
                }
                SqlConnection con = new SqlConnection(strCon);
                SqlCommand cmd = new SqlCommand(SQL, con);
                cmd.Parameters.AddWithValue("@docsetid", r.Number);
                cmd.Parameters.AddWithValue("@now", DateTime.Now);
                cmd.Parameters.AddWithValue("@rmuri", ruri);
                cmd.Parameters.AddWithValue("@ver", rRev);
                cmd.Parameters.AddWithValue("@issue", 0);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                if (bDetailMsg)
                {
                    Console.WriteLine("Query executed ok:");
                    Console.ReadLine();
                }
            }
            catch (Exception exp)
            {
                if (bDetailMsg)
                {
                    Console.WriteLine("Query execution failed: Error " + exp.Message.ToString());
                    Console.ReadLine();
                }
            }

        }
        private static void UpdateMasterNoRm(string recnum, int prob)
        {
            try
            {
                if (bDetailMsg)
                {
                    Console.WriteLine("Starting UpdateMasterinRM");
                    Console.ReadLine();
                }
                string SQL = "UPDATE ECM_MasterCheck SET Checked=@now, Issue=@issue WHERE DocumentSetID=@docsetid";
                if (bDetailMsg)
                {
                    Console.WriteLine("UpdateMasterinRM query: " + SQL);
                    Console.ReadLine();
                }
                SqlConnection con = new SqlConnection(strCon);
                SqlCommand cmd = new SqlCommand(SQL, con);
                cmd.Parameters.AddWithValue("@docsetid", recnum);
                cmd.Parameters.AddWithValue("@now", DateTime.Now);
                cmd.Parameters.AddWithValue("@issue", prob);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                if (bDetailMsg)
                {
                    Console.WriteLine("Query executed ok:");
                    Console.ReadLine();
                }
            }
            catch (Exception exp)
            {
                if (bDetailMsg)
                {
                    Console.WriteLine("Query execution failed: Error " + exp.Message.ToString());
                    Console.ReadLine();
                }
            }
        }
        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\ECM\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\ECM\");
            }
            FileStream fs = new FileStream(@"C:\Temp\ECM\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\ECM\log_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
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
