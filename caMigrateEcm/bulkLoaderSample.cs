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
    {   class ECMMigration
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
        class ECMMigration1
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
                    m_origin.DefaultRecordType = db.FindTrimObjectByUri(BaseObjectTypes.RecordType, 16) as RecordType;
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
                PropertyOrFieldValue[] recordFields = new PropertyOrFieldValue[111];
                recordFields[0] = new PropertyOrFieldValue(PropertyIds.RecordTitle);
                recordFields[1] = new PropertyOrFieldValue(PropertyIds.RecordDateCreated);
                recordFields[2] = new PropertyOrFieldValue(PropertyIds.RecordNotes);
                recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordContainer);
                recordFields[4] = new PropertyOrFieldValue(PropertyIds.RecordNumber);
                recordFields[110] = new PropertyOrFieldValue(PropertyIds.RecordBarcode);

                //
                FieldDefinition fd_ECMApplicationRAM_Application_Number = new FieldDefinition(db, "Application Number List");
                recordFields[5] = new PropertyOrFieldValue(fd_ECMApplicationRAM_Application_Number);

                FieldDefinition fd_ECMBarCodeBarCodeString = new FieldDefinition(db, "ECM BarCode:BarCodeString");
                recordFields[6] = new PropertyOrFieldValue(fd_ECMBarCodeBarCodeString);

                FieldDefinition fd_ECMBCSRetention_Period = new FieldDefinition(db, "ECM BCS:Retention_Period");
                recordFields[7] = new PropertyOrFieldValue(fd_ECMBCSRetention_Period);

                FieldDefinition fd_ECMCaseCaseDescription = new FieldDefinition(db, "ECM Case:CaseDescription");
                recordFields[8] = new PropertyOrFieldValue(fd_ECMCaseCaseDescription);

                FieldDefinition fd_ECMCaseCaseName = new FieldDefinition(db, "ECM Case:CaseName");
                recordFields[9] = new PropertyOrFieldValue(fd_ECMCaseCaseName);

                FieldDefinition fd_ECMCaseCaseNumber = new FieldDefinition(db, "ECM Case:CaseNumber");
                recordFields[10] = new PropertyOrFieldValue(fd_ECMCaseCaseNumber);

                FieldDefinition fd_ECMCaseCaseOfficer = new FieldDefinition(db, "ECM Case:CaseOfficer");
                recordFields[11] = new PropertyOrFieldValue(fd_ECMCaseCaseOfficer);

                FieldDefinition fd_ECMCaseCaseStatus = new FieldDefinition(db, "ECM Case:CaseStatus");
                recordFields[12] = new PropertyOrFieldValue(fd_ECMCaseCaseStatus);

                FieldDefinition fd_ECMCaseCaseType = new FieldDefinition(db, "ECM Case:CaseType");
                recordFields[13] = new PropertyOrFieldValue(fd_ECMCaseCaseType);

                FieldDefinition fd_ECMCaseEndDate = new FieldDefinition(db, "ECM Case:EndDate");
                recordFields[14] = new PropertyOrFieldValue(fd_ECMCaseEndDate);

                FieldDefinition fd_ECMCaseStartDate = new FieldDefinition(db, "ECM Case:StartDate");
                recordFields[15] = new PropertyOrFieldValue(fd_ECMCaseStartDate);

                FieldDefinition fd_ECMCorrespondentDescription = new FieldDefinition(db, "ECM Correspondent:Description");
                recordFields[16] = new PropertyOrFieldValue(fd_ECMCorrespondentDescription);

                FieldDefinition fd_ECMECMDriveID = new FieldDefinition(db, "ECM ECMDriveID");
                recordFields[17] = new PropertyOrFieldValue(fd_ECMECMDriveID);

                FieldDefinition fd_ECMEmployeeCommencement_Date = new FieldDefinition(db, "ECM Employee:Commencement_Date");
                recordFields[18] = new PropertyOrFieldValue(fd_ECMEmployeeCommencement_Date);

                FieldDefinition fd_ECMEmployeeEmployee_Given_Name = new FieldDefinition(db, "ECM Employee:Employee_Given_Name");
                recordFields[19] = new PropertyOrFieldValue(fd_ECMEmployeeEmployee_Given_Name);

                FieldDefinition fd_ECMEmployeeEmployee_Surname = new FieldDefinition(db, "ECM Employee:Employee_Surname");
                recordFields[20] = new PropertyOrFieldValue(fd_ECMEmployeeEmployee_Surname);

                FieldDefinition fd_ECMEmployeeEmployeeCode = new FieldDefinition(db, "ECM Employee:EmployeeCode");
                recordFields[21] = new PropertyOrFieldValue(fd_ECMEmployeeEmployeeCode);

                FieldDefinition fd_ECmEmployeeHR_Case_Description = new FieldDefinition(db, "ECm Employee:HR_Case_Description");
                recordFields[22] = new PropertyOrFieldValue(fd_ECmEmployeeHR_Case_Description);

                FieldDefinition fd_ECMEmployeePreferred_Name = new FieldDefinition(db, "ECM Employee:Preferred_Name");
                recordFields[23] = new PropertyOrFieldValue(fd_ECMEmployeePreferred_Name);

                FieldDefinition fd_ECMEmployeeStatus = new FieldDefinition(db, "ECM Employee:Status");
                recordFields[24] = new PropertyOrFieldValue(fd_ECMEmployeeStatus);

                FieldDefinition fd_ECMEmployeeTermination_Date = new FieldDefinition(db, "ECM Employee:Termination_Date");
                recordFields[25] = new PropertyOrFieldValue(fd_ECMEmployeeTermination_Date);

                FieldDefinition fd_ECMFileName = new FieldDefinition(db, "ECM FileName");
                recordFields[26] = new PropertyOrFieldValue(fd_ECMFileName);

                FieldDefinition fd_ECMFileNetFoldername = new FieldDefinition(db, "ECM FileNet:Foldername");
                recordFields[27] = new PropertyOrFieldValue(fd_ECMFileNetFoldername);

                FieldDefinition fd_ECMHR_Case_Code = new FieldDefinition(db, "ECM HR_Case_Code");
                recordFields[28] = new PropertyOrFieldValue(fd_ECMHR_Case_Code);

                FieldDefinition fd_ECMInfoExpertLevel1_Name = new FieldDefinition(db, "ECM InfoExpert:Level1_Name");
                recordFields[29] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel1_Name);

                FieldDefinition fd_ECMInfoExpertLevel2_Name = new FieldDefinition(db, "ECM InfoExpert:Level2_Name");
                recordFields[30] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel2_Name);

                FieldDefinition fd_ECMInfoExpertLevel3_Description = new FieldDefinition(db, "ECM InfoExpert:Level3_Description");
                recordFields[31] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel3_Description);

                FieldDefinition fd_ECMInfoExpertLevel3_Name = new FieldDefinition(db, "ECM InfoExpert:Level3_Name");
                recordFields[32] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel3_Name);

                FieldDefinition fd_ECMInfoExpertLevel4_Description = new FieldDefinition(db, "ECM InfoExpert:Level4_Description");
                recordFields[33] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel4_Description);

                FieldDefinition fd_ECMInfoExpertLevel4_Name = new FieldDefinition(db, "ECM InfoExpert:Level4_Name");
                recordFields[34] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel4_Name);

                FieldDefinition fd_ECMInfoExpertLevel5_Description = new FieldDefinition(db, "ECM InfoExpert:Level5_Description");
                recordFields[35] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel5_Description);

                FieldDefinition fd_ECMInfoExpertLevel5_Name = new FieldDefinition(db, "ECM InfoExpert:Level5_Name");
                recordFields[36] = new PropertyOrFieldValue(fd_ECMInfoExpertLevel5_Name);

                FieldDefinition fd_ECMInfringementInfringementID = new FieldDefinition(db, "ECM Infringement:InfringementID");
                recordFields[37] = new PropertyOrFieldValue(fd_ECMInfringementInfringementID);

                FieldDefinition fd_ECMNotes = new FieldDefinition(db, "ECM Notes");
                recordFields[38] = new PropertyOrFieldValue(fd_ECMNotes);

                FieldDefinition fd_ECMPositionClose_Date = new FieldDefinition(db, "ECM Position:Close_Date");
                recordFields[39] = new PropertyOrFieldValue(fd_ECMPositionClose_Date);

                FieldDefinition fd_ECMPositionOpen_Date = new FieldDefinition(db, "ECM Position:Open_Date");
                recordFields[40] = new PropertyOrFieldValue(fd_ECMPositionOpen_Date);

                FieldDefinition fd_ECMPositionPosition_ID = new FieldDefinition(db, "ECM Position:Position_ID");
                recordFields[41] = new PropertyOrFieldValue(fd_ECMPositionPosition_ID);

                FieldDefinition fd_ECMPositionPosition_Title = new FieldDefinition(db, "ECM Position:Position_Title");
                recordFields[42] = new PropertyOrFieldValue(fd_ECMPositionPosition_Title);

                FieldDefinition fd_ECMPositionVacancy_ID = new FieldDefinition(db, "ECM Position:Vacancy_ID");
                recordFields[43] = new PropertyOrFieldValue(fd_ECMPositionVacancy_ID);

                FieldDefinition fd_ECMPositionVacancy_Status = new FieldDefinition(db, "ECM Position:Vacancy_Status");
                recordFields[44] = new PropertyOrFieldValue(fd_ECMPositionVacancy_Status);

                FieldDefinition fd_ECMProjectAndContractProject_Name = new FieldDefinition(db, "ECM ProjectAndContract:[Project_Name");
                recordFields[45] = new PropertyOrFieldValue(fd_ECMProjectAndContractProject_Name);

                FieldDefinition fd_ECMProjectAndContractContract_or_Activity_Desc = new FieldDefinition(db, "ECM ProjectAndContract:Contract_or_Activity_Desc");
                recordFields[46] = new PropertyOrFieldValue(fd_ECMProjectAndContractContract_or_Activity_Desc);

                FieldDefinition fd_ECMProjectAndContractContract_or_Activity_Name = new FieldDefinition(db, "ECM ProjectAndContract:Contract_or_Activity_Name");
                recordFields[47] = new PropertyOrFieldValue(fd_ECMProjectAndContractContract_or_Activity_Name);

                FieldDefinition fd_ECMProjectAndContractContract_or_Activity_No = new FieldDefinition(db, "ECM ProjectAndContract:Contract_or_Activity_No");
                recordFields[48] = new PropertyOrFieldValue(fd_ECMProjectAndContractContract_or_Activity_No);

                FieldDefinition fd_ECMProjectAndContractEnd_Date = new FieldDefinition(db, "ECM ProjectAndContract:End_Date");
                recordFields[49] = new PropertyOrFieldValue(fd_ECMProjectAndContractEnd_Date);

                FieldDefinition fd_ECMProjectAndContractProject_Description = new FieldDefinition(db, "ECM ProjectAndContract:Project_Description");
                recordFields[50] = new PropertyOrFieldValue(fd_ECMProjectAndContractProject_Description);

                FieldDefinition fd_ECMProjectAndContractProject_End_Date = new FieldDefinition(db, "ECM ProjectAndContract:Project_End_Date");
                recordFields[51] = new PropertyOrFieldValue(fd_ECMProjectAndContractProject_End_Date);

                FieldDefinition fd_ECMProjectAndContractProject_Officer = new FieldDefinition(db, "ECM ProjectAndContract:Project_Officer");
                recordFields[52] = new PropertyOrFieldValue(fd_ECMProjectAndContractProject_Officer);

                FieldDefinition fd_ECMProjectAndContractProject_Start_Date = new FieldDefinition(db, "ECM ProjectAndContract:Project_Start_Date");
                recordFields[53] = new PropertyOrFieldValue(fd_ECMProjectAndContractProject_Start_Date);

                FieldDefinition fd_ECMProjectAndContractProjectNo = new FieldDefinition(db, "ECM ProjectAndContract:ProjectNo");
                recordFields[54] = new PropertyOrFieldValue(fd_ECMProjectAndContractProjectNo);

                FieldDefinition fd_ECMProjectAndContractResponsible_Officer = new FieldDefinition(db, "ECM ProjectAndContract:Responsible_Officer");
                recordFields[55] = new PropertyOrFieldValue(fd_ECMProjectAndContractResponsible_Officer);

                FieldDefinition fd_ECMProjectAndContractStart_Date = new FieldDefinition(db, "ECM ProjectAndContract:Start_Date");
                recordFields[56] = new PropertyOrFieldValue(fd_ECMProjectAndContractStart_Date);

                FieldDefinition fd_ECMProjectAndContractStatus = new FieldDefinition(db, "ECM ProjectAndContract:Status");
                recordFields[57] = new PropertyOrFieldValue(fd_ECMProjectAndContractStatus);

                FieldDefinition fd_ECMPropertyHouse_No = new FieldDefinition(db, "ECM Property:House_No");
                recordFields[58] = new PropertyOrFieldValue(fd_ECMPropertyHouse_No);

                FieldDefinition fd_ECMPropertyHouse_No_Suffix = new FieldDefinition(db, "ECM Property:House_No_Suffix");
                recordFields[59] = new PropertyOrFieldValue(fd_ECMPropertyHouse_No_Suffix);

                FieldDefinition fd_ECMPropertyHouse_No_To = new FieldDefinition(db, "ECM Property:House_No_To");
                recordFields[60] = new PropertyOrFieldValue(fd_ECMPropertyHouse_No_To);

                FieldDefinition fd_ECMPropertyHouse_No_To_Suffix = new FieldDefinition(db, "ECM Property:House_No_To_Suffix");
                recordFields[61] = new PropertyOrFieldValue(fd_ECMPropertyHouse_No_To_Suffix);

                FieldDefinition fd_ECMPropertyLocality_Name = new FieldDefinition(db, "ECM Property:Locality_Name");
                recordFields[62] = new PropertyOrFieldValue(fd_ECMPropertyLocality_Name);

                FieldDefinition fd_ECMPropertyPostcode = new FieldDefinition(db, "ECM Property:Postcode");
                recordFields[63] = new PropertyOrFieldValue(fd_ECMPropertyPostcode);

                FieldDefinition fd_ECMPropertyProperty_Name = new FieldDefinition(db, "ECM Property:Property_Name");
                recordFields[64] = new PropertyOrFieldValue(fd_ECMPropertyProperty_Name);

                FieldDefinition fd_ECMPropertyProperty_No = new FieldDefinition(db, "Property Number List");
                recordFields[65] = new PropertyOrFieldValue(fd_ECMPropertyProperty_No);

                FieldDefinition fd_ECMPropertyStreet_Name = new FieldDefinition(db, "ECM Property:Street_Name");
                recordFields[66] = new PropertyOrFieldValue(fd_ECMPropertyStreet_Name);

                FieldDefinition fd_ECMPropertyUnit_No = new FieldDefinition(db, "ECM Property:Unit_No");
                recordFields[67] = new PropertyOrFieldValue(fd_ECMPropertyUnit_No);

                FieldDefinition fd_ECMPropertyUnit_No_Suffix = new FieldDefinition(db, "ECM Property:Unit_No_Suffix");
                recordFields[68] = new PropertyOrFieldValue(fd_ECMPropertyUnit_No_Suffix);

                FieldDefinition fd_ECMreference = new FieldDefinition(db, "ECM reference");
                recordFields[69] = new PropertyOrFieldValue(fd_ECMreference);

                FieldDefinition fd_ECMRelatedDocDocSetID = new FieldDefinition(db, "ECM RelatedDoc:DocSetID");
                recordFields[70] = new PropertyOrFieldValue(fd_ECMRelatedDocDocSetID);

                FieldDefinition fd_ECMSTDAllRevTitle = new FieldDefinition(db, "ECM STD:AllRevTitle");
                recordFields[71] = new PropertyOrFieldValue(fd_ECMSTDAllRevTitle);

                FieldDefinition fd_ECMSTDApplicationNo = new FieldDefinition(db, "ECM STD:ApplicationNo");
                recordFields[72] = new PropertyOrFieldValue(fd_ECMSTDApplicationNo);

                FieldDefinition fd_ECMSTDBusinessCode = new FieldDefinition(db, "ECM STD:BusinessCode");
                recordFields[73] = new PropertyOrFieldValue(fd_ECMSTDBusinessCode);

                FieldDefinition fd_ECMSTDClassName = new FieldDefinition(db, "ECM STD:ClassName");
                recordFields[74] = new PropertyOrFieldValue(fd_ECMSTDClassName);

                FieldDefinition fd_ECMSTDCorrespondant = new FieldDefinition(db, "ECM STD:Correspondant");
                recordFields[75] = new PropertyOrFieldValue(fd_ECMSTDCorrespondant);

                FieldDefinition fd_ECMSTDCustomerRequest = new FieldDefinition(db, "ECM STD:CustomerRequest");
                recordFields[76] = new PropertyOrFieldValue(fd_ECMSTDCustomerRequest);

                FieldDefinition fd_ECMSTDDateLastAccessed = new FieldDefinition(db, "ECM STD:DateLastAccessed");
                recordFields[77] = new PropertyOrFieldValue(fd_ECMSTDDateLastAccessed);

                FieldDefinition fd_ECMSTDDateReceived = new FieldDefinition(db, "ECM STD:DateReceived");
                recordFields[78] = new PropertyOrFieldValue(fd_ECMSTDDateReceived);

                FieldDefinition fd_ECMSTDDateRegistered = new FieldDefinition(db, "ECM STD:DateRegistered");
                recordFields[79] = new PropertyOrFieldValue(fd_ECMSTDDateRegistered);

                FieldDefinition fd_ECMSTDDeclaredDate = new FieldDefinition(db, "ECM STD:DeclaredDate");
                recordFields[80] = new PropertyOrFieldValue(fd_ECMSTDDeclaredDate);

                FieldDefinition fd_ECMSTDDestructionDue = new FieldDefinition(db, "ECM STD:DestructionDue");
                recordFields[81] = new PropertyOrFieldValue(fd_ECMSTDDestructionDue);

                FieldDefinition fd_ECMSTDDocumentDate = new FieldDefinition(db, "ECM STD:DocumentDate");
                recordFields[82] = new PropertyOrFieldValue(fd_ECMSTDDocumentDate);

                FieldDefinition fd_ECMSTDDocumentType = new FieldDefinition(db, "ECM STD:DocumentType");
                recordFields[83] = new PropertyOrFieldValue(fd_ECMSTDDocumentType);

                FieldDefinition fd_ECMSTDExternalReference = new FieldDefinition(db, "ECM STD:ExternalReference");
                recordFields[84] = new PropertyOrFieldValue(fd_ECMSTDExternalReference);

                FieldDefinition fd_ECMSTDInfringementNo = new FieldDefinition(db, "Infringement Number");
                recordFields[85] = new PropertyOrFieldValue(fd_ECMSTDInfringementNo);

                FieldDefinition fd_ECMSTDInternalReference = new FieldDefinition(db, "ECM STD:InternalReference");
                recordFields[86] = new PropertyOrFieldValue(fd_ECMSTDInternalReference);

                FieldDefinition fd_ECMSTDInternalScanRequest = new FieldDefinition(db, "ECM STD:InternalScanRequest");
                recordFields[87] = new PropertyOrFieldValue(fd_ECMSTDInternalScanRequest);

                FieldDefinition fd_ECMSTDJobNo = new FieldDefinition(db, "ECM STD:JobNo");
                recordFields[88] = new PropertyOrFieldValue(fd_ECMSTDJobNo);

                FieldDefinition fd_ECMSTDOtherReferences = new FieldDefinition(db, "ECM STD:OtherReferences");
                recordFields[89] = new PropertyOrFieldValue(fd_ECMSTDOtherReferences);

                FieldDefinition fd_ECMSTDPropertyNo = new FieldDefinition(db, "ECM STD:PropertyNo");
                recordFields[90] = new PropertyOrFieldValue(fd_ECMSTDPropertyNo);

                FieldDefinition fd_ECMSTDSummaryText = new FieldDefinition(db, "ECM STD:SummaryText");
                recordFields[91] = new PropertyOrFieldValue(fd_ECMSTDSummaryText);

                FieldDefinition fd_ECMStreetsLocality = new FieldDefinition(db, "ECM Streets:Locality");
                recordFields[92] = new PropertyOrFieldValue(fd_ECMStreetsLocality);

                FieldDefinition fd_ECMStreetsPostcode = new FieldDefinition(db, "ECM Streets:Postcode");
                recordFields[93] = new PropertyOrFieldValue(fd_ECMStreetsPostcode);

                FieldDefinition fd_ECMStreetsStreet_Name = new FieldDefinition(db, "ECM Streets:Street_Name");
                recordFields[94] = new PropertyOrFieldValue(fd_ECMStreetsStreet_Name);

                FieldDefinition fd_ECMUserGroupsDescription = new FieldDefinition(db, "ECM UserGroups:Description");
                recordFields[95] = new PropertyOrFieldValue(fd_ECMUserGroupsDescription);

                FieldDefinition fd_ECMUserGroupsExtNo = new FieldDefinition(db, "ECM UserGroups:ExtNo");
                recordFields[96] = new PropertyOrFieldValue(fd_ECMUserGroupsExtNo);

                FieldDefinition fd_ECMUserGroupsGivenName = new FieldDefinition(db, "ECM UserGroups:GivenName");
                recordFields[97] = new PropertyOrFieldValue(fd_ECMUserGroupsGivenName);

                FieldDefinition fd_ECMUserGroupsOrgEmail = new FieldDefinition(db, "ECM UserGroups:OrgEmail");
                recordFields[98] = new PropertyOrFieldValue(fd_ECMUserGroupsOrgEmail);

                FieldDefinition fd_ECMUserGroupsSurname = new FieldDefinition(db, "ECM UserGroups:Surname");
                recordFields[99] = new PropertyOrFieldValue(fd_ECMUserGroupsSurname);

                FieldDefinition fd_ECMVolumeFilename = new FieldDefinition(db, "ECM Volume:Filename");
                recordFields[100] = new PropertyOrFieldValue(fd_ECMVolumeFilename);

                FieldDefinition fd_ECMVolumeLastModified = new FieldDefinition(db, "ECM Volume:LastModified");
                recordFields[101] = new PropertyOrFieldValue(fd_ECMVolumeLastModified);

                FieldDefinition fd_ECMVolumeMedia = new FieldDefinition(db, "ECM Volume:Media");
                recordFields[102] = new PropertyOrFieldValue(fd_ECMVolumeMedia);

                FieldDefinition fd_ECMVolumePrimaryCacheName = new FieldDefinition(db, "ECM Volume:PrimaryCacheName");
                recordFields[103] = new PropertyOrFieldValue(fd_ECMVolumePrimaryCacheName);

                FieldDefinition fd_ECMVolumeRenditionTypeName = new FieldDefinition(db, "ECM Volume:RenditionTypeName");
                recordFields[104] = new PropertyOrFieldValue(fd_ECMVolumeRenditionTypeName);

                FieldDefinition fd_ECMVolumeShareDrive = new FieldDefinition(db, "ECM Volume:ShareDrive");
                recordFields[105] = new PropertyOrFieldValue(fd_ECMVolumeShareDrive);

                FieldDefinition fd_ECMVolumeSize = new FieldDefinition(db, "ECM Volume:Size");
                recordFields[106] = new PropertyOrFieldValue(fd_ECMVolumeSize);

                FieldDefinition fd_ECMVolumeStatus = new FieldDefinition(db, "ECM Volume:Status");
                recordFields[107] = new PropertyOrFieldValue(fd_ECMVolumeStatus);

                FieldDefinition fd_ECMVolumeStorageLocation = new FieldDefinition(db, "ECM Volume:StorageLocation");
                recordFields[108] = new PropertyOrFieldValue(fd_ECMVolumeStorageLocation);

                FieldDefinition fd_ECMVolumeUpdatable = new FieldDefinition(db, "ECM Volume:Updatable");
                recordFields[109] = new PropertyOrFieldValue(fd_ECMVolumeUpdatable);
                //
                //recordFields[110] = new PropertyOrFieldValue(PropertyIds.RecordSecurityLevel);
                //SecurityLevel sl = new SecurityLevel(db, 2);
                //recordFields[110].SetValue(sl.Uri);

                
                //recordFields[111] = new PropertyOrFieldValue(PropertyIds.RecordDateFinalized);

                




                //string connectionString = "Data Source=yfdev.cloudapp.net;Initial Catalog=SCC_ECM;Persist Security Info=True;User ID=EPL;Password=Password1!";
                //string connectionString = "Data Source=MSI-GS60;Initial Catalog=ECMNew;Integrated Security=True";
                //Console.ReadLine();
                //
                Console.WriteLine("Enter Connection string");

                SqlConnection con = new SqlConnection(Console.ReadLine());
                //{
                con.Open();
                Console.WriteLine("Enter SQL string");
                List<ECMMigration> lstecm = new List<ECMMigration>();
                using (SqlCommand command = new SqlCommand(Console.ReadLine(), con))
                {
                    DataSet dataSet = new DataSet();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = command;
                    da.Fill(dt);                  
                    SqlDataReader reader = command.ExecuteReader();
                    //
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

                    con.Close();
                    con.Dispose();
                }

                    string strIndexLoc = null;
                    //Console.WriteLine("Enter the network location of the index");
                    //string strInloc = Console.ReadLine();
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
                    //if (!Directory.Exists(strInloc))
                    //{
                    //    strInloc = null;
                    //}

                    string strTotal = lstecm.Count().ToString();
                    int intCount = 0;

                //int intMaxVersion = lstecm.Max(x => x.ECMSTDVersion);
                //for (int i = 1; i < intMaxVersion; i++)
                //{

                //}
                foreach (ECMMigration h in lstecm.Where(x => x.ECMSTDVersion == 1))
                {
                    intCount = intCount + 1;
                    try
                    {
                        string strStorageLoc = FindMappedStorage(h.ECMVolumeStorageLocation) + "\\" + h.ECMVolumeFilename;
                        //long chkExistingrec = m_loader.FindRecord(h.DocSetID.ToString());
                        if (m_loader.FindRecord(h.DocSetID.ToString()) == 0)
                        {
                            if (File.Exists(strStorageLoc))
                            {
                                Record reccont = GetFolderUri(h.ECMBCSFunction_Name + " - " + h.ECMBCSActivity_Name.Replace("-", "~"), h.ECMBCSSubject_Name, db);
                                if (reccont != null)
                                {
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
                                    recordFields[1].SetValue((TrimDateTime)h.ECMSTDDocumentDate);
                                    //recordFields[1].SetValue((TrimDateTime)DateTime.Now);
                                    recordFields[2].SetValue("Migrated from EMS T1 using " + m_origin.Name);
                                    recordFields[3].SetValue(reccont.Uri); //Container
                                    recordFields[4].SetValue(h.DocSetID); //record Number
                                    try
                                    {
                                        recordFields[110].SetValue(h.ECMBarCodeBarCodeString);
                                    }
                                    catch (Exception exp)
                                    {
                                        strErrorCol = strErrorCol + "ECMBarCodeBarCodeString error " + exp.Message.ToString() + Environment.NewLine;
                                    }
                                    //recordFields[111].SetValue((DateTime)h.ECMSTDDateLastAccessed);
                                    //recordFields[111].SetValue((TrimDateTime)h.ECMSTDDateLastAccessed);

                                    if (h.ECMSTDDocumentDate != null) { recordFields[1].SetValue((TrimDateTime)h.ECMSTDDocumentDate); } else { recordFields[5].ClearValue(); }//ECMApplicationRAM_Application_Number
                                                                                                                                                                              //Auto generated
                                    if (h.ECMApplicationRAM_Application_Number != null) { recordFields[5].SetValue(h.ECMApplicationRAM_Application_Number); } else { recordFields[5].ClearValue(); }//ECMApplicationRAM_Application_Number
                                    if (h.ECMBarCodeBarCodeString != null) { recordFields[6].SetValue(h.ECMBarCodeBarCodeString); } else { recordFields[6].ClearValue(); }//ECMBarCodeBarCodeString
                                    if (h.ECMBCSRetention_Period != null) { recordFields[7].SetValue(h.ECMBCSRetention_Period); } else { recordFields[7].ClearValue(); }//ECMBCSRetention_Period
                                    if (h.ECMCaseCaseDescription != null) { recordFields[8].SetValue(h.ECMCaseCaseDescription); } else { recordFields[8].ClearValue(); }//ECMCaseCaseDescription
                                    if (h.ECMCaseCaseName != null) { recordFields[9].SetValue(h.ECMCaseCaseName); } else { recordFields[9].ClearValue(); }//ECMCaseCaseName
                                    if (h.ECMCaseCaseNumber != null) { recordFields[10].SetValue(h.ECMCaseCaseNumber); } else { recordFields[10].ClearValue(); }//ECMCaseCaseNumber
                                    if (h.ECMCaseCaseOfficer != null) { recordFields[11].SetValue(h.ECMCaseCaseOfficer); } else { recordFields[11].ClearValue(); }//ECMCaseCaseOfficer
                                    if (h.ECMCaseCaseStatus != null) { recordFields[12].SetValue(h.ECMCaseCaseStatus); } else { recordFields[12].ClearValue(); }//ECMCaseCaseStatus
                                    if (h.ECMCaseCaseType != null) { recordFields[13].SetValue(h.ECMCaseCaseType); } else { recordFields[13].ClearValue(); }//ECMCaseCaseType
                                    if (h.ECMCaseEndDate != null) { recordFields[14].SetValue(h.ECMCaseEndDate); } else { recordFields[14].ClearValue(); }//ECMCaseEndDate
                                    if (h.ECMCaseStartDate != null) { recordFields[15].SetValue(h.ECMCaseStartDate); } else { recordFields[15].ClearValue(); }//ECMCaseStartDate
                                    if (h.ECMCorrespondentDescription != null) { recordFields[16].SetValue(h.ECMCorrespondentDescription); } else { recordFields[16].ClearValue(); }//ECMCorrespondentDescription
                                                                                                                                                                                    //if (h.ECMECMDriveID != null) { recordFields[17].SetValue(h.ECMECMDriveID); } else { recordFields[17].ClearValue(); }//ECMECMDriveID
                                    if (h.ECMEmployeeCommencement_Date != null) { recordFields[18].SetValue(h.ECMEmployeeCommencement_Date); } else { recordFields[18].ClearValue(); }//ECMEmployeeCommencement_Date
                                    if (h.ECMEmployeeEmployee_Given_Name != null) { recordFields[19].SetValue(h.ECMEmployeeEmployee_Given_Name); } else { recordFields[19].ClearValue(); }//ECMEmployeeEmployee_Given_Name
                                    if (h.ECMEmployeeEmployee_Surname != null) { recordFields[20].SetValue(h.ECMEmployeeEmployee_Surname); } else { recordFields[20].ClearValue(); }//ECMEmployeeEmployee_Surname
                                    if (h.ECMEmployeeEmployeeCode != null) { recordFields[21].SetValue(h.ECMEmployeeEmployeeCode); } else { recordFields[21].ClearValue(); }//ECMEmployeeEmployeeCode
                                    if (h.ECmEmployeeHR_Case_Description != null) { recordFields[22].SetValue(h.ECmEmployeeHR_Case_Description); } else { recordFields[22].ClearValue(); }//ECmEmployeeHR_Case_Description
                                    if (h.ECMEmployeePreferred_Name != null) { recordFields[23].SetValue(h.ECMEmployeePreferred_Name); } else { recordFields[23].ClearValue(); }//ECMEmployeePreferred_Name
                                    if (h.ECMEmployeeStatus != null) { recordFields[24].SetValue(h.ECMEmployeeStatus); } else { recordFields[24].ClearValue(); }//ECMEmployeeStatus
                                    if (h.ECMEmployeeTermination_Date != null) { recordFields[25].SetValue(h.ECMEmployeeTermination_Date); } else { recordFields[25].ClearValue(); }//ECMEmployeeTermination_Date
                                    if (h.ECMFileName != null) { recordFields[26].SetValue(h.ECMFileName); } else { recordFields[26].ClearValue(); }//ECMFileName
                                    if (h.ECMFileNetFoldername != null) { recordFields[27].SetValue(h.ECMFileNetFoldername); } else { recordFields[27].ClearValue(); }//ECMFileNetFoldername
                                    if (h.ECMHR_Case_Code != null) { recordFields[28].SetValue(h.ECMHR_Case_Code); } else { recordFields[28].ClearValue(); }//ECMHR_Case_Code
                                    if (h.ECMInfoExpertLevel1_Name != null) { recordFields[29].SetValue(h.ECMInfoExpertLevel1_Name); } else { recordFields[29].ClearValue(); }//ECMInfoExpertLevel1_Name
                                    if (h.ECMInfoExpertLevel2_Name != null) { recordFields[30].SetValue(h.ECMInfoExpertLevel2_Name); } else { recordFields[30].ClearValue(); }//ECMInfoExpertLevel2_Name
                                    if (h.ECMInfoExpertLevel3_Description != null) { recordFields[31].SetValue(h.ECMInfoExpertLevel3_Description); } else { recordFields[31].ClearValue(); }//ECMInfoExpertLevel3_Description
                                    if (h.ECMInfoExpertLevel3_Name != null) { recordFields[32].SetValue(h.ECMInfoExpertLevel3_Name); } else { recordFields[32].ClearValue(); }//ECMInfoExpertLevel3_Name
                                    if (h.ECMInfoExpertLevel4_Description != null) { recordFields[33].SetValue(h.ECMInfoExpertLevel4_Description); } else { recordFields[33].ClearValue(); }//ECMInfoExpertLevel4_Description
                                    if (h.ECMInfoExpertLevel4_Name != null) { recordFields[34].SetValue(h.ECMInfoExpertLevel4_Name); } else { recordFields[34].ClearValue(); }//ECMInfoExpertLevel4_Name
                                    if (h.ECMInfoExpertLevel5_Description != null) { recordFields[35].SetValue(h.ECMInfoExpertLevel5_Description); } else { recordFields[35].ClearValue(); }//ECMInfoExpertLevel5_Description
                                    if (h.ECMInfoExpertLevel5_Name != null) { recordFields[36].SetValue(h.ECMInfoExpertLevel5_Name); } else { recordFields[36].ClearValue(); }//ECMInfoExpertLevel5_Name
                                    if (h.ECMInfringementInfringementID != null) { recordFields[37].SetValue(h.ECMInfringementInfringementID); } else { recordFields[37].ClearValue(); }//ECMInfringementInfringementID
                                    if (h.ECMNotes != null) { recordFields[38].SetValue(h.ECMNotes); } else { recordFields[38].ClearValue(); }//ECMNotes
                                    if (h.ECMPositionClose_Date != null) { recordFields[39].SetValue(h.ECMPositionClose_Date); } else { recordFields[39].ClearValue(); }//ECMPositionClose_Date
                                    if (h.ECMPositionOpen_Date != null) { recordFields[40].SetValue(h.ECMPositionOpen_Date); } else { recordFields[40].ClearValue(); }//ECMPositionOpen_Date
                                    if (h.ECMPositionPosition_ID != null) { recordFields[41].SetValue(h.ECMPositionPosition_ID); } else { recordFields[41].ClearValue(); }//ECMPositionPosition_ID
                                    if (h.ECMPositionPosition_Title != null) { recordFields[42].SetValue(h.ECMPositionPosition_Title); } else { recordFields[42].ClearValue(); }//ECMPositionPosition_Title
                                    if (h.ECMPositionVacancy_ID != null) { recordFields[43].SetValue(h.ECMPositionVacancy_ID); } else { recordFields[43].ClearValue(); }//ECMPositionVacancy_ID
                                    if (h.ECMPositionVacancy_Status != null) { recordFields[44].SetValue(h.ECMPositionVacancy_Status); } else { recordFields[44].ClearValue(); }//ECMPositionVacancy_Status
                                    if (h.ECMProjectAndContractProject_Name != null) { recordFields[45].SetValue(h.ECMProjectAndContractProject_Name); } else { recordFields[45].ClearValue(); }//ECMProjectAndContractProject_Name
                                    if (h.ECMProjectAndContractContract_or_Activity_Desc != null) { recordFields[46].SetValue(h.ECMProjectAndContractContract_or_Activity_Desc); } else { recordFields[46].ClearValue(); }//ECMProjectAndContractContract_or_Activity_Desc
                                    if (h.ECMProjectAndContractContract_or_Activity_Name != null) { recordFields[47].SetValue(h.ECMProjectAndContractContract_or_Activity_Name); } else { recordFields[47].ClearValue(); }//ECMProjectAndContractContract_or_Activity_Name
                                    if (h.ECMProjectAndContractContract_or_Activity_No != null) { recordFields[48].SetValue(h.ECMProjectAndContractContract_or_Activity_No); } else { recordFields[48].ClearValue(); }//ECMProjectAndContractContract_or_Activity_No
                                    if (h.ECMProjectAndContractEnd_Date != null) { recordFields[49].SetValue(h.ECMProjectAndContractEnd_Date); } else { recordFields[49].ClearValue(); }//ECMProjectAndContractEnd_Date
                                    if (h.ECMProjectAndContractProject_Description != null) { recordFields[50].SetValue(h.ECMProjectAndContractProject_Description); } else { recordFields[50].ClearValue(); }//ECMProjectAndContractProject_Description
                                    if (h.ECMProjectAndContractProject_End_Date != null) { recordFields[51].SetValue(h.ECMProjectAndContractProject_End_Date); } else { recordFields[51].ClearValue(); }//ECMProjectAndContractProject_End_Date
                                    if (h.ECMProjectAndContractProject_Officer != null) { recordFields[52].SetValue(h.ECMProjectAndContractProject_Officer); } else { recordFields[52].ClearValue(); }//ECMProjectAndContractProject_Officer
                                    if (h.ECMProjectAndContractProject_Start_Date != null) { recordFields[53].SetValue(h.ECMProjectAndContractProject_Start_Date); } else { recordFields[53].ClearValue(); }//ECMProjectAndContractProject_Start_Date
                                    if (h.ECMProjectAndContractProjectNo != null) { recordFields[54].SetValue(h.ECMProjectAndContractProjectNo); } else { recordFields[54].ClearValue(); }//ECMProjectAndContractProjectNo
                                    if (h.ECMProjectAndContractResponsible_Officer != null) { recordFields[55].SetValue(h.ECMProjectAndContractResponsible_Officer); } else { recordFields[55].ClearValue(); }//ECMProjectAndContractResponsible_Officer
                                    if (h.ECMProjectAndContractStart_Date != null) { recordFields[56].SetValue(h.ECMProjectAndContractStart_Date); } else { recordFields[56].ClearValue(); }//ECMProjectAndContractStart_Date
                                    if (h.ECMProjectAndContractStatus != null) { recordFields[57].SetValue(h.ECMProjectAndContractStatus); } else { recordFields[57].ClearValue(); }//ECMProjectAndContractStatus
                                    if (h.ECMPropertyHouse_No != null) { recordFields[58].SetValue(h.ECMPropertyHouse_No); } else { recordFields[58].ClearValue(); }//ECMPropertyHouse_No
                                    if (h.ECMPropertyHouse_No_Suffix != null) { recordFields[59].SetValue(h.ECMPropertyHouse_No_Suffix); } else { recordFields[59].ClearValue(); }//ECMPropertyHouse_No_Suffix
                                    if (h.ECMPropertyHouse_No_To != null) { recordFields[60].SetValue(h.ECMPropertyHouse_No_To); } else { recordFields[60].ClearValue(); }//ECMPropertyHouse_No_To
                                    if (h.ECMPropertyHouse_No_To_Suffix != null) { recordFields[61].SetValue(h.ECMPropertyHouse_No_To_Suffix); } else { recordFields[61].ClearValue(); }//ECMPropertyHouse_No_To_Suffix
                                    if (h.ECMPropertyLocality_Name != null) { recordFields[62].SetValue(h.ECMPropertyLocality_Name); } else { recordFields[62].ClearValue(); }//ECMPropertyLocality_Name
                                    if (h.ECMPropertyPostcode != null) { recordFields[63].SetValue(h.ECMPropertyPostcode); } else { recordFields[63].ClearValue(); }//ECMPropertyPostcode
                                    if (h.ECMPropertyProperty_Name != null) { recordFields[64].SetValue(h.ECMPropertyProperty_Name); } else { recordFields[64].ClearValue(); }//ECMPropertyProperty_Name
                                    if (h.ECMPropertyProperty_No != null) { recordFields[65].SetValue(h.ECMPropertyProperty_No); } else { recordFields[65].ClearValue(); }//ECMPropertyProperty_No
                                    if (h.ECMPropertyStreet_Name != null) { recordFields[66].SetValue(h.ECMPropertyStreet_Name); } else { recordFields[66].ClearValue(); }//ECMPropertyStreet_Name
                                    if (h.ECMPropertyUnit_No != null) { recordFields[67].SetValue(h.ECMPropertyUnit_No); } else { recordFields[67].ClearValue(); }//ECMPropertyUnit_No
                                    if (h.ECMPropertyUnit_No_Suffix != null) { recordFields[68].SetValue(h.ECMPropertyUnit_No_Suffix); } else { recordFields[68].ClearValue(); }//ECMPropertyUnit_No_Suffix
                                                                                                                                                                                //if (h.ECMreference != null) { recordFields[69].SetValue(h.ECMreference); } else { recordFields[69].ClearValue(); }//ECMreference
                                                                                                                                                                                //if (h.ECMRelatedDocDocSetID != null) { recordFields[70].SetValue(h.ECMRelatedDocDocSetID); } else { recordFields[70].ClearValue(); }//ECMRelatedDocDocSetID
                                    if (h.ECMSTDAllRevTitle != null) { recordFields[71].SetValue(h.ECMSTDAllRevTitle); } else { recordFields[71].ClearValue(); }//ECMSTDAllRevTitle
                                    if (h.ECMSTDApplicationNo != null) { recordFields[72].SetValue(h.ECMSTDApplicationNo); } else { recordFields[72].ClearValue(); }//ECMSTDApplicationNo
                                    if (h.ECMSTDBusinessCode != null) { recordFields[73].SetValue(h.ECMSTDBusinessCode); } else { recordFields[73].ClearValue(); }//ECMSTDBusinessCode
                                    if (h.ECMSTDClassName != null) { recordFields[74].SetValue(h.ECMSTDClassName); } else { recordFields[74].ClearValue(); }//ECMSTDClassName
                                    if (h.ECMSTDCorrespondant != null) { recordFields[75].SetValue(h.ECMSTDCorrespondant); } else { recordFields[75].ClearValue(); }//ECMSTDCorrespondant
                                    if (h.ECMSTDCustomerRequest != null) { recordFields[76].SetValue(h.ECMSTDCustomerRequest); } else { recordFields[76].ClearValue(); }//ECMSTDCustomerRequest
                                    if (h.ECMSTDDateLastAccessed != null) { recordFields[77].SetValue(h.ECMSTDDateLastAccessed); } else { recordFields[77].ClearValue(); }//ECMSTDDateLastAccessed
                                    if (h.ECMSTDDateReceived != null) { recordFields[78].SetValue(h.ECMSTDDateReceived); } else { recordFields[78].ClearValue(); }//ECMSTDDateReceived
                                    if (h.ECMSTDDateRegistered != null) { recordFields[79].SetValue(h.ECMSTDDateRegistered); } else { recordFields[79].ClearValue(); }//ECMSTDDateRegistered
                                    if (h.ECMSTDDeclaredDate != null) { recordFields[80].SetValue(h.ECMSTDDeclaredDate); } else { recordFields[80].ClearValue(); }//ECMSTDDeclaredDate
                                    if (h.ECMSTDDestructionDue != null) { recordFields[81].SetValue(h.ECMSTDDestructionDue); } else { recordFields[81].ClearValue(); }//ECMSTDDestructionDue
                                    if (h.ECMSTDDocumentDate != null) { recordFields[82].SetValue(h.ECMSTDDocumentDate); } else { recordFields[82].ClearValue(); }//ECMSTDDocumentDate
                                    if (h.ECMSTDDocumentType != null) { recordFields[83].SetValue(h.ECMSTDDocumentType); } else { recordFields[83].ClearValue(); }//ECMSTDDocumentType
                                    if (h.ECMSTDExternalReference != null) { recordFields[84].SetValue(h.ECMSTDExternalReference); } else { recordFields[84].ClearValue(); }//ECMSTDExternalReference
                                    if (h.ECMSTDInfringementNo != null) { recordFields[85].SetValue(h.ECMSTDInfringementNo); } else { recordFields[85].ClearValue(); }//ECMSTDInfringementNo
                                    if (h.ECMSTDInternalReference != null) { recordFields[86].SetValue(h.ECMSTDInternalReference); } else { recordFields[86].ClearValue(); }//ECMSTDInternalReference
                                    if (h.ECMSTDInternalScanRequest != null) { recordFields[87].SetValue(h.ECMSTDInternalScanRequest); } else { recordFields[87].ClearValue(); }//ECMSTDInternalScanRequest
                                    if (h.ECMSTDJobNo != null) { recordFields[88].SetValue(h.ECMSTDJobNo); } else { recordFields[88].ClearValue(); }//ECMSTDJobNo
                                    if (h.ECMSTDOtherReferences != null) { recordFields[89].SetValue(h.ECMSTDOtherReferences); } else { recordFields[89].ClearValue(); }//ECMSTDOtherReferences
                                    if (h.ECMSTDPropertyNo != null) { recordFields[90].SetValue(h.ECMSTDPropertyNo); } else { recordFields[90].ClearValue(); }//ECMSTDPropertyNo
                                    if (h.ECMSTDSummaryText != 0) { recordFields[91].SetValue(h.ECMSTDSummaryText); } else { recordFields[91].ClearValue(); }//ECMSTDSummaryText
                                    if (h.ECMStreetsLocality != null) { recordFields[92].SetValue(h.ECMStreetsLocality); } else { recordFields[92].ClearValue(); }//ECMStreetsLocality
                                    if (h.ECMStreetsPostcode != null) { recordFields[93].SetValue(h.ECMStreetsPostcode); } else { recordFields[93].ClearValue(); }//ECMStreetsPostcode
                                    if (h.ECMStreetsStreet_Name != null) { recordFields[94].SetValue(h.ECMStreetsStreet_Name); } else { recordFields[94].ClearValue(); }//ECMStreetsStreet_Name
                                    if (h.ECMUserGroupsDescription != null) { recordFields[95].SetValue(h.ECMUserGroupsDescription); } else { recordFields[95].ClearValue(); }//ECMUserGroupsDescription
                                    if (h.ECMUserGroupsExtNo != null) { recordFields[96].SetValue(h.ECMUserGroupsExtNo); } else { recordFields[96].ClearValue(); }//ECMUserGroupsExtNo
                                    if (h.ECMUserGroupsGivenName != null) { recordFields[97].SetValue(h.ECMUserGroupsGivenName); } else { recordFields[97].ClearValue(); }//ECMUserGroupsGivenName
                                    if (h.ECMUserGroupsOrgEmail != null) { recordFields[98].SetValue(h.ECMUserGroupsOrgEmail); } else { recordFields[98].ClearValue(); }//ECMUserGroupsOrgEmail
                                    if (h.ECMUserGroupsSurname != null) { recordFields[99].SetValue(h.ECMUserGroupsSurname); } else { recordFields[99].ClearValue(); }//ECMUserGroupsSurname
                                    if (h.ECMVolumeFilename != null) { recordFields[100].SetValue(h.ECMVolumeFilename); } else { recordFields[100].ClearValue(); }//ECMVolumeFilename
                                    if (h.ECMVolumeLastModified != null) { recordFields[101].SetValue(h.ECMVolumeLastModified); } else { recordFields[101].ClearValue(); }//ECMVolumeLastModified
                                    if (h.ECMVolumeMedia != null) { recordFields[102].SetValue(h.ECMVolumeMedia); } else { recordFields[102].ClearValue(); }//ECMVolumeMedia
                                    if (h.ECMVolumePrimaryCacheName != null) { recordFields[103].SetValue(h.ECMVolumePrimaryCacheName); } else { recordFields[103].ClearValue(); }//ECMVolumePrimaryCacheName
                                    if (h.ECMVolumeRenditionTypeName != null) { recordFields[104].SetValue(h.ECMVolumeRenditionTypeName); } else { recordFields[104].ClearValue(); }//ECMVolumeRenditionTypeName
                                    if (h.ECMVolumeShareDrive != null) { recordFields[105].SetValue(h.ECMVolumeShareDrive); } else { recordFields[105].ClearValue(); }//ECMVolumeShareDrive
                                                                                                                                                                      //if (h.ECMVolumeSize != null) { recordFields[106].SetValue(h.ECMVolumeSize); } else { recordFields[106].ClearValue(); }//ECMVolumeSize
                                    if (h.ECMVolumeStatus != null) { recordFields[107].SetValue(h.ECMVolumeStatus); } else { recordFields[107].ClearValue(); }//ECMVolumeStatus
                                    if (h.ECMVolumeStorageLocation != null) { recordFields[108].SetValue(h.ECMVolumeStorageLocation); } else { recordFields[108].ClearValue(); }//ECMVolumeStorageLocation
                                                                                                                                                                                //if (h.ECMVolumeUpdatable != null) { recordFields[109].SetValue(h.ECMVolumeUpdatable); } else { recordFields[109].ClearValue(); }//ECMVolumeUpdatable

                                    Record importRec = m_loader.NewRecord();
                                    importRec.DateReceived = h.ECMSTDDateReceived;
                                    //SecurityLevel sl = new SecurityLevel(db, 2);
                                    //importRec.SecurityProfile.SecurityLevel = sl;

                                    var imax = lstecm.Where(x => x.DocSetID == h.DocSetID).Max(x => x.ECMSTDVersion);
                                    //if (imax == 1)
                                    //{
                                    //    importRec.SetAsFinal(false);
                                    //    importRec.DateFinalized = (TrimDateTime)h.ECMSTDDateLastAccessed;
                                    //}


                                    InputDocument doc = new InputDocument();
                                    doc.SetAsFile(strStorageLoc);

                                    m_loader.SetDocument(importRec, doc, windowsmode);
                                    //}

                                    m_loader.SetProperties(importRec, recordFields);
                                    m_loader.SubmitRecord(importRec);

                                    if (m_loader.Error.Bad == true)
                                    {
                                        Console.WriteLine("Error -  " + m_loader.Error.Message + " " + m_loader.Error.Source);
                                        //strErrorCol = strErrorCol + "Loader Error: " + h.DocSetID.ToString() + ", " + " ECMDescription: " + h.ECMDescription + " Bcs: " + h.Bcs + " Folder" + h.Folder + " reccont.Uri:" + reccont.Uri.ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                                    }
                                    else
                                    {
                                        Console.WriteLine("Imported ECM T1 record " + h.DocSetID.ToString() + " Row " + intCount.ToString() + " of " + strTotal);

                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Could no find folder: " + h.DocSetID.ToString());
                                    strErrorCol = strErrorCol + "Cant find folder to set: " + h.DocSetID.ToString() + " - " + m_loader.Error.Message + Environment.NewLine;
                                }
                            }
                            else
                            {
                                Console.WriteLine("File does not exist: " + strStorageLoc);
                                strErrorCol = strErrorCol + "File does not exist: " + strStorageLoc + Environment.NewLine;


                            }
                        }
                        else
                        {
                            Console.WriteLine("Record already Exists " + h.DocSetID.ToString() + " Row " + intCount.ToString() + " of " + strTotal);
                            strErrorCol = strErrorCol + "Record already Exists " + h.DocSetID.ToString() + " Row " + intCount.ToString() + " of " + strTotal + Environment.NewLine;

                        }
                    }
                    catch (Exception exp)
                    {
                        Loggitt("ReadLine Error: " + exp.Message.ToString(), " DocSetID: " + h.DocSetID + ", ECMDescription: " + h.ECMDescription + ", GetString(2): " + h.ECMSTDDocumentDate.ToString() + ", FileLocation: " + h.ECMVolumeFilename + ", STDDateLastAccessed: " + h.ECMSTDDateLastAccessed.ToString() + ", STDDateRegistered: " + h.ECMSTDDateRegistered.ToString() + ", Bcs: " + h.ECMBCSActivity_Name + " - " + h.ECMBCSFunction_Name + ", Folder:" + h.ECMBCSSubject_Name);
                    }
                }

                stopwatch.Stop();
                    //Console.WriteLine("Process import batch? Enter to continue.");
                    //Console.ReadLine();
                    
                    m_loader.ProcessAccumulatedData();
                
                    Int64 runHistoryUri = m_loader.RunHistoryUri;


                    Console.WriteLine("Processing complete ...");
                    m_loader.EndRun();

                    TimeSpan ts = stopwatch.Elapsed;

                    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Console.WriteLine("Process Time " + elapsedTime);
                    var sdsd = m_loader.ErrorMessage;
                    // just for interest, lets look at the origin history object and output what it did
                    OriginHistory hist = new OriginHistory(db, runHistoryUri);
                
                    Console.WriteLine("Number of records created ..... " + System.Convert.ToString(hist.RecordsCreated));
                    Loggitt("Bulk loading run:" + Environment.NewLine + "Number of records created: " + hist.RecordsCreated.ToString() + Environment.NewLine + "Time to run: " + elapsedTime + Environment.NewLine + "Loader errors: " + sdsd + Environment.NewLine + "Records in error: : " + hist.RecordsInError.ToString() + Environment.NewLine + Environment.NewLine + strErrorCol, "BulkLoader Log");

            }
            catch (Exception exp)
            {
                Loggitt("Main Error: " + exp.Message.ToString(), "Error");
            }
            return true;
        }

        private string FindMappedStorage(string eCMVolumeStorageLocation)
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
            switch(skjfdh)
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
