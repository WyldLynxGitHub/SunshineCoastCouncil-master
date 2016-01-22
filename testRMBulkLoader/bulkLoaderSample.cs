using HP.HPTRIM.SDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testRMBulkLoader
{
    public class bulkLoaderSample
    {

        Origin m_origin = null;
        BulkDataLoader m_loader = null;
        int m_recordCount = 0;
        long lCont = 0;
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
        bool getNextRecord(PropertyOrFieldValue[] fields)
        {
            // this function would normally be processing some input data source,
            // reading the next item, setting yp the fields array from this input.
            // For the example, we "hard code" some simple property values and
            // provide 5 records.
            m_recordCount++;
            if (m_recordCount > 5)
            {
                return false;
            }
            // a simple title
            String title = "Test Bulk Loaded Record Import #" + System.Convert.ToString(m_recordCount);
            title += ", imported from run number " + System.Convert.ToString(m_loader.RunHistoryUri);
            // some notes
            String notes = "Some notes for the bulk imported record #"
            + System.Convert.ToString(m_recordCount)
            + " data imported at: "
            + TrimDateTime.Now.ToLongDateTimeString();
            // a random date
            TrimDateTime dateCreated = new TrimDateTime(2007, 12, 14, 12, 0, 0);
            // create a location "on the fly" for the Author property
            //PropertyOrFieldValue[] authorFields = new PropertyOrFieldValue[2];
            //authorFields[0] = new PropertyOrFieldValue(PropertyIds.LocationTypeOfLocation);
            //authorFields[1] = new PropertyOrFieldValue(PropertyIds.LocationSortName);
            //authorFields[0].SetValue(LocationType.Position);
            //authorFields[1].SetValue("Position #" + System.Convert.ToString(m_recordCount));
            ////12 12 HP TRIM Bulk Data Importing Programming Guide
            //// if we are rerunning, check to see if it already exists
            //PropertyValue positionName = new PropertyValue(PropertyIds.LocationSortName);
            //positionName.SetValue("Position #" + System.Convert.ToString(m_recordCount));
            //Int64 authorUri = m_loader.FindLocation(positionName);
            //if (authorUri == 0)
            //{
            //    // use the loader to setup the new location's properties
            //    Location authorLoc = m_loader.NewLocation();
                
            //        m_loader.SetProperties(authorLoc, authorFields);
            //        // now submit the location to the bulk loader queue
            //        authorUri = m_loader.SubmitLocation(authorLoc);
                
            //}
            // now set up the fields array based on these values
            fields[0].SetValue(title);
            fields[1].SetValue(dateCreated);
            fields[2].SetValue(notes);
            fields[3].SetValue(lCont);
            //fields[3].SetValue(authorUri);
            return true;
        }
        public bool run(Database db)
        {
            Console.WriteLine("Initialise BulkDataLoader sample ...");
            //
            Console.WriteLine("Enter Container Uri");
            lCont = Convert.ToInt64(Console.ReadLine());

            // create an origin to use for this sample. Look it up first just so you can rerun the code.
            m_origin = db.FindTrimObjectByName(BaseObjectTypes.Origin, "Bulk Loader Sample") as Origin;
            if (m_origin == null)
            {
                m_origin = new Origin(db, OriginType.Custom1);
                m_origin.Name = "Bulk Loader Sample";
                m_origin.OriginLocation = "n.a";
                // sample code assumes you have a record type defined called "Document"
                m_origin.DefaultRecordType = db.FindTrimObjectByUri(BaseObjectTypes.RecordType,2) as RecordType;
                // don't bother with other origin defaults for the sample, just save it so we can use it
                m_origin.Save();
            }
            // construct a BulkDataLoader for this origin
            m_loader = new BulkDataLoader(m_origin);
            // initialise it as per instructions
            if (!m_loader.Initialise())
            {
                // this sample has no way of dealing with the error.
                Console.WriteLine(m_loader.ErrorMessage);
                return false;
            }
            Console.WriteLine("Starting up an import run ...");
            // the sample is going to do just one run, let's get started...
            // you will need to specify a working folder that works for you (see programming guide)
            m_loader.LogFileName = "Logtest";
            
            m_loader.StartRun("Simulated Input Data", "C:\\junk");
            //13 13 HP TRIM Bulk Data Importing Programming Guide
            // setup the property array that will be used to transfer record metadata
            PropertyOrFieldValue[] recordFields = new PropertyOrFieldValue[4];
            recordFields[0] = new PropertyOrFieldValue(PropertyIds.RecordTitle);
            recordFields[1] = new PropertyOrFieldValue(PropertyIds.RecordDateCreated);
            recordFields[2] = new PropertyOrFieldValue(PropertyIds.RecordNotes);
            recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordContainer);
            //recordFields[3] = new PropertyOrFieldValue(PropertyIds.RecordAuthor);
            // now lets add some records
            while (getNextRecord(recordFields))
            {
                Console.WriteLine("Importing record #" + System.Convert.ToString(m_recordCount) + " ...");
                Record importRec = m_loader.NewRecord();
                
                
                    // set the record properties
                    m_loader.SetProperties(importRec, recordFields);
                    // attach an electronic. The sample just does this for the first record, and uses
                    // a mail message to avoid having to find a document on the hard drive somewhere.
                
                    //InputDocument doc = new InputDocument();
                    
                    //    doc.SetAsMailMessage();
                    //    doc.Subject = "Imported mail message #" + System.Convert.ToString(m_recordCount);
                    //    doc.Content = "Some mail messages have very little content";
                    //    doc.SentDate = new TrimDateTime(2007, 12, 14, 10, 30, 15);
                    //    doc.ReceivedDate = new TrimDateTime(2007, 12, 14, 14, 30, 30);
                    //    EmailParticipant mailContactFrom = new EmailParticipant("random.kindness@oneworld.com","Random Kindness", "SMTP");
                        
                    //        doc.SetAuthor(mailContactFrom);
                        
                    //    EmailParticipant mailContactTo = new EmailParticipant("little.animals@oneworld.com", "Little Animals", "SMTP");
                        
                    //        doc.AddRecipient(MailRecipientType.To, mailContactTo);
                        
                    //    m_loader.SetDocument(importRec, doc, BulkLoaderCopyMode.WindowsCopy);
                    
                    // submit it to the bulk loader
                    m_loader.SubmitRecord(importRec);
                
            }
            // by now the loader has accumulated 5 record inserts and 5 location inserts
            // if you set a breakpoint here and look into the right subfolder of your working folder
            // you will see a bunch of temporary files ready to be loaded into the SQL engine
            // process this batch
            Console.WriteLine("Processing import batch ...");
            m_loader.ProcessAccumulatedData();
            
            // grab a copy of the history object (it is not available in bulk loader after we end the run
            Int64 runHistoryUri = m_loader.RunHistoryUri;
            // we're done, lets end the run now
            Console.WriteLine("Processing complete ...");
            m_loader.EndRun();

            // just for interest, lets look at the origin history object and output what it did
            OriginHistory hist = new OriginHistory(db, runHistoryUri);
            Console.WriteLine("Number of records created ..... "
        + System.Convert.ToString(hist.RecordsCreated));
            Console.WriteLine("Number of locations created ... "
            + System.Convert.ToString(hist.LocationsCreated));
            Console.WriteLine("Logfile; " + m_loader.LogFileName);

            return true;
        }
    }
}
