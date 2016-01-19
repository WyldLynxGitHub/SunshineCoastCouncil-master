using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace CreateLink
{

    public class CreateLink : ITrimAddIn
    {
        private const string html = @"Version:0.9
StartHTML:<<<<<<<1
EndHTML:<<<<<<<2
StartFragment:<<<<<<<3
EndFragment:<<<<<<<4
SourceURL: {0}
<html>
<body>
<!--StartFragment-->
<a href='{0}'>{1}</a>
<!--EndFragment-->
</body>
</html>";
        public static long intUri;
        public static string strRecNumber;
        public override string ErrorMessage
        {
            //get { throw new NotImplementedException(); }
            get { return "SetLink Error"; }
            //Loggitt(ErrorMessage, "Record Lifecycle test: Error");
        }

        public override void ExecuteLink(int cmdId, TrimMainObjectSearch forTaggedObjects)
        {
            //throw new NotImplementedException();
            //Loggitt("Record Uri: " + forTaggedObjects.SkipToUri.ToString(), "Record Lifecycle test: ExecuteLink(int cmdId, TrimMainObjectSearch forTaggedObjects)");
            MessageBox.Show("forTaggedObjects: " + forTaggedObjects.SkipToUri.ToString());

        }

        public override void ExecuteLink(int cmdId, TrimMainObject forObject, ref bool itemWasChanged)
        {
            ////MessageBox.Show("itemWasChanged: " + forObject.Uri.ToString());
            //string strallrecs = null;
            ////throw new NotImplementedException();
            //System.Windows.Forms.Clipboard
            //Loggitt("Record Uri: " + forObject.Uri.ToString(), "Record Lifecycle test: ExecuteLink(int cmdId, TrimMainObject forObject, ref bool itemWasChanged)");
            //using(TrimMainObjectSearch obj = new TrimMainObjectSearch(forObject.Database, BaseObjectTypes.Record))
            //{
            //    obj.SetSearchString("allrelated:" + forObject.Uri);
            //    foreach(Record r in obj)
            //    {
            //        strallrecs += r.Number+Environment.NewLine;
            //    }
            //}
            //MessageBox.Show("All related: " + Environment.NewLine + strallrecs);
            //MessageBox.Show("itemWasChanged: " + forObject.Uri.ToString() + " - " + forObject.WebURL);

            //Clipboard.SetText(forObject.WebURL, TextDataFormat.Html);

            //Clipboard.SetText("http://www.news.com.au", TextDataFormat.Html);
            //
            //
            //Record rec = forObject.Database.FindTrimObjectByUri(BaseObjectTypes.Record, forObject.Uri) as Record;

            //intUri = forObject.Uri;
            //strRecNumber = rec.Number;
            //Thread clipboardThread = new Thread(somethingToRunInThread);
            
            //clipboardThread.SetApartmentState(ApartmentState.STA);
            //clipboardThread.IsBackground = false;
            //clipboardThread.Start();
            //MessageBox.Show("Link Created to clipboard.", "", MessageBoxButtons.OK);

            if(forObject!=null)
            {
                using(RMRecord form = new RMRecord(forObject))
                {
                    form.ShowDialog();
                }
            }
            //try
            //{
            //    if (forTaggedObjects.FastCount > 0)
            //    {
            //        using (Database db = forTaggedObjects.Database)
            //        {
            //            // process of multiple objects was not supported. Only process the first one
            //            foreach (TrimMainObject obj in forTaggedObjects)
            //            {
            //                using (RMRecord form = new RMRecord(obj))
            //                {
            //                    form.ShowDialog();
            //                }
            //                // To do....Processing
            //                obj.Dispose();
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }
        public static void somethingToRunInThread()
        {

            //System.Windows.Forms.Clipboard.SetText("String to be copied to Clipboard");
            //http://localhost/HPRMDemoServiceAPI/Record/858/file/document
            ///Record/858/file/document
            //string strClipAddress = "<a href="http://localhost/HPRMDemoServiceAPI/Record/" + intUri + "/file/document">Link"</a>";
            //Clipboard.SetText(strClipAddress);
            string link = String.Format(html, "http://localhost/HPRMServiceAPI/Record/" + intUri + "/file/document", strRecNumber);
            Clipboard.SetText(link, TextDataFormat.Html);
        }
        public override TrimMenuLink[] GetMenuLinks()
        {
            //throw new NotImplementedException();
            TrimMenuLink[] links = new TrimMenuLink[] { new SetLinkMenu() };
            return links;
        }

        public override void Initialise(Database db)
        {
            //throw new NotImplementedException();
            //Loggitt("Logged on as: "+db.CurrentUser.Name, "Record Lifecycle test: Initialise");        
        }

        public override bool IsMenuItemEnabled(int cmdId, TrimMainObject forObject)
        {
            //throw new NotImplementedException();
            return true;
        }

        public override void PostDelete(TrimMainObject deletedObject)
        {
            throw new NotImplementedException();
        }

        public override void PostSave(TrimMainObject savedObject, bool itemWasJustCreated)
        {
            //throw new NotImplementedException();
            //Loggitt("Record Uri: "+savedObject.Uri, "Record Lifecycle test: PostSave");
        }

        public override bool PreDelete(TrimMainObject modifiedObject)
        {
            //throw new NotImplementedException();
            return true;
        }

        public override bool PreSave(TrimMainObject modifiedObject)
        {
            //Loggitt("Record Uri: " + modifiedObject.Uri, "Record Lifecycle test: PreSave");
            return true;
        }

        public override bool SelectFieldValue(FieldDefinition field, TrimMainObject trimObject, string previousValue)
        {
            //throw new NotImplementedException();
            return false;
        }

        public override void Setup(TrimMainObject newObject)
        {
            //Loggitt(newObject.Name, "Record Lifecycle test: Setup");
        }

        public override bool SupportsField(FieldDefinition field)
        {
            //throw new NotImplementedException();
            return false;
        }

        public override bool VerifyFieldValue(FieldDefinition field, TrimMainObject trimObject, string newValue)
        {
            //throw new NotImplementedException();
            bool isValid = (newValue != "");
            return isValid;
        }

        public static void Loggitt(string msg, string title)
        {
            if (!(System.IO.Directory.Exists(@"C:\Temp\AddinRecordCreateTest\")))
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp\AddinRecordCreateTest\");
            }
            FileStream fs = new FileStream(@"C:\Temp\AddinRecordCreateTest\tstlog_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(@"C:\Temp\AddinRecordCreateTest\tstlog_" + DateTime.Today.ToShortDateString().Replace("/", "") + ".txt", FileMode.Append, FileAccess.Write);
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
    class SetLinkMenu : TrimMenuLink
    {
        public SetLinkMenu()
        {
        }
        // Summary:
        //     Gets a string that should appear as a tool tip when hovering over a menu
        //     item
        public override string Description
        {
            get { return "Select this menu item to save download link to clipboard."; }
        }
        //
        // Summary:
        //     Gets an ID number that identifies this menu item.
        public override int MenuID
        {
            get { return 42; }
        }
        //
        // Summary:
        //     Gets a string that should appear on the context menu.
        public override string Name
        {
            get { return "Clipboard Link"; }
        }
        //
        // Summary:
        //     Gets a boolean value indicating whether this menu item supports TRIM tagged
        //     processing
        public override bool SupportsTagged
        {
            get { return false; }
        }

    };
}
