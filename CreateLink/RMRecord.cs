using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HP.HPTRIM.SDK;
using System.Threading;

namespace CreateLink
{
    public partial class RMRecord : Form
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
        private TrimMainObject obj = null;
        private Database TrimDb = null;
        public static long intUri;
        public static string strRecNumber;
        private static string httpname = null;
        public RMRecord(TrimMainObject obj)
        {
            InitializeComponent();
            //
            this.obj = obj;
            if (obj != null)
            {
                this.TrimDb = obj.Database;
            }

            Record rec = obj.Database.FindTrimObjectByUri(BaseObjectTypes.Record, obj.Uri) as Record;

            intUri = obj.Uri;
            strRecNumber = rec.Number;
            
            textBox1.Text = strRecNumber;
            //MessageBox.Show("Link Created to clipboard.", "", MessageBoxButtons.OK);
        }
        public static void somethingToRunInThread()
        {

            //System.Windows.Forms.Clipboard.SetText("String to be copied to Clipboard");
            //http://localhost/HPRMDemoServiceAPI/Record/858/file/document
            ///Record/858/file/document
            //string strClipAddress = "<a href="http://localhost/HPRMDemoServiceAPI/Record/" + intUri + "/file/document">Link"</a>";
            //Clipboard.SetText(strClipAddress);
            string link = String.Format(html, "http://localhost/HPRMServiceAPI/Record/" + intUri + "/file/document", httpname);
            Clipboard.SetText(link, TextDataFormat.Html);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            httpname = textBox1.Text;
            Thread clipboardThread = new Thread(somethingToRunInThread);

            clipboardThread.SetApartmentState(ApartmentState.STA);
            clipboardThread.IsBackground = false;
            clipboardThread.Start();
            MessageBox.Show("Record download link copied to clipboard", "Create link", MessageBoxButtons.OK);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr =  MessageBox.Show("Are you sure you want to delete record?", "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(dr==DialogResult.Yes)
            {
                Record rec = obj.Database.FindTrimObjectByUri(BaseObjectTypes.Record, obj.Uri) as Record;
                //
                List<Location> lstAclView = new List<Location>();
                Location lPol = new Location(TrimDb, obj.Database.CurrentUser.Uri);
                lstAclView.Add(lPol);
                TrimAccessControlList acl = rec.AccessControlList;
                acl.set_AccessLocations((int)RecordAccess.ViewRecord, lstAclView.ToArray());
                rec.AccessControlList = acl;
                //
                


                rec.SetNotes("Record set to delete: reason - "+comboBox1.SelectedText, NotesUpdateType.AppendWithNewLine);
                rec.Save();
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
