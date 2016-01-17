using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace TemplateTitle
{
    public class TemplateTitle : TrimEventProcessorAddIn
    {
        public override void ProcessEvent(Database db, TrimEvent evt)
        {
            if(evt.ObjectType == BaseObjectTypes.Record)
            {
                if (evt.EventType == Events.ObjectAdded)
                {
                    try
                    {
                        Record rec = new Record(db, evt.ObjectUri);
                        if (rec.IsElectronic)
                        {
                            var docdetails = rec.ESource;

                            //string Location = "C:\\Program Files\\hello.txt";

                            string FileName = docdetails.Substring(docdetails.LastIndexOf('\\') +
                                1);
                            //
                            int fileExtPos = FileName.LastIndexOf(".");
                            if (fileExtPos >= 0)
                                FileName = FileName.Substring(0, fileExtPos);

                            var rectitle = rec.Title;
                            if (rectitle != FileName)
                            {
                                rec.Title = FileName;
                            }
                            rec.SetNotes("Notes added: DocDetails: " + FileName + " RecTitle: " + rectitle, NotesUpdateType.PrependWithNewLine);
                            rec.Save();
                            db.LogExternalEvent("Successful template title change.", BaseObjectTypes.Record, evt.ObjectUri, true);
                        }
                    }
                    catch (Exception exp)
                    {
                        db.LogExternalEvent("Template title change error - "+exp.Message.ToString(), BaseObjectTypes.Record, evt.ObjectUri, true);
                    }
                }
            }
        }
    }
}
