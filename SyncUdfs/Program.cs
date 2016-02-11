using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;
using System.IO;
using System.Xml.Serialization;

namespace SyncUdfs
{
    public class Program
    {
        public class ADDITIONALFIELD
        {
            [XmlElement(ElementName = "ABBREVIATION")]
            public string ABBREVIATION { get; set; }
            [XmlElement(ElementName = "CURRENCYSYMBOL")]
            public string CURRENCYSYMBOL { get; set; }
            [XmlElement(ElementName = "DEFAULTSTRING")]
            public string DEFAULTSTRING { get; set; }
            [XmlElement(ElementName = "EXTERNALID")]
            public string EXTERNALID { get; set; }
            [XmlElement(ElementName = "FORMAT")]
            public string FORMAT { get; set; }
            [XmlElement(ElementName = "LENGTH")]
            public string LENGTH { get; set; }
            [XmlElement(ElementName = "LOWERLIMIT")]
            public string LOWERLIMIT { get; set; }
            [XmlElement(ElementName = "MASK")]
            public string MASK { get; set; }
            [XmlElement(ElementName = "NAME")]
            public string NAME { get; set; }
            [XmlElement(ElementName = "NOTES")]
            public string NOTES { get; set; }
            [XmlElement(ElementName = "PLUGINID")]
            public string PLUGINID { get; set; }
            [XmlElement(ElementName = "UPPERLIMIT")]
            public string UPPERLIMIT { get; set; }
            [XmlAttribute(AttributeName = "uri")]
            public string Uri { get; set; }
        }
        static void Main(string[] args)
        {
            SyncUdfs();
            
        }
        static void SyncUdfs()
        {
            List<ADDITIONALFIELD> lstaf = new List<ADDITIONALFIELD>();
            ADDITIONALFIELD af = new ADDITIONALFIELD();
            af.LENGTH = "65536";
            af.NAME = "ECM Application:RAM_Application_Number";
            lstaf.Add(af);

            af.LENGTH = "65536";
            af.NAME = "ECM Application:RAM_Application_Number1";
            lstaf.Add(af);


            //MySerializableClass myObject;
            // Construct an instance of the XmlSerializer with the type
            // of object that is being deserialized.

            XmlSerializer mySerializer = new XmlSerializer(typeof(List<ADDITIONALFIELD>));
            // To read the file, create a FileStream.

            //TextWriter textWriter = new StreamWriter(@"D:\Users\WyldLynx\Documents\ECMTEST.xml");

            //mySerializer.Serialize(textWriter, lstaf);
            //textWriter.Close();


            //
            FileStream myFileStream = new FileStream(@"D:\Users\WyldLynx\Documents\ECMTEST.xml", FileMode.Open);
            // Call the Deserialize method and cast to the object type.
            lstaf = (List<ADDITIONALFIELD>)mySerializer.Deserialize(myFileStream);

            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                foreach (ADDITIONALFIELD aff in lstaf)
                {
                    if (db.FindTrimObjectByName(BaseObjectTypes.FieldDefinition, aff.NAME) == null)
                    {
                        FieldDefinition fd = new FieldDefinition(db);
                        fd.Name = aff.NAME;


                        switch (Convert.ToInt64(aff.FORMAT))
                        {
                            case 0:
                                fd.Format = UserFieldFormats.String;
                                break;
                            case 1:
                                fd.Format = UserFieldFormats.Number;
                                break;
                            case 2:
                                fd.Format = UserFieldFormats.Boolean;
                                break;
                            case 3:
                                fd.Format = UserFieldFormats.Date;
                                break;
                            case 4:
                                fd.Format = UserFieldFormats.Datetime;
                                break;
                            case 5:
                                fd.Format = UserFieldFormats.Decimal;
                                break;
                            case 6:
                                fd.Format = UserFieldFormats.Text;
                                break;
                            case 7:
                                fd.Format = UserFieldFormats.Currency;
                                break;
                            //case 8:
                            //    fd.Format = UserFieldFormats.
                            //    break;
                            case 9:
                                fd.Format = UserFieldFormats.Object;
                                break;
                            //case 10:
                            //    fd.Format = UserFieldFormats.
                            //    break;
                            case 11:
                                fd.Format = UserFieldFormats.String;
                                break;

                            case 13:
                                fd.Format = UserFieldFormats.Geography;
                                break;
                        }
                        fd.UpperLimit = aff.UPPERLIMIT;
                        fd.Length = Convert.ToInt32(aff.LENGTH);
                        fd.SetNotes("Additional field syncronised automatically", NotesUpdateType.AppendWithUserStamp);
                        RecordType rt = new RecordType(db, 21);

                        fd.Save();
                        fd.SetIsUsedForRecord(rt, true);
                        Console.WriteLine("new additional field added: " + fd.Name);

                    }
                    else
                    {
                        Console.WriteLine("Additional field already exists: " + aff.NAME);
                    }

                }

            }
        }
    }
}
