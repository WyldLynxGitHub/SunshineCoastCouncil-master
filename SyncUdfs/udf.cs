using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SyncUdfs
{
    public class udf
    {
        [XmlRoot(ElementName = "ABBREVIATION")]
        public class ABBREVIATION
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
            [XmlText]
            public string Text { get; set; }
        }

        [XmlRoot(ElementName = "CURRENCYSYMBOL")]
        public class CURRENCYSYMBOL
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
        }

        [XmlRoot(ElementName = "DEFAULTSTRING")]
        public class DEFAULTSTRING
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
        }

        [XmlRoot(ElementName = "EXTERNALID")]
        public class EXTERNALID
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
        }

        [XmlRoot(ElementName = "FORMAT")]
        public class FORMAT
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
            [XmlText]
            public string Text { get; set; }
        }

        [XmlRoot(ElementName = "LENGTH")]
        public class LENGTH
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
            [XmlText]
            public string Text { get; set; }
        }

        [XmlRoot(ElementName = "LOWERLIMIT")]
        public class LOWERLIMIT
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
            [XmlText]
            public string Text { get; set; }
        }

        [XmlRoot(ElementName = "MASK")]
        public class MASK
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
        }

        [XmlRoot(ElementName = "NAME")]
        public class NAME
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
            [XmlText]
            public string Text { get; set; }
        }

        [XmlRoot(ElementName = "NOTES")]
        public class NOTES
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
            [XmlText]
            public string Text { get; set; }
        }

        [XmlRoot(ElementName = "PLUGINID")]
        public class PLUGINID
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
        }

        [XmlRoot(ElementName = "UPPERLIMIT")]
        public class UPPERLIMIT
        {
            [XmlAttribute(AttributeName = "propId")]
            public string PropId { get; set; }
        }

        [XmlRoot(ElementName = "ADDITIONALFIELD")]
        public class ADDITIONALFIELD
        {
            [XmlElement(ElementName = "ABBREVIATION")]
            public ABBREVIATION ABBREVIATION { get; set; }
            [XmlElement(ElementName = "CURRENCYSYMBOL")]
            public CURRENCYSYMBOL CURRENCYSYMBOL { get; set; }
            [XmlElement(ElementName = "DEFAULTSTRING")]
            public DEFAULTSTRING DEFAULTSTRING { get; set; }
            [XmlElement(ElementName = "EXTERNALID")]
            public EXTERNALID EXTERNALID { get; set; }
            [XmlElement(ElementName = "FORMAT")]
            public FORMAT FORMAT { get; set; }
            [XmlElement(ElementName = "LENGTH")]
            public LENGTH LENGTH { get; set; }
            [XmlElement(ElementName = "LOWERLIMIT")]
            public LOWERLIMIT LOWERLIMIT { get; set; }
            [XmlElement(ElementName = "MASK")]
            public MASK MASK { get; set; }
            [XmlElement(ElementName = "NAME")]
            public NAME NAME { get; set; }
            [XmlElement(ElementName = "NOTES")]
            public NOTES NOTES { get; set; }
            [XmlElement(ElementName = "PLUGINID")]
            public PLUGINID PLUGINID { get; set; }
            [XmlElement(ElementName = "UPPERLIMIT")]
            public UPPERLIMIT UPPERLIMIT { get; set; }
            [XmlAttribute(AttributeName = "uri")]
            public string Uri { get; set; }
        }

        [XmlRoot(ElementName = "TRIM")]
        public class TRIM
        {
            [XmlElement(ElementName = "ADDITIONALFIELD")]
            public List<ADDITIONALFIELD> ADDITIONALFIELD { get; set; }
            [XmlAttribute(AttributeName = "version")]
            public string Version { get; set; }
            [XmlAttribute(AttributeName = "siteID")]
            public string SiteID { get; set; }
            [XmlAttribute(AttributeName = "databaseID")]
            public string DatabaseID { get; set; }
            [XmlAttribute(AttributeName = "dataset")]
            public string Dataset { get; set; }
            [XmlAttribute(AttributeName = "date")]
            public string Date { get; set; }
            [XmlAttribute(AttributeName = "user")]
            public string User { get; set; }
        }
    }
}
