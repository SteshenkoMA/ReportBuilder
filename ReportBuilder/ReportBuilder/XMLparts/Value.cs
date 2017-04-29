using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder.XMLparts
{
    public class Value
    {
        [XmlAttribute("id")]
        public string id { get; set; }
        [XmlElement("path")]
        public string path { get; set; }
    }
}
