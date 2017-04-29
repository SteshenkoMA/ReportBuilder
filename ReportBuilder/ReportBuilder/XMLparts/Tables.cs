using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder.XMLparts
{
    public class Tables
    {
        [XmlElement("Table")]
        public List<Table> data { get; set; }
    }
}
