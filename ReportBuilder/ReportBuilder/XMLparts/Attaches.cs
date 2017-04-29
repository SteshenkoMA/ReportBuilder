using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder.XMLparts
{
    public class Attaches
    {
        [XmlElement("Attach")]
        public List<Attach> data { get; set; }

    }
}
