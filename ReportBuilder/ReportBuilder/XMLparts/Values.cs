using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder.XMLparts
{
    public class Values
    {
        [XmlElement("Value")]
        public List<Value> data { get; set; }
    }
}
