using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder.XMLparts
{
    public class Images
    {
        [XmlElement("Image")]
        public List<Image> data { get; set; }

    }
}
