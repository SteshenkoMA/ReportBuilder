using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder.XMLparts
{
    //В папке XMLpart классы, которые описывают содержимое sourses.xml
    [Serializable]
    [XmlRoot("Sourses")]
    public class Sourses
    {
        [XmlElement(ElementName = "Images", IsNullable = true)]
        public Images images;

        [XmlElement(ElementName = "Tables", IsNullable = true)]
        public Tables tables;

        [XmlElement(ElementName ="Attaches", IsNullable = true)]
        public Attaches attaches;

        [XmlElement(ElementName = "Values", IsNullable = true)]
        public Values values;
    }
        
}
