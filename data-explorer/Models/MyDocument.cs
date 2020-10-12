using System.Collections;
using System.Xml.Serialization;

namespace com.truewindglobal.aspose.Models
{
    [XmlRoot("MyDocument", Namespace = "http://www.dotnetcoretutorials.com/namespace")]
    public class MyDocument
    {
        public string MyProperty { get; set; }

        public MyAttributeProperty MyAttributeProperty { get; set; }

        [XmlArray]
        [XmlArrayItem(ElementName = "MyListItem")]
        public IList MyList { get; set; }
    }
    public class MyAttributeProperty
    {
        [XmlAttribute("value")]
        public int Value { get; set; }
    }
}
