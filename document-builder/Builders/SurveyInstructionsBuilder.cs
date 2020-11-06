using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyInstructionsBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            Section section = new Section(builder.Document);
            builder.Document.AppendChild(section);
            Body body = new Body(builder.Document);
            section.AppendChild(body);
            var i = 0;
            foreach (var ele in query)
            {
                i += 1;
                var para = new Paragraph(builder.Document);
                builder.Document.LastSection.Body.AppendChild(para);
                builder.MoveToDocumentEnd();
                Console.WriteLine("INSTRUCTIONS");
                Console.WriteLine(ele.Element("text").Value);
                builder.Write(ele.Element("text").Value);
            }
        }
    }
}
