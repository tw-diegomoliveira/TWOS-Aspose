using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveySectionsBuilder
    {
        public static void Build(DocumentBuilder builder, XElement element)
        {
#if false
            Console.WriteLine("SECTION");
            Console.WriteLine(query.Attribute("id").Value);
            Console.WriteLine(query.Attribute("name").Value);
            Console.WriteLine();
#endif
            Section section = new Section(builder.Document);
            builder.Document.AppendChild(section);

            // Set some properties for the section
            section.PageSetup.SectionStart = SectionStart.NewPage;
            section.PageSetup.PaperSize = PaperSize.Letter;

            Body body = new Body(builder.Document);
            section.AppendChild(body);

            Paragraph para = new Paragraph(builder.Document);
            body.AppendChild(para);

            // We can set some formatting for the paragraph
            para.ParagraphFormat.StyleName = "Title";
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            Run runId = new Run(builder.Document);
            Run runName = new Run(builder.Document);
            runId.Text = element.Attribute("id").Value;
            para.AppendChild(runId);
            runName.Text = element.Attribute("name").Value;
            para.AppendChild(runName);
        }
    }
}
