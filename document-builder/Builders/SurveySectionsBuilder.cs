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
            //Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyTitleStyle");
            //paraStyle.Font.Bold = true;
            //paraStyle.Font.Size = 28;
            //paraStyle.Font.Name = GlobalProperties.FontName;
            //paraStyle.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            body.AppendChild(para);

            // We can set some formatting for the paragraph
            para.ParagraphFormat.StyleName = "MyHeading1Style";
            //para.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            Run run = new Run(builder.Document)
            {
                Text = "Section " + element.Attribute("id").Value
            };
            para.AppendChild(run);

            para = new Paragraph(builder.Document);
            //paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyNormalStyle");
            //paraStyle.Font.Bold = false;
            //paraStyle.Font.Size = GlobalProperties.FontSize;
            //paraStyle.Font.Name = GlobalProperties.FontName;
            //paraStyle.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            
            para.ParagraphFormat.StyleName = "MyNormalStyle";
            //para.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            body.AppendChild(para);

            run = new Run(builder.Document)
            {
                Text = element.Attribute("name").Value
            };
            para.AppendChild(run);
            builder.MoveToDocumentEnd();
            builder.Writeln();
        }
    }
}
