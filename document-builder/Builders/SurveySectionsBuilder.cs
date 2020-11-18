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
        public static void Build(DocumentBuilder builder, XElement ele)
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Document.LastSection.PageSetup.SectionStart = SectionStart.NewPage;
            builder.Document.LastSection.PageSetup.PaperSize = PaperSize.Letter;

            //builder.ListFormat.List = builder.Document.Lists[0];
            //builder.ListFormat.ListLevelNumber = 0;

            //builder.InsertStyleSeparator();
            builder.ParagraphFormat.StyleName = "qHeading 1";
            builder.Writeln("Section - " + ele.Attribute("name").Value);
            //builder.ListFormat.RemoveNumbers();
            //builder.ListFormat.List = null;
            //builder.ParagraphFormat.StyleName = "qNormal";
            //builder.Writeln(ele.Attribute("name").Value);
        }
    }
}
