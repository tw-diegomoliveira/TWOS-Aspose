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
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.ParagraphFormat.StyleName = "qTitle";
            builder.Writeln("INSTRUCTIONS");
            builder.Writeln();

            foreach (var ele in query)
            {
                builder.ParagraphFormat.StyleName = "qInstruction";
                builder.Writeln(ele.Element("text").Value);
            }
        }
    }
}