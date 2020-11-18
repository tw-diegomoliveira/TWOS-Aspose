using Aspose.Words;
using Aspose.Words.Drawing;
using com.truewindglobal.aspose.Models;
using document_builder.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyCoverBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            //builder.Document.RemoveAllChildren();
            //builder.MoveToDocumentEnd();
            //builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.InsertStyleSeparator();

            foreach (var ele in query)
            {
                // Logo
                Shape logo = builder.InsertImage(Resources.logo, width: 71, height: 35);
                logo.WrapType = WrapType.None;
                logo.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
                logo.HorizontalAlignment = HorizontalAlignment.Center;
                logo.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
                logo.Top = 30;
                
                builder.Writeln();
                builder.Writeln();
                builder.Writeln();
                builder.Writeln();
                builder.Writeln();
                builder.Writeln();
                builder.ParagraphFormat.StyleName = "qTitle";
                builder.Writeln("15(c) Questionnaire");
                builder.Writeln(ele.Element("clientName").Value);
                builder.Writeln();
                builder.Writeln();
                builder.ParagraphFormat.Style.Font.Size = 24;
                builder.Writeln("Funds");
                builder.Writeln();
                builder.Writeln();
                builder.ParagraphFormat.StyleName = "qSubtitle";
                builder.Writeln(ele.Element("name").Value);
                builder.ParagraphFormat.StyleName = "qQuote";
                builder.Writeln(ele.Element("adviserName").Value);
            }

        }
    }
}
