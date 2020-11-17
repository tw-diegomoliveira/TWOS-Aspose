using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace com.truewindglobal.aspose.Builders
{
    public class TableOfContentsBuilder
    {
        public static void Build(DocumentBuilder builder)
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.ParagraphFormat.StyleName = "qTitle";
            builder.Writeln("INDEX");
            builder.Writeln();
            builder.InsertTableOfContents("\\o \"1-1\" \\h \\z \\u");
            builder.Document.Styles[StyleIdentifier.Toc1].Font.Name = GlobalProperties.FontName;
        }
    }
}
