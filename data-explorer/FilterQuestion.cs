using Aspose.Words;
using Aspose.Words.Lists;
using System.Collections.Generic;
using System.Xml.Linq;

namespace com.truewindglobal.aspose
{
    internal class FilterQuestion
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            builder.ListFormat.List = builder.Document.Lists[0];
            builder.ListFormat.ListLevelNumber = 0;
            foreach (var q in query)
            {
                builder.Writeln(q.Element("wording").Value);
                var con = q.GetElementsUsingXPath("//self::fQuestion[@id=" + q.Attribute("id").Value + "]/rQuestion");
                ConnectedQuestions.Build(builder, con);
            }
        }
    }
}