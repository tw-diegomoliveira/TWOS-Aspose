using Aspose.Words;
using System.Collections.Generic;
using System.Xml.Linq;

namespace com.truewindglobal.aspose
{
    internal class ConnectedQuestions
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            builder.ListFormat.List = builder.Document.Lists[1];
            builder.ListFormat.ListLevelNumber = 1;
            foreach (var q in query)
            {
                builder.Writeln(q.Element("wording").Value);
            }

        }
    }
}