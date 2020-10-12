using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyRegularQuestionsBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            foreach (var ele in query)
            {
                var qType = ele.Attribute("answerType").Value;
                var qSubtypeType = ele.Attribute("answerSubType").Value;
#if false
                Console.WriteLine("\t" + ele.Element("wording").Value);
                Console.WriteLine("\t" + ele.Element("description").Value);
#endif
                Paragraph para = new Paragraph(builder.Document);
                builder.Document.LastSection.Body.AppendChild(para);

                Run run = new Run(builder.Document);
                run.Text = ele.Element("wording").Value;
                para.AppendChild(run);

                builder.Writeln(ele.Element("answer").Element("content").Value);

                switch (qType)
                {
                    case "Text":
                        //Console.WriteLine("\t\t\t" + ele.Element("answer").Element("content").Value);
                        builder.Writeln(ele.Element("answer").Element("content").Value);
                        break;
                    default:
                        break;
                }
                    
                //get answers
            }

        }
    }
}
