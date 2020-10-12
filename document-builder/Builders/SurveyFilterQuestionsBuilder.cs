using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyFilterQuestionsBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            foreach (var ele in query)
            {
                //Filter question
                var id = ele.Attribute("id").Value;
#if false
                Console.Write("\t" + ele.Element("wording").Value);
                Console.WriteLine(ele.Element("answer").Element("content").Value);
#endif
                Paragraph para = new Paragraph(builder.Document);
                builder.Document.LastSection.Body.AppendChild(para);

                Run run = new Run(builder.Document);
                run.Text = ele.Element("wording").Value;
                para.AppendChild(run);

                builder.Writeln(ele.Element("answer").Element("content").Value);

                //Awnsers
                var q = ele.GetElementsUsingXPath("//self::fQuestion[@id='" + id + "']//rQuestion");
                SurveyRegularQuestionsBuilder.Build(builder, q);
                //foreach (var e in q)
                //{
                //    Console.WriteLine("\t\t" + e.Element("wording").Value);
                //    Console.WriteLine("\t\t" + e.Element("description").Value);
                //    //get answers
                //}


            }
        }
    }
}
