using Aspose.Words;
using Aspose.Words.Lists;
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
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query, int sectionIndex, List list)
        {
            foreach (var ele in query)
            {
                builder.MoveToDocumentEnd();
                builder.InsertStyleSeparator();
                builder.ParagraphFormat.StyleName = "qHeading 2";
                builder.ListFormat.List = list;
                builder.ListFormat.ListLevelNumber = 0;
                builder.Writeln(ele.Element("wording").Value);
                builder.ListFormat.RemoveNumbers();
                builder.ListFormat.List = null;

                if (ele.Element("answer").Element("content").Value == "1")
                {
                    builder.ParagraphFormat.StyleName = "qFilterYes";
                    builder.Writeln("\t>YES");
                }
                else
                {
                    builder.ParagraphFormat.StyleName = "qFilterYes";
                    builder.Writeln("\t>NO");
                }

                //Connected Questions
                var q = ele.GetElementsUsingXPath("//self::fQuestion[@id='" + ele.Attribute("id").Value + "']//rQuestion");
                
                List groupList = builder.Document.Lists.Add(ListTemplate.BulletSquare);
                groupList.ListLevels[0].StartAt = 1;
                builder.ListFormat.ListLevelNumber = 2;

                SurveyRegularQuestionsBuilder.Build(builder, q, true, groupList);
            }
        }
    }
}
