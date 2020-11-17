using Aspose.Words;
using Aspose.Words.Lists;
using System;
using System.Xml.Linq;

namespace com.truewindglobal.aspose
{
    class DataExplorer
    {
        static void Main()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            List listMain = doc.Lists.Add(ListTemplate.NumberArabicDot);
            listMain.IsRestartAtEachSection = true;
            List listSecondary = doc.Lists.Add(ListTemplate.NumberLowercaseLetterDot);
            listSecondary.IsRestartAtEachSection = true;

            var data = XElement.Load(GlobalProperties.Filename);
            var sections = data.GetElementsUsingXPath("//self::section");

            foreach (var s in sections)
            {
                var filter = s.GetElementsUsingXPath("//self::section[@id=" + s.Attribute("id").Value + "]/fQuestion");
                FilterQuestion.Build(builder, filter);

                var regular = s.GetElementsUsingXPath("//self::section[@id=" + s.Attribute("id").Value + "]/rQuestion");
                RegularQuestion.Build(builder, regular);

                builder.InsertBreak(BreakType.SectionBreakNewPage);
            }
            doc.Save(GlobalProperties.DocumentPath + "Lists.pdf", SaveFormat.Pdf);
            //Console.ReadLine();

        }
    }
}
