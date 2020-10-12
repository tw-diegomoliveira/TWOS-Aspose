using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Aspose.Words;
using com.truewindglobal.aspose;
using com.truewindglobal.aspose.Builders;

namespace com.truewindglobal.aspose
{
    class DocumentExport
    {
        static void Main()
        {
            // Applies Aspose.Words Licence from file
            //HelperMethods.ApplyLicence();

            // Creates Word document object
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Loads the entire XML data file
            var xmlData = XElement.Load(GlobalProperties.XMLfile);

            //Document Global Setup
            DocumentIniSettings(builder);

            //extract Cover data
            var query = xmlData.GetElementsUsingXPath("//self::document[@id=90]/cover");

            /// <summary>
            /// Build document Cover page.
            /// </summary>
            SurveyCoverBuilder.Build(builder, query);

            //extract Instructions data
            query = xmlData.GetElementsUsingXPath("//self::document[@id=90]/instructions");

            /// <summary>
            /// Build document Instructions page
            /// </summary>
            SurveyInstructionsBuilder.Build(builder, query);

            //extract Section data
            query = xmlData.GetElementsUsingXPath("//self::document[@id=90]/section");

            //Loop thru the section
            foreach (var ele in query)
            {
                var id = ele.Attribute("id").Value;
                SurveySectionsBuilder.Build(builder, ele);

                //extract filter
                var q = ele.GetElementsUsingXPath("//self::section[@id='"+ id +"']/fQuestion");
                //create filter
                SurveyFilterQuestionsBuilder.Build(builder, q);

                //extract regular
                q = ele.GetElementsUsingXPath("//section[@id='" + id + "']/rQuestion");
                //create regular
                SurveyRegularQuestionsBuilder.Build(builder, q);
            }


            //Set Pagenumber
            //Set TOC

            /// <summary>
            /// Creates/Saves document the extension will provide the format.
            /// </summary>
            doc.Save(GlobalProperties.Filename);
            Console.ReadLine();
        }

        private static void DocumentIniSettings(DocumentBuilder builder)
        {
            builder.Font.Name = GlobalProperties.FontName;
            builder.Font.Size = GlobalProperties.FontSize;
        }
    }
}
