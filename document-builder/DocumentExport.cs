using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using com.truewindglobal.aspose;
using com.truewindglobal.aspose.Builders;
using com.truewindglobal.aspose.Models;

namespace com.truewindglobal.aspose
{
    class DocumentExport
    {
        static void Main()
        {
            // Applies Aspose.Words Licence from file
            HelperMethods.ApplyLicence();

            // Creates Word document object
            var doc = new Document();
            var stylesDoc = new Document(GlobalProperties.stylesDoc);
            var builder = new DocumentBuilder(doc);

            // Loads the entire XML data file
            var xmlData = XElement.Load(GlobalProperties.xmlFile);

            //Document Global Setup
            DocumentIniSettings(builder);
            
            //extract Cover data
            var query = xmlData.GetElementsUsingXPath("//self::cover");

            /// <summary>
            /// Build document Cover page.
            /// </summary>
            SurveyCoverBuilder.Build(builder, query);

            //extract Instructions data
            query = xmlData.GetElementsUsingXPath("//self::instruction");

            /// <summary>
            /// Build document Instructions page
            /// </summary>
            SurveyInstructionsBuilder.Build(builder, query);

            //extract Section data
            query = xmlData.GetElementsUsingXPath("//self::document/section");

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
                builder.MoveToDocumentEnd();
                builder.Writeln();
                SurveyRegularQuestionsBuilder.Build(builder, q);
            }

            // Setup Header and Footer
            HeaderAndFooterBuilder.Build(builder);

            //Set TOC

            /// <summary>
            /// Creates/Saves document the extension will provide the format.
            /// </summary>
            doc.Save(GlobalProperties.outDoc);
            Console.ReadLine();

        }
        private static void DocumentIniSettings(DocumentBuilder builder)
        {
            builder.Font.Name = GlobalProperties.FontName;
            builder.Font.Size = GlobalProperties.FontSize;
            
            /// <summary>
            /// Custom styles for the all document
            /// </summary>
            
            //Title
            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyTitleStyle");
            paraStyle.Font.Bold = true;
            paraStyle.Font.Size = 28;
            paraStyle.Font.Name = "Whitney HTF Semi";
            paraStyle.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            //SubTitle
            paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MySubTitleStyle");
            paraStyle.Font.Bold = true;
            paraStyle.Font.Size = 22;
            paraStyle.Font.Name = "Whitney HTF Semi";
            paraStyle.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            //Heading1
            paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyHeading1Style");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 20;
            paraStyle.Font.Name = "Whitney HTF";
            paraStyle.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            //Normal
            StyleCollection styles = builder.Document.Styles;
            styles.DefaultFont.Name = "Whitney HTF Book";
            styles.DefaultFont.Size = 12;
            styles.Add(StyleType.Paragraph, "MyNormalStyle");
        }
    }
}
