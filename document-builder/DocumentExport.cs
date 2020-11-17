using System;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Lists;
using com.truewindglobal.aspose.Builders;

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

            List sectionList = doc.Lists.AddCopy(stylesDoc.Lists[5]);
            sectionList.IsRestartAtEachSection = false;

            List lvlOneList = doc.Lists.Add(ListTemplate.NumberArabicDot);
            lvlOneList.IsRestartAtEachSection = true;

            List lvlTwoList = doc.Lists.Add(ListTemplate.NumberLowercaseLetterParenthesis);
            lvlTwoList.IsRestartAtEachSection = true;

            HelperMethods.CopyStyles(stylesDoc, doc);


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

            ///<summary>
            /// Insert TOC
            /// </summary>
            TableOfContentsBuilder.Build(builder);

            //extract Instructions data
            query = xmlData.GetElementsUsingXPath("//self::instruction");

            /// <summary>
            /// Build document Instructions page
            /// </summary>
            SurveyInstructionsBuilder.Build(builder, query);

            //extract Section data
            query = xmlData.GetElementsUsingXPath("//self::document/section");

            //Loop thru the section
            var i = 1;
            foreach (var ele in query)
            {
                //var id = ;
                SurveySectionsBuilder.Build(builder, ele);

                if (i > 1)
                {
                    //extract filter
                    var q = ele.GetElementsUsingXPath("//self::section[@id='" + ele.Attribute("id").Value + "']/fQuestion");
                    // creates a new list
                    List newList = builder.Document.Lists.Add(ListTemplate.NumberArabicDot);
                    newList.ListLevels[0].StartAt = 1;
                    builder.ListFormat.ListLevelNumber = 2;
                    //create filter
                    SurveyFilterQuestionsBuilder.Build(builder, q, i, newList);

                    //extract regular
                    q = ele.GetElementsUsingXPath("//section[@id='" + ele.Attribute("id").Value + "']/rQuestion");
                    //create regular
                    //builder.MoveToDocumentEnd();
                    //builder.Writeln();
                    SurveyRegularQuestionsBuilder.Build(builder, q, false, newList);
                }
                else
                {
                    //extract filter
                    var q = ele.GetElementsUsingXPath("//self::section[@id='" + ele.Attribute("id").Value + "']/fQuestion");
                    SurveyFilterQuestionsBuilder.Build(builder, q, i, lvlOneList);
                    //extract regular
                    q = ele.GetElementsUsingXPath("//section[@id='" + ele.Attribute("id").Value + "']/rQuestion");
                    //create regular
                    //builder.MoveToDocumentEnd();
                    //builder.Writeln();
                    SurveyRegularQuestionsBuilder.Build(builder, q, false, lvlOneList);
                }

                i += 1;
            }

            // Setup Header and Footer
            HeaderAndFooterBuilder.Build(builder);

            //Set TOC
            //TableOfContentsBuilder.Build(builder);
            builder.Document.UpdateFields();

            /// <summary>
            /// Creates/Saves document the extension will provide the format.
            /// </summary>
            FontInfoCollection fontInfos = doc.FontInfos;
            fontInfos.EmbedTrueTypeFonts = true;
            fontInfos.EmbedSystemFonts = false;
            fontInfos.SaveSubsetFonts = false;
            doc.Save(GlobalProperties.outDoc);
            //Console.ReadLine();

        }
        private static void DocumentIniSettings(DocumentBuilder builder)
        {
            builder.Font.Name = GlobalProperties.FontName;
            builder.Font.Size = GlobalProperties.FontSize;

            /// <summary>
            /// Custom styles for the all document
            /// </summary>
#if false
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
#endif
        }
    }
}
