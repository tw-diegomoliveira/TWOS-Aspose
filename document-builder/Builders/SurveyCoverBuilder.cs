using Aspose.Words;
using Aspose.Words.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyCoverBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            foreach (var ele in query)
            {
#if debug
                Console.WriteLine("COVER");
                Console.WriteLine(ele);
#endif
                builder.Document.RemoveAllChildren();
                
                Section section = new Section(builder.Document);
                builder.Document.AppendChild(section);

                // Set some properties for the section
                section.PageSetup.SectionStart = SectionStart.NewPage;
                section.PageSetup.PaperSize = PaperSize.Letter;

                Body body = new Body(builder.Document);
                section.AppendChild(body);

                Paragraph titleParagraph = new Paragraph(builder.Document);
                body.AppendChild(titleParagraph);

                titleParagraph.ParagraphFormat.StyleName = "Title";
                titleParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                // Title
                Run titleRun = new Run(builder.Document);
                titleRun.Text = "Questionnaire Export" + ControlChar.CrLf + ControlChar.CrLf;
                titleRun.Font.Color = Color.Black;
                titleRun.Font.Size = 44;
                titleRun.Font.Name = "Whitney HTF";
                titleParagraph.AppendChild(titleRun);

                // Logo
                var logo = new Shape(builder.Document, ShapeType.Rectangle);

                logo.Width = 460;
                logo.Height = 60;

                logo.FillColor = Color.LightGray;
                logo.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                logo.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                logo.WrapType = WrapType.None;
                logo.Stroked = false;
                logo.Top = 140;
                logo.HorizontalAlignment = HorizontalAlignment.Center;
                titleParagraph.AppendChild(logo);

                // Subtitle
                Paragraph subTitleParagraph = new Paragraph(builder.Document);

                subTitleParagraph.ParagraphFormat.StyleName = "Subtitle";
                subTitleParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                Run subTitleRun = new Run(builder.Document);
                subTitleRun.Text = "Second full process test Sept 3rd";
                subTitleRun.Font.Color = Color.Black;
                subTitleRun.Font.Size = 20;
                subTitleRun.Font.Name = "Whitney HTF Semi";
                subTitleParagraph.AppendChild(subTitleRun);
                body.AppendChild(subTitleParagraph);

                //Adviser Name
                Paragraph adviserNameParagraph = new Paragraph(builder.Document);

                adviserNameParagraph.ParagraphFormat.StyleName = "Subtitle";
                adviserNameParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                Run adviserNameRun = new Run(builder.Document);
                adviserNameRun.Text = "Arcadian Asset Management LLC";
                adviserNameRun.Font.Color = Color.Gray;
                adviserNameRun.Font.Size = 20;
                subTitleRun.Font.Name = "Whitney HTF Semi";
                adviserNameParagraph.AppendChild(adviserNameRun);
                body.AppendChild(adviserNameParagraph);

            }

        }
    }
}
