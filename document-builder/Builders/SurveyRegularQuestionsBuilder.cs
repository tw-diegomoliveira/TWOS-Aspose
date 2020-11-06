using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Drawing;
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
                var qId = ele.Attribute("id").Value;
                var qType = ele.Attribute("answerType").Value;
                var qSubtypeType = ele.Attribute("answerSubType").Value;
#if false
                Console.WriteLine("\t" + ele.Element("wording").Value);
                Console.WriteLine("\t" + ele.Element("description").Value);
#endif
                //Body body = new Body(builder.Document);              
                //Paragraph para = new Paragraph(builder.Document);
                //body.AppendChild(para);

                //Run run = new Run(builder.Document);
                //run.Text = ele.Element("wording").Value;
                //para.AppendChild(run);
                builder.Writeln(ele.Element("wording").Value);
                //builder.Writeln(ele.Element("answer").Element("content").Value);

                //section.AppendChild(para);

                //builder.Writeln(ele.Element("answer").Element("content").Value);
                var style = @"<style>body { font-family:Whitney HTF Book !important; }</style>";

                switch (qType)
                {
                    
                    case "Text":
                        //Console.WriteLine("\t\t\t" + ele.Element("answer").Element("content").Value);
                        builder.InsertHtml(style + ele.Element("answer").Element("content").Value);
                        break;
                    case "Yes/No":
                        //Console.WriteLine("\t\t\t" + ele.Element("answer").Element("content").Value);
                        builder.InsertHtml(style + ele.Element("answer").Element("content").Value);
                        break;

                    case "Upload":
                        builder.InsertHtml(style + "<p>Upload type</p>");
                        break;
                    case "Grouped":
                        //"//section[@id=11]/rQuestion[@answerType="Grouped"]/answer/answerDataExtended"
                        builder.InsertHtml(style + "<p>Grouped type</p>");
                        break;
                    case "Table":
                        if (qSubtypeType == "Standard Table")
                        {
                            //cols(header) -> "//self::rQuestion[@id=119][@answerType='Table']/answer/answerDataExtended[@order=1]/wording"
                            //rows(content) -> "//self::rQuestion[@id=119][@answerType='Table']/answer/answerDataExtended[@order=1]/content"
                            var qc = ele.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table']/answer/answerDataExtended[@order=1]/wording");
                            var qr = ele.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table']/answer/answerDataExtended[@order=1]/content");
                            
                            builder.StartTable();
                            foreach (var e in qc)
                            {
                                builder.InsertCell();
                                // Some special features for the header row.
                                builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                builder.Write(e.Value);
                            }
                            builder.EndRow();
                            
                            foreach (var e in qr)
                            {
                                // Set features for the other rows and cells.
                                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                builder.CellFormat.Width = 100.0;
                                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

                                // Reset height and define a different height rule for table body
                                builder.RowFormat.Height = 30.0;
                                builder.RowFormat.HeightRule = HeightRule.Auto;

                                builder.InsertCell();
                                builder.Write(e.Value);
                            }
                            builder.EndRow();
                            builder.EndTable();
                        }
                        builder.InsertHtml(style + "<p>Table type</p>");
                        break;
                    default:
                        break;
                }
                    
                //get answers
            }

        }
    }
}
