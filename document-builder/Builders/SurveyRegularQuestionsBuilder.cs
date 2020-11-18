using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyRegularQuestionsBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query, bool isConnectedQuestion, List list)
        {
            var style = @"<style>body { font-family:Tahoma !important }</style>";
            foreach (var e in query)
            {
                var qId = e.Attribute("id").Value;
                var qType = e.Attribute("answerType").Value;
                var qSubtypeType = e.Attribute("answerSubType").Value;

                builder.MoveToDocumentEnd();
                builder.InsertStyleSeparator();
                switch (qType)
                {

                //builder.InsertHtml(style + e.Element("answer").Element("content").Value);
                //builder.InsertStyleSeparator();
                    case "Text":
                        builder.ParagraphFormat.StyleName = isConnectedQuestion ? "qHeading 4" : "qHeading 2";
                        //builder.ListFormat.List = list;
                        builder.Writeln(e.Element("wording").Value);
                        //builder.ListFormat.RemoveNumbers();
                        //builder.ListFormat.List = null;
                        builder.ParagraphFormat.StyleName = "qDescription";
                        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level4;
                        builder.Writeln("\t" + e.Element("description").Value);
                        builder.ParagraphFormat.StyleName = "qNormal";
                        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level4;
                        builder.InsertHtml(style + e.Element("answer").Element("content").Value);
                        break;

                    case "Yes/No":
                        builder.ParagraphFormat.StyleName = isConnectedQuestion ? "qHeading 4" : "qHeading 2";
                        //builder.ListFormat.List = list; 
                        builder.Write(e.Element("wording").Value + "\n");
                        //builder.ListFormat.RemoveNumbers();
                        //builder.ListFormat.List = null;
                        builder.ParagraphFormat.StyleName = "qDescription";
                        builder.Writeln(e.Element("description").Value);
                        builder.ParagraphFormat.StyleName = "qNormal";
                        builder.Writeln("\t" + e.Element("answer").Element("content").Value);
                        builder.ListFormat.RemoveNumbers();
                        break;

                    case "Upload":
                        builder.ParagraphFormat.StyleName = isConnectedQuestion ? "qHeading 4" : "qHeading 2";
                        //builder.ListFormat.List = list;
                        builder.Writeln(e.Element("wording").Value);
                        //builder.ListFormat.RemoveNumbers();
                        //builder.Writeln(e.Element("answer").Element("content").Value);
                        builder.Writeln();
                        var qu = e.GetElementsUsingXPath("//self::rQuestion[@answerType='Upload']/answer/binaryData");
                        foreach (var q in qu)
                        {
                            builder.ParagraphFormat.StyleName = "qHeading 4";
                            builder.Writeln(q.Attribute("name").Value);
                        }
                        break;

                    case "Grouped":
                        builder.MoveToDocumentEnd();
                        //builder.InsertStyleSeparator();
                        builder.ParagraphFormat.StyleName = isConnectedQuestion ? "qHeading 4" : "qHeading 2";
                        builder.ListFormat.List = list;
                        builder.Writeln(e.Element("wording").Value);
                        builder.ListFormat.RemoveNumbers();
                        builder.ListFormat.List = null;

                        builder.ParagraphFormat.StyleName = "qDescription";
                        builder.Writeln(e.Element("description").Value);

                        var grouped = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Grouped']/answer/answerDataExtended");

                        List newList = builder.Document.Lists.Add(ListTemplate.NumberLowercaseLetterParenthesis);
                        newList.ListLevels[0].StartAt = 1;
                        builder.ListFormat.ListLevelNumber = 2;
                        //builder.InsertStyleSeparator();
                        //builder.ListFormat.List = newList;

                        foreach (var t in grouped)
                        {
                            switch (t.Element("answerType").Value)
                            {
                                case "Yes/No":
                                    builder.ParagraphFormat.StyleName = "qHeading 4";
                                    builder.ListFormat.List = newList;
                                    builder.ListFormat.ListLevelNumber = 1;
                                    builder.Writeln(t.Element("wording").Value);
                                    builder.ListFormat.RemoveNumbers();
                                    //builder.ListFormat.List = null;
                                    builder.Writeln(t.Element("content").Value);
                                    break;
                                case "Upload":
                                    builder.InsertHtml(style + t.Element("wording").Value);
                                    builder.Writeln();
                                    builder.InsertHtml(style + t.Element("binaryData").Attribute("name").Value);
                                    //builder.InsertHtml(style + "<p>Upload</p>");
                                    builder.Writeln();
                                    break;
                                case "Text":
                                    builder.ParagraphFormat.StyleName = "qHeading 4";
                                    builder.ListFormat.List = newList;
                                    builder.ListFormat.ListLevelNumber = 1;
                                    builder.Writeln(t.Element("wording").Value);
                                    builder.ListFormat.RemoveNumbers();
                                    //builder.ListFormat.List = null;
                                    builder.Writeln();
                                    builder.ParagraphFormat.StyleName = "qNormal";
                                    builder.InsertHtml(style + t.Element("content").Value);
                                    builder.Writeln();
                                    break;
                                default:
                                    break;
                            }
                        }
                        break;
                    case "Table":
                        builder.ParagraphFormat.StyleName = isConnectedQuestion ? "qHeading 4" : "qHeading 2";
                        builder.ListFormat.List = list;
                        builder.Writeln(e.Element("wording").Value);
                        builder.ListFormat.RemoveNumbers();
                        //builder.ListFormat.List = null;
                        builder.ParagraphFormat.StyleName = "qDescription";
                        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level4;
                        builder.Writeln("\t" + e.Element("description").Value);
                        builder.ParagraphFormat.StyleName = "qNormal";

                        if (e.Element("answer").Element("content").Value == "")
                        {
                            if (e.Element("answer").Element("binaryData").Attribute("name").Value == "")
                            {
                                if (qSubtypeType == "Standard Table")
                                {
                                    //cols(header) -> "//self::rQuestion[@id=119][@answerType='Table']/answer/answerDataExtended[@order=1]/wording"
                                    //rows(content) -> "//self::rQuestion[@id=119][@answerType='Table']/answer/answerDataExtended[@order=1]/content"
                                    var qc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table']/answer/answerDataExtended[@order=1]/wording");
                                    var count = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Standard Table']/answer/answerDataExtended/content");

                                    builder.StartTable();
                                    var cols = 0;
                                    var rows = 0;
                                    foreach (var ee in qc)
                                    {
                                        cols += 1;
                                        builder.InsertCell();
                                        // Some special features for the header row.
                                        builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                        builder.Write(ee.Value);
                                    }
                                    builder.EndRow();
                                    rows = count.Count() / cols;
                                    for (int i = 1; i <= rows; i++)
                                    {
                                        var qr = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table']/answer/answerDataExtended[@order=" + i + "]/content");
                                        foreach (var ee in qr)
                                        {
                                            // Set features for the other rows and cells.
                                            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                            builder.CellFormat.Width = 100.0;
                                            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

                                            // Reset height and define a different height rule for table body
                                            builder.RowFormat.Height = 30.0;
                                            builder.RowFormat.HeightRule = HeightRule.Auto;

                                            builder.InsertCell();
                                            builder.Write(ee.Value);
                                        }
                                        builder.EndRow();
                                    }
                                    builder.EndTable();
                                }
                                if (qSubtypeType == "Fund as Columns")
                                {
                                    //cols ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended/gName"
                                    //rows ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended[@order=1]/wording"
                                    //content ->
                                    var qc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended/gName");
                                    var qq = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended[@order=1]/wording");
                                    var qr = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended/content");
                                    builder.StartTable();
                                    builder.InsertCell();
                                    builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);

                                    var i = 1;
                                    foreach (var ee in qc)
                                    {
                                        if (i % 2 != 0)
                                        {
                                            // Some special features for the header row.
                                            builder.InsertCell();
                                            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                            builder.Write(ee.Value);
                                        }
                                        i += 1;
                                    }
                                    builder.EndRow();
                                    var j = 1;
                                    foreach (var ee in qq)
                                    {
                                        builder.InsertCell();
                                        builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                        builder.Write(ee.Value);
                                        foreach (var eee in qr)
                                        {
                                            if (j % 2 != 0)
                                            {
                                                builder.InsertCell();
                                                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                                builder.Write(eee.Value);
                                            }
                                            j += 1;
                                        }
                                        builder.EndRow();
                                    }
                                    builder.EndTable();
                                }
                                if (qSubtypeType == "Fund as Rows")
                                {
                                    //cols ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended[@order=1]/wording"
                                    //rows ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Columns']/answer/answerDataExtended/gName"
                                    //content ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Rows']/answer/answerDataExtended[@order=1]/content"
                                    var qc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Rows']/answer/answerDataExtended[@order=1]/wording");
                                    var qq = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Rows']/answer/answerDataExtended/gName");
                                    builder.StartTable();
                                    builder.InsertCell();
                                    builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);

                                    foreach (var ee in qc)
                                    {
                                        // Some special features for the header row.
                                        builder.InsertCell();
                                        builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                        builder.Write(ee.Value);
                                    }
                                    builder.EndRow();

                                    var i = 1;
                                    var j = 1;
                                    foreach (var ee in qq)
                                    {
                                        if (i % 2 != 0)
                                        {
                                            builder.InsertCell();
                                            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                            builder.Write(ee.Value);
                                            var qr = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Rows']/answer/answerDataExtended[@order=" + j + "]/content");
                                            foreach (var eee in qr)
                                            {
                                                //if (j % 2 != 0)
                                                //{
                                                builder.InsertCell();
                                                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                                builder.Write(eee.Value);
                                                //}
                                                //j += 1;
                                            }
                                            builder.EndRow();
                                            j += 1;
                                        }
                                        i += 1;
                                    }
                                    builder.EndTable();
                                }
                                if (qSubtypeType == "Fund as Title")
                                {
                                    //cols ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended/gName"
                                    //rows ->"//self::rQuestion[@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended[@order=1]/wording"
                                    //content ->
                                    
                                    //get first FundId
                                    var cc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended[@order=1]/gId");
                                    var ic = cc.FirstOrDefault().Value;
                                    
                                    // get number of question(col) per FundId
                                    cc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended[@order=1][" + ic + "]/gName");
                                    

                                    var qq = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended[@order=1]/wording");
                                    var qr = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended/content");
                                    //builder.InsertCell();
                                    //builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);

                                    for (int j=0; j < cc.Count();j++)
                                    {
                                        var qc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Fund as Title']/answer/answerDataExtended[@order=1][" + j + "]/gName");


                                    }

                                    //var i = 1;
                                    //foreach (var ee in qc)
                                    //{
                                    //    builder.StartTable();
                                    //    if (i % 2 != 0)
                                    //    {
                                    //        // Some special features for the header row.
                                    //        builder.InsertCell();
                                    //        builder.CellFormat.HorizontalMerge = CellMerge.First;
                                    //        builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                    //        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                    //        builder.Write(ee.Value);
                                    //        for (int c = 0; c < cc.Count(); c++)
                                    //        {
                                    //            builder.InsertCell();
                                    //            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                                    //        }
                                    //        builder.EndRow();

                                    //    }
                                    //    i += 1;
                                    //    builder.EndTable();
                                    //}

                                    //var j = 1;
                                    //foreach (var e in qq)
                                    //{
                                    //    builder.InsertCell();
                                    //    builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                    //    builder.Write(e.Value);
                                    //    foreach (var ee in qr)
                                    //    {
                                    //        if (j % 2 != 0)
                                    //        {
                                    //            builder.InsertCell();
                                    //            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                    //            builder.Write(ee.Value);
                                    //        }
                                    //        j += 1;
                                    //    }
                                    //    builder.EndRow();
                                    //    builder.EndTable();
                                    //}

                                }
                                if (qSubtypeType == "Ranking Table")
                                {
                                    builder.Writeln("Rnk");
                                    // Gets all the FundId
                                    var ad = e.GetElementsUsingXPath("//self::rQuestion[@answerType='Table'][@answerSubType='Ranking Table']/answer/answerDataExtended[@order=1]/gId");

                                    // get only the first id
                                    var fId = ad.FirstOrDefault().Value;

                                    // get number of question(col) per FundId
                                    var cc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Ranking Table']/answer/answerDataExtended[@order=1][gId=" + fId + "]/content").Count();

                                    var qc = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Ranking Table']/answer/answerDataExtended[@order=1]/gName");
                                    var i = 1;
                                    foreach (var q in qc)
                                    {
                                        if (i % 2 != 0)
                                        {
                                            builder.StartTable();
                                            builder.InsertCell();
                                            builder.CellFormat.HorizontalMerge = CellMerge.First;
                                            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                            builder.Write(q.Value);

                                            for (int c = 0; c < cc; c++)
                                            {
                                                builder.InsertCell();
                                                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                                            }
                                            builder.EndRow();

                                            builder.InsertCell();
                                            builder.CellFormat.HorizontalMerge = CellMerge.First;
                                            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                            for (var j = 1; j <= cc; j++)
                                            {
                                                var ft = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Ranking Table']/answer/answerDataExtended[@order=1][" + j + "]/wording");
                                                // Some special features for the header row.
                                                //builder.InsertCell();
                                                //builder.CellFormat.HorizontalMerge = CellMerge.First;
                                                //builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                                //builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                //Console.WriteLine();
                                                builder.InsertCell();
                                                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                                builder.Write(ft.FirstOrDefault().Value);
                                            }
                                            builder.EndRow();


                                            //Questions (rows)
                                            builder.InsertCell();
                                            builder.CellFormat.HorizontalMerge = CellMerge.First;
                                            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                            builder.Write("Rank {1}");
                                            for (var j = 1; j <= cc; j++)
                                            {
                                                var ft = e.GetElementsUsingXPath("//self::rQuestion[@id=" + qId + "][@answerType='Table'][@answerSubType='Ranking Table']/answer/answerDataExtended[@order=1][" + j + "]/content");
                                                // Some special features for the header row.
                                                //Console.WriteLine();
                                                builder.InsertCell();
                                                builder.Write(ft.FirstOrDefault().Value);

                                            }
                                            builder.EndRow();
                                            builder.EndTable();
                                            builder.Writeln();
                                        }
                                        i += 1;
                                    }
                                    builder.InsertHtml(style + "<p>Ranking Table</p>");
                                    builder.InsertHtml(style + "<p>Not yet Implemented</p>");
                                }
                            }
                            else
                            {
                                builder.InsertHtml(style + e.Element("answer").Element("binaryData").Attribute("name").Value);
                                builder.Writeln();
                            }
                        }
                        else
                        {
                            builder.InsertHtml(style + e.Element("answer").Element("content").Value);
                        }
                        break;

                    default:
                        break;
                }
            }
        }
    }
}
