using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using com.truewindglobal.aspose.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class HeaderAndFooterBuilder
    {
        //public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        public static void Build(DocumentBuilder builder)
        {
            builder.MoveToDocumentStart();
            PageSetup pageSetup = builder.CurrentSection.PageSetup;

            /// <summary>
            /// Specify if we want headers/footers of the first page to be different from other pages
            /// You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
            /// different headers/footers for odd and even pages
            /// </summary>
            pageSetup.DifferentFirstPageHeaderFooter = true;

#if false
            // --- Create header for the first page ---
            //pageSetup.HeaderDistance = 20;
            //builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            //builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Set font properties for header text
            //builder.Font.Name = "Arial";
            //builder.Font.Bold = true;
            //builder.Font.Size = 14;
            // Specify header title for the first page
            //builder.Write("Aspose.Words Header/Footer Creation Primer - Title Page.");
#endif                
            // --- Create header for pages other than first ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert absolutely positioned image into the top/left corner of the header
            // Distance from the top/left edges of the page is set to 10 points
            string imageFileName = GlobalProperties.logo;
            builder.InsertImage(imageFileName, 
                                RelativeHorizontalPosition.Margin, 
                                0, 
                                RelativeVerticalPosition.Page, 
                                20, 
                                71, 
                                35, 
                                WrapType.Through);

            // --- Create footer for pages other than first ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.InsertField("PAGE", "");

#if false
            // We use table with two cells to make one part of the text on the line (with page numbering)
            // to be aligned left, and the other part of the text (with copyright) to be aligned right
            builder.StartTable();

            // Clear table borders
            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            // Set first cell to 1/3 of the page width
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100.0F / 3);

            // Insert page numbering text here
            // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            // Align this text to the left
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.InsertCell();
            // Set the second cell to 2/3 of the page width
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100.0F * 2 / 3);

            builder.Write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

            // Align this text to the right
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();
#endif
            //Save the resulting document
            builder.Document.Save(@"C:\AsposeWordsDemo\HeaderFooter.Primer.docx");


        }

        /// <summary>
        /// Clones and copies headers/footers form the previous section to the specified section.
        /// </summary>
        private static void CopyHeadersFootersFromPreviousSection(Section section)
        {
            Section previousSection = (Section)section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
                section.HeadersFooters.Add(headerFooter.Clone(true));
        }

    }
}
