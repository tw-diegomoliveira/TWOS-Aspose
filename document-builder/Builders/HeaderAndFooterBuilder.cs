using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using com.truewindglobal.aspose.Models;
using document_builder.Properties;
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

            // --- Create header for pages other than first ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert absolutely positioned image into the top/left corner of the header
            // Distance from the top/left edges of the page is set to 10 points
            Shape logo = builder.InsertImage(Resources.logo, width: 71, height: 35);
            logo.WrapType = WrapType.Through;
            logo.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
            logo.HorizontalAlignment = HorizontalAlignment.Center;
            logo.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

            // --- Create footer for pages other than first ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.InsertField("PAGE", "");
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
