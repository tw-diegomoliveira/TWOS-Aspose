using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace com.truewindglobal.aspose
{
    public class HelperMethods
    {
        public static void ApplyLicence()
        {
            License license = new License();
            license.SetLicense(@"Aspose.Total.NET.lic");
        }

        public static void CopyStyles(Document srcDoc, Document destDoc)
        {
            // Normal
            Style srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Normal"]);
            srcStyle.Name = "qNormal";
            destDoc.Styles.AddCopy(srcStyle);

            // Title
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Title"]);
            srcStyle.Name = "qTitle";
            destDoc.Styles.AddCopy(srcStyle);

            // SubTitle
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Subtitle"]);
            srcStyle.Name = "qSubtitle";
            destDoc.Styles.AddCopy(srcStyle);

            // Adviser (SubTitle2)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Quote"]);
            srcStyle.Name = "qQuote";
            destDoc.Styles.AddCopy(srcStyle);

            // Instructions
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Instruction"]);
            srcStyle.Name = "qInstruction";
            destDoc.Styles.AddCopy(srcStyle);

            // Section (Heading 1)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Heading 1"]);
            srcStyle.Name = "qHeading 1";
            destDoc.Styles.AddCopy(srcStyle);

            // Filter Question (Heading 2)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Heading 2"]);
            srcStyle.Name = "qHeading 2";
            //srcStyle.ParagraphFormat.OutlineLevel = OutlineLevel.Level2; 
            destDoc.Styles.AddCopy(srcStyle);

            // (Heading 3)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Heading 3"]);
            srcStyle.Name = "qHeading 3";
            //srcStyle.ParagraphFormat.OutlineLevel = OutlineLevel.Level3;
            destDoc.Styles.AddCopy(srcStyle);

            // Connected Question (Heading 4)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Heading 4"]);
            srcStyle.Name = "qHeading 4";
            destDoc.Styles.AddCopy(srcStyle);

            // Filter Question (Yes)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["FilterYes"]);
            srcStyle.Name = "qFilterYes";
            destDoc.Styles.AddCopy(srcStyle);

            // Filter Question (No)
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["FilterNo"]);
            srcStyle.Name = "qFilterNo";
            destDoc.Styles.AddCopy(srcStyle);

            // Description
            srcStyle = srcDoc.Styles.AddCopy(srcDoc.Styles["Description"]);
            srcStyle.Name = "qDescription";
            destDoc.Styles.AddCopy(srcStyle);
        }

    }
}
