using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace document_analyser
{
    public class HelperMethods
    {
        public static void ApplyLicence()
        {
            License license = new License();
            license.SetLicense(@"Aspose.Total.NET.lic");
        }

        public static string GetAllDocumentProperties(DocumentBuilder builder)
        {
            Console.WriteLine("2. Built-in Properties");
            foreach (DocumentProperty prop in builder.Document.BuiltInDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);

            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty prop in builder.Document.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);

            CustomDocumentProperties props = builder.Document.CustomDocumentProperties;
            if (props["Authorized"] == null)
            {
                props.Add("Authorized", true);
                props.Add("Authorized By", "John Smith");
                props.Add("Authorized Date", DateTime.Today);
                props.Add("Authorized Revision", builder.Document.BuiltInDocumentProperties.RevisionNumber);
                props.Add("Authorized Amount", 123.45);
            }
            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty prop in builder.Document.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);
            builder.Document.CustomDocumentProperties.Remove("Authorized Date");
            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty prop in builder.Document.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);

            //Console.WriteLine(builder.Document.ToString());
            Console.WriteLine("4. Variables");
            string variables = "";
            foreach (KeyValuePair<string, string> entry in builder.Document.Variables)
            {
                string name = entry.Key.ToString();
                string value = entry.Value.ToString();
                if (variables == "")
                {
                    // Do something useful.
                    variables = "Name: " + name + "," + "Value: {1}" + value;
                }
                else
                {
                    variables = variables + "Name: " + name + "," + "Value: {1}" + value;
                }
            }
            PageSetup pageSetup = builder.PageSetup;
            //pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
            //pageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0);
            //pageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5);
            //pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5);
            //pageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2);
            //pageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2);
            Console.WriteLine(pageSetup.TopMargin);
            Console.WriteLine(pageSetup.BottomMargin);
            Console.WriteLine(pageSetup.LeftMargin);
            Console.WriteLine(pageSetup.RightMargin);
            Console.WriteLine(pageSetup.HeaderDistance);
            Console.WriteLine(pageSetup.FooterDistance);

            // Create and attach collector before the document before page layout is built.
            LayoutCollector layoutCollector = new LayoutCollector(builder.Document);

            // This will build layout model and collect necessary information.
            builder.Document.UpdatePageLayout();

            // Print the details of each document node including the page numbers. 
            foreach (Node node in builder.Document.Sections[1].Body.GetChildNodes(NodeType.Paragraph, true))
            {
                Console.WriteLine(" --------- ");
                Console.WriteLine("NodeType:   " + Node.NodeTypeToString(node.NodeType));
                Console.WriteLine("Text:       \"" + node.ToString(SaveFormat.Text).Trim() + "\"");
                Console.WriteLine("Page Start: " + layoutCollector.GetStartPageIndex(node));
                Console.WriteLine("Page End:   " + layoutCollector.GetEndPageIndex(node));
                Console.WriteLine(" --------- ");
                Console.WriteLine();
            }

            // Detatch the collector from the document.
            layoutCollector.Document = null;

            Console.WriteLine("\nFound the page numbers of all nodes successfully.");

            StyleCollection styles = builder.Document.Styles;
            // Iterate through all the styles.
            foreach (Style style in styles)
            {
                Console.WriteLine(style.Name);
            }
            
            return "";
        }
    }
}
