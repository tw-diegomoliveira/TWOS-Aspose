using Aspose.Html;
using Aspose.Html.Dom;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace html_ngin
{
    class Program
    {
        static void Main()
        {
            var code = @"<table border=""1"" cellpadding=""1"" cellspacing=""1"">
                          <thead>
                            <tr>
                              <th scope=""col"">hardware</th>
                              <th scope=""col"">software</th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td>brand new server</td>
                              <td>all security patchs applyed</td>
                            </tr>
                            <tr>
                              <td>upgrade VPN gateway</td>
                              <td>&nbsp;</td>
                            </tr>
                          </tbody>
                        </table>
                        <p>&nbsp;</p>";

            // Initialize a document based on the prepared code
            using (var document = new HTMLDocument(code, "."))
            {
                // Here we evaluate the XPath expression where we select all child SPAN elements from elements whose 'class' attribute equals to 'happy':
                var result = document.Evaluate("//table//th",
                    document,
                    null,
                    Aspose.Html.Dom.XPath.XPathResultType.Any,
                    null);

                // Iterate over the resulted nodes
                for (Node node; (node = result.IterateNext()) != null;)
                {
                    Console.WriteLine(node.TextContent);
                    // output: Hello
                    // output: World!
                }
            }
            Console.ReadLine();
        }
    }
}
