using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace com.truewindglobal.aspose
{
    static public class XpathExtensionClass
    {
        public static IEnumerable<XElement> GetElementsUsingXPath(this XNode node, string xpath, params string[] namespaceDeclarations)
        {
            var nav = node.CreateNavigator();
            var xnm = GetXmlNamespaceManager(nav, namespaceDeclarations);
            var q = from ele in GetNodesUsingXPath(nav, xpath, xnm) select (XElement)ele;

            return q;
        }

        public static object GetEvaluationUsingXPath(this XNode node, string xpath, params string[] namespaceDeclarations)
        {
            var nav = node.CreateNavigator();
            var xnm = GetXmlNamespaceManager(nav, namespaceDeclarations);

            return nav.Evaluate(xpath, xnm);
        }

        public static IEnumerable<XNode> GetNodesUsingXPath(XPathNavigator nav, string xpath, IXmlNamespaceResolver nm)
        {
            var itr = nav.Select(xpath, nm);
            while (itr.MoveNext())
            {
                var uo = itr.Current.UnderlyingObject;
                yield return uo as XNode;
            }
        }

        private static readonly Regex DeclarationsPAttern = new Regex(@"(?x:^xmlns:(?<prefix>[^:\s]+)=(?<namespace>[^\s]+)$|
^(?<prefix>[^:\s]+):(?<namespace>[^\s]+)$)$", RegexOptions.Compiled);
        private static XmlNamespaceManager GetXmlNamespaceManager(XPathNavigator nav,
            string[] namespaceDeclarations)
        {
            var xnm = new XmlNamespaceManager(nav.NameTable);
            if (namespaceDeclarations != null)
            {
                foreach (string namespaceDeclaration in namespaceDeclarations)
                {
                    var match = DeclarationsPAttern.Match(namespaceDeclaration);
                    if (!match.Success)
                    {
                        throw new ArgumentException("incorrect namespace declaration", "namespaceDeclaration");
                    }
                    xnm.AddNamespace(match.Groups["prefix"].Value, match.Groups["namespace"].Value);
                }
            }
            if (string.IsNullOrEmpty(xnm.LookupNamespace("xhtml")))
            {
                xnm.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");
            }

            return xnm;
        }


    }
}
