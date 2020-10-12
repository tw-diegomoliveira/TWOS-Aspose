using System;
using System.Xml.Linq;

namespace com.truewindglobal.aspose
{
    class DataExplorer
    {
        static void Main()
        {
            //HelperMethods.ListAllTablesinDataSet(HelperMethods.LoadXMLDataIntoDataSet());
            //HelperMethods.LoadQuestionnaireDocumentData();
            //HelperMethods.ListTableContents(HelperMethods.LoadXMLDataIntoDataSet().Tables["rQuestion88"]);
            //HelperMethods.ListAllXMLFile();

            //LoadWithReader();
            //LoadWithXPath();
            //LoadWithXmlSerializer();

            var doc = XElement.Load("dummy.xml");
            //var query = from ele in doc.Elements(XName.Get("section"))
            //            let p = ele.Parent
            //            where p.Name == XName.Get("document")
            //            && p.Parent == null
            //            select ele;

            //foreach (var ele in query)
            //{
            //    Console.WriteLine(ele);
            //}
            //doc = XElement.Load("accountsNS.xml");
            var query = doc.GetElementsUsingXPath("//self::document/section[@id=28]/fQuestion[@id=42]/rQuestion");

            foreach (var ele in query)
            {
                var q = ele.GetElementsUsingXPath("//rQuestion[@id=119]//answer//answerDataExtended");
                Console.WriteLine(ele);
                foreach (var e in q)
                {
                    Console.WriteLine("\t\t" + e);
                }
            }

            Console.ReadLine();

        }
    }
}
