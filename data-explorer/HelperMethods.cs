using com.truewindglobal.aspose.Models;
using System;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace com.truewindglobal.aspose
{
    class HelperMethods
    {
        private static void LoadWithXmlSerializer()
        {
            using (var fileStream = File.Open(GlobalProperties.SerializerTestData, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(MyDocument));
                var myDocument = (MyDocument)serializer.Deserialize(fileStream);

                Console.WriteLine($"My Property : {myDocument.MyProperty}");
                Console.WriteLine($"My Attribute : {myDocument.MyAttributeProperty.Value}");

                foreach (var item in myDocument.MyList)
                {
                    Console.WriteLine(item);
                }
            }
        }

        private static void LoadWithXPath()
        {
            using (var fileStream = File.Open("test.xml", FileMode.Open))
            {
                //Load the file and create a navigator object. 
                XPathDocument xPath = new XPathDocument(fileStream);
                var navigator = xPath.CreateNavigator();

                //Compile the query with a namespace prefix. 
                XPathExpression query = navigator.Compile("ns:MyDocument/ns:MyProperty");

                //Do some BS to get the default namespace to actually be called ns. 
                var nameSpace = new XmlNamespaceManager(navigator.NameTable);
                nameSpace.AddNamespace("ns", "http://www.dotnetcoretutorials.com/namespace");
                query.SetContext(nameSpace);

                Console.WriteLine("My Property Value : " + navigator.SelectSingleNode(query).Value);
            }
        }

        private static void LoadWithReader()
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;

            using (var filestream = File.OpenText(GlobalProperties.Filename))
            using (var reader = XmlReader.Create(filestream, settings))
            {
                //// Move to the root element
                //reader.MoveToContent();

                //var query = reader
                //    // Read all child elements <Widget>
                //    .ReadElements("section88", "")
                //    // And extract the text content of their first child element <Description>
                //    .SelectMany(r => r.ReadElements("fQuestion88", "").Select(i => i.ReadElementContentAsString()).Take(1));
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            Console.WriteLine($"Start Element: {reader.Name}. Has Attributes? : {reader.HasAttributes}");
                            break;
                        case XmlNodeType.Text:
                            Console.WriteLine($"Inner Text: {reader.Value}");
                            break;
                        case XmlNodeType.EndElement:
                            Console.WriteLine($"End Element: {reader.Name}");
                            break;
                        default:
                            Console.WriteLine($"Unknown: {reader.NodeType}");
                            break;
                    }
                }
            }

        }

        public static void ListAllXMLFile()
        {
            //XmlDocument document = new XmlDocument();
            //XmlReader reader = XmlReader.Create(GlobalProperties.Filename);
            //document.Load(reader);

            XmlDocument document = new XmlDocument();
            document.Load(GlobalProperties.Filename);



            //from element in document.SelectSingleNode("document88").SelectSingleNode("sections").SelectNodes("section88")
            //where element.Att


            XmlNode node = document.SelectSingleNode("/document88/sections");
            XmlNodeList nodes = document.SelectNodes("/document88/sections/section88/filterQuestions/fQuestion88");

            //from element in node.SelectNodes("section88").GetEnumerator()
            //where element.



            DataSet ds = new DataSet();

            Console.WriteLine(node.OuterXml);
            Console.ReadLine();

            //Console.WriteLine(node.InnerText);

            foreach (XmlNode n in nodes)
            {
                //Console.Write(n.InnerText);
                Console.WriteLine(n.OuterXml);
                String xmlString = n.OuterXml;
                StringReader sr = new StringReader(xmlString);
                ds.ReadXml(sr);
                ListAllTablesinDataSet(ds);
                ListTableContents(ds.Tables["rQuestion88"]);
                ListTableContents(ds.Tables["answer"]);
            }
            //document.Save(Console.Out);
        }
        public static DataSet LoadXMLDataIntoDataSet()
        {
            DataSet ds = new DataSet();
            ds.ReadXml(GlobalProperties.Filename);
            return ds;
        }
        public static void ListAllTablesinDataSet(DataSet ds)
        {
            var i = 0;
            foreach (DataTable table in ds.Tables)
            {
                Console.WriteLine(i + "\t" + table.TableName);
                i += 1;
                foreach (DataColumn dc in table.Columns)
                {
                    Console.WriteLine("\t\t" + dc.ColumnName);
                }
            }
            Console.ReadLine();
        }

        public static void ListTableContents(DataTable table)
        {
            Console.WriteLine(table.TableName);

            foreach (DataColumn column in table.Columns)
            {
                Console.Write(column.ColumnName + "\t");
            }
            Console.WriteLine();

            foreach (DataRow row in table.Rows)
            {
                foreach (object item in row.ItemArray)
                {
                    Console.Write(item.ToString() + " | ");
                }
                Console.WriteLine();
            }
        }

        public static void LoadQuestionnaireDocumentData()
        {
            DataSet ds = new DataSet();
            ds.ReadXml(@"C:\AsposeWordsDemo\FSdata.xml");

            //String xml = "<root><answer>yes</answer></root>";
            //StringReader sr = new StringReader(xml);
            //ds.ReadXml(sr);

            Console.WriteLine(ds.Tables[2]);
            Console.ReadLine();

            DataTable tableQuestions = ds.Tables["SECTION88"];

            Console.Write("Table Name: " + tableQuestions.TableName);
            Console.WriteLine();

            foreach (DataColumn column in tableQuestions.Columns)
            {
                Console.Write(column.ColumnName + "\t");
            }
            Console.WriteLine();
            ////DataRow[] filterRowTableQuestion = tableQuestion.Select("Section_id = 1");

            foreach (DataRow row in tableQuestions.Rows)
            {
                foreach (object item in row.ItemArray)
                {
                    Console.Write(item.ToString() + " | ");
                }
                Console.WriteLine();

            }

            DataTable tableConnectedQuestions = ds.Tables["rQuestion"];

            Console.Write("Table Name: " + tableConnectedQuestions.TableName);
            Console.WriteLine();

            foreach (DataColumn column in tableConnectedQuestions.Columns)
            {
                Console.Write(column.ColumnName + "\t");
            }
            Console.WriteLine();
            ////DataRow[] filterRowTableQuestion = tableQuestion.Select("Section_id = 1");

            foreach (DataRow row in tableConnectedQuestions.Rows)
            {
                foreach (object item in row.ItemArray)
                {
                    Console.Write(item.ToString() + " | ");
                }
                Console.WriteLine();

            }

            Console.ReadLine();

            DataTable tableAnswers = ds.Tables["answerData"];

            Console.Write("Table Name: " + tableAnswers.TableName);
            Console.WriteLine();

            foreach (DataColumn column in tableAnswers.Columns)
            {
                Console.Write(column.ColumnName + "\t");
            }
            Console.WriteLine();
            ////DataRow[] filterRowTableQuestion = tableQuestion.Select("Section_id = 1");

            foreach (DataRow row in tableAnswers.Rows)
            {
                foreach (object item in row.ItemArray)
                {
                    Console.Write(item.ToString() + " | ");
                }
                Console.WriteLine();

            }

            Console.ReadLine();
        }
    }
}
