using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace document_analyser
{
    class Analyse
    {
        static void Main()
        {
            HelperMethods.ApplyLicence();
            Document doc = new Document(@"C:\AsposeWordsDemo\Styles_Questionnaire_v1.0.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //string properties = HelperMethods.GetAllDocumentProperties(builder);
            var i = 0;
            foreach (Shape img in doc.GetChildNodes(NodeType.Shape, true))
            {
                i += 1;
                Shape imported = (Shape)doc.ImportNode(img, true);
                Console.WriteLine(i);
                //doc.FirstSection.Body.FirstParagraph.AppendChild(imported);

            }
#if false
            MemoryStream outStream = new MemoryStream();
            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            PageRange pageRange = new PageRange(0, doc.PageCount - 1);
            imageSaveOptions.PageSet = new PageSet(pageRange);

            /// <summary>
            /// Saving all pages to disk as image
            /// </summary>
            imageSaveOptions.PageSavingCallback = new HandlePageSavingCallbackImg();

            /// <summary>
            /// Saving all pages as stream
            /// </summary>
            //HandlePageSavingCallback handler = new HandlePageSavingCallback();
            //imageSaveOptions.PageSavingCallback = handler;

            doc.Save(outStream, imageSaveOptions);
#endif
            //Console.WriteLine(outStream.ToArray());

            //Console.WriteLine(properties);
#if false
            using (IEnumerator<Style> stylesEnum = doc.Styles.GetEnumerator())
            {
                while (stylesEnum.MoveNext())
                {
                    Style curStyle = stylesEnum.Current;
                    Console.WriteLine($"Style name:\t\"{curStyle.Name}\", of type \"{curStyle.Type}\"");
                    Console.WriteLine($"\tSubsequent style:\t{curStyle.NextParagraphStyleName}");
                    Console.WriteLine($"\tIs heading:\t\t\t{curStyle.IsHeading}");
                    Console.WriteLine($"\tIs QuickStyle:\t\t{curStyle.IsQuickStyle}");
                }
            }
#endif
            Console.ReadLine();
        }

        /// <summary>
        /// Saving all pages as stream
        /// </summary>
        private class HandlePageSavingCallback : IPageSavingCallback
        {
            public ArrayList jpeg_streams = new ArrayList();
            public void PageSaving(PageSavingArgs args)
            {
                args.PageStream = new MemoryStream();
                args.KeepPageStreamOpen = true;
                jpeg_streams.Add(args.PageStream);
            }
        }
        /// <summary>
        /// Saving all pages to disk as image
        /// </summary>
        private class HandlePageSavingCallbackImg : IPageSavingCallback
        {
            public void PageSaving(PageSavingArgs args)
            {
                args.PageFileName = string.Format(@"C:\AsposeWordsDemo\Questionnaire_{0}.jpeg", args.PageIndex);
            }
        }
    }

}
