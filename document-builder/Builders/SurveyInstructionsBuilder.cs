using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.truewindglobal.aspose.Builders
{
    public class SurveyInstructionsBuilder
    {
        public static void Build(DocumentBuilder builder, IEnumerable<XElement> query)
        {
            foreach (var ele in query)
            {
                Console.WriteLine("INSTRUCTIONS");
                Console.WriteLine(ele);
            }
        }
    }
}
