using Aspose.Words;
using System;
using System.Collections.Generic;
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
            license.SetLicense(@"Aspose.Words.lic");
        }

    }
}
