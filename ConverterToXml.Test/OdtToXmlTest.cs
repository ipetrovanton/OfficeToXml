using System;
using System.Collections.Generic;
using System.IO;
using ConverterToXML;
using Xunit;

namespace ConverterToXml.Test
{
    public class OdtToXmlTest
    {
        [Fact]
        public void XlsxConverterTest_NotNull()
        {
            OdtToXml converter = new OdtToXml();
            string curDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/odt1.odt";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
