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
        public void OdtConverterTest_NotNull()
        {
            OdtToXml converter = new OdtToXml();
            string path = @$"Files/odt1.odt";
            if (!Path.IsPathFullyQualified(path))
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
            }

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
