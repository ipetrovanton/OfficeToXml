using ConverterToXml.Converters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;

namespace ConverterToXml.Test
{
    public class DocxToXmlTest
    {
        [Fact]
        public void XlsxConverterTest_NotNull()
        {
            DocxToXml converter = new DocxToXml();
            string curDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/Result.docx";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
