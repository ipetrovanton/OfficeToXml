using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ConverterToXml.Converters;
using Xunit;

namespace ConverterToXml.Test
{
    public class XlsxToXmlTest
    {
        [Fact]
        public void XlsxConverterTest_NotNull()
        {
            XlsxToXml converter = new XlsxToXml();
            string curDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/xlsx.xlsx";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
