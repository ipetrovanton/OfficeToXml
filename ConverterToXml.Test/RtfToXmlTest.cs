using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ConverterToXml.Converters;
using Xunit;

namespace ConverterToXml.Test
{
    public class RtfToXmlTest
    {
        [Fact]
        public void RtfConverterTest_NotNull()
        {
            RtfToXml converter = new RtfToXml();
            string curDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/rtf.rtf";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
