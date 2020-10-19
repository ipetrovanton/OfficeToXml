using System;
using System.IO;
using Xunit;

namespace ConverterToXml.Test
{
    
    public class DocToXmlTest
    {
        [Fact]
        public void DocToDocxConvertToDocx_NotNull()
        {
            Converters.DocToXml converter = new Converters.DocToXml();
            string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/doc1.doc";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
