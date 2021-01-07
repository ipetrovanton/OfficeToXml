using System.IO;
using ConverterToXml.Converters;
using Xunit;

namespace ConverterToXml.Test
{
    public class OdsToXmlTest
    {
        [Fact]
        public void OdsToXmlTest_NotNull()
        {
            OdsToXml converter = new OdsToXml();
            string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/ods.ods";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);
                Assert.NotNull(result);
            }
        }
    }
}
