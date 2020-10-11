using System.IO;
using Xunit;

namespace ConverterToXml.Test
{
    public class DocToDocxTest
    {
        [Fact]
        public void DocToDocxConvertToDocx_NotNull()
        {
            Converters.DocToDocx converter = new Converters.DocToDocx();
            string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"\Files\doc.doc";
            converter.ConvertToDocx(path, curDir + @"\Files\Result.docx");
            Assert.True(File.Exists(curDir + @"\Files\Result.docx"));
        }
    }
}
