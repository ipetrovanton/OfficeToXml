using ConverterToXml.Converters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;

namespace ConverterToXml.Test
{
    public class XlsToXlsxTest
    {
        [Fact]
        public void XlsxConverterTest_NotNull()
        {
            XlsToXlsx converter = new XlsToXlsx();
            string curDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = curDir + @"/Files/xls.xls";
            converter.ConvertToXlsxFile(path, curDir + @"/Files/Result.xlsx");
            Assert.True(File.Exists(curDir + @"/Files/Result.xlsx"));
        }
    }
}
