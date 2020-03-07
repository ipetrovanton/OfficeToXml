using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConverterToXml.Converters
{
    public class XlsToXml : IConvertable
    {
        public string Convert(Stream stream)
        {
            XlsToXlsx excelConvert = new XlsToXlsx();
            Stream str = excelConvert.Convert(stream);
            XlsToXml converter = new XlsToXml();
            return converter.Convert(str);
        }

        public string ConvertByFile(string path)
        {
            using (FileStream fs = File.OpenRead(path))
            {
                return Convert(fs);
            }
        }
    }
}
