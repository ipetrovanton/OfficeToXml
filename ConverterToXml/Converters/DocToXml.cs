using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConverterToXml.Converters
{
    public class DocToXml: IConvertable
    {
        public string Convert(Stream stream)
        {
            DocToDocx docToDocx = new DocToDocx();
            MemoryStream ms = docToDocx.ConvertFromStreamToDocxMemoryStream(stream);
            DocxFromDocToXml docxToXml = new DocxFromDocToXml();
            string xml = docxToXml.Convert(ms);
            return xml;
        }

        public string ConvertByFile(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                DocToDocx docToDocx = new DocToDocx();
                MemoryStream ms = docToDocx.ConvertFromStreamToDocxMemoryStream(fs);
                DocxFromDocToXml docxToXml = new DocxFromDocToXml();
                string xml = docxToXml.Convert(ms);
                return xml;
            }
        }
    }
}
