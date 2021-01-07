using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ConverterToXml.Converters
{
    public class OdsToXml: IConvertable
    {
        public string Convert(Stream stream)
        {
            var content = Unzip(stream);
            using (Stream memoryStream = new MemoryStream(content))
            {
                return ClearXml(memoryStream);
            }
        }

        public string ConvertByFile(string path)
        {
            using (FileStream fs = File.OpenRead(path))
            {
                return Convert(fs);
            }
        }

        /// <summary>
        /// Достать xml из потока ods
        /// </summary>
        /// <param name="stream">Обрабатываемый поток</param>
        /// <returns></returns>
        private byte[] Unzip(Stream stream)
        {
            // Необходимо копировать входящий поток в MemoryStream, т.к. байтовый массив нельзя вернуть из Stream
            using (MemoryStream ms = new MemoryStream())
            {
                ZipArchive archive = new ZipArchive(stream);
                var unzippedEntryStream = archive.GetEntry("content.xml")?.Open();
                unzippedEntryStream?.CopyTo(ms);
                return ms.ToArray();
            }
        }
        /// <summary>
        /// Обработка xml
        /// </summary>
        /// <param name="xmlStream"></param>
        /// <param name="charset"></param>
        /// <returns></returns>
        private string ClearXml(Stream xmlStream, string charset = "UTF-8")
        {
            StringBuilder sb = new StringBuilder(1000); // Врядли будет меньше 1k символов
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xmlStream);
            var spreadsheet = xDoc.GetElementsByTagName("office:spreadsheet")[0];
            XmlNodeList lists = spreadsheet.ChildNodes;
            sb.Append($"<?xml version=\"1.0\" encoding=\"{charset}\" standalone=\"yes\"?>");
            sb.Append("<documents><document>");
            foreach (XmlElement list in lists)
            {
                ClearLists(list, sb);
            }
            sb.Append("</document></documents>");
            return sb.ToString();
        }

        /// <summary>
        /// Обработка листов (таблиц)
        /// </summary>
        /// <param name="list"></param>
        /// <param name="sb"></param>
        private void ClearLists(XmlElement list, StringBuilder sb)
        {
            if (list.LocalName != "table")
                return;
            sb.Append($"<list name=\"{list.GetAttribute("table:name")}\">");
            foreach (XmlElement row in list.ChildNodes)
            {
                ClearRow(row, sb);
            }
            sb.Append($"</list name=\"{list.GetAttribute("table:name")}\">");
        }

        /// <summary>
        /// Обработка строк
        /// </summary>
        /// <param name="row"></param>
        /// <param name="sb"></param>
        private void ClearRow(XmlElement row, StringBuilder sb)
        {
            if (row.LocalName != "table-row")
                return;
            sb.Append(@"<row>");
            foreach (XmlElement cell in row.ChildNodes)
            {
                ClearCell(cell, sb);
            }
            sb.Append(@"</row>");
        }

        /// <summary>
        /// Обработка ячеек
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="sb"></param>
        private void ClearCell(XmlElement cell, StringBuilder sb)
        {
            if (cell.LocalName != "table-cell" || !cell.HasChildNodes)
                return;
            sb.Append($"<cell>{cell.InnerText}</cell>");
        }
    }
}
