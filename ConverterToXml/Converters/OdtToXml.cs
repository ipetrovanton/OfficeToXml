using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using ConverterToXml.Converters;

namespace ConverterToXML
{
    public class OdtToXml : IConvertable
    {
        public string Convert(Stream memStream)
        {
            var content = Unzip(memStream);
            using (Stream stream = new MemoryStream(content))
            {
                return ClearXml(stream);
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
                var unzippedEntryStream = archive.GetEntry("content.xml").Open();
                unzippedEntryStream.CopyTo(ms);
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
            XmlNode text = xDoc.GetElementsByTagName("office:body")[0].FirstChild;
            sb.Append($"<?xml version=\"1.0\" encoding=\"{charset}\" standalone=\"yes\"?>");
            sb.Append("<documents><document>");
            // Основной метод для очистки разметки xml
            ClearText(text, sb);
            sb.Append("</document></documents>");
            return sb.ToString();
        }

        /// <summary>
        /// Выбираем подходящий метод разбора текста в зависимости от тега
        /// </summary>
        /// <param name="text"></param>
        /// <param name="sb"></param>
        private void ClearText(XmlNode text, StringBuilder sb)
        {
            foreach (XmlNode node in text.ChildNodes)
            {
                switch (node.Name)
                {
                    case "text:p":
                        ParagraphClear(node, sb);
                        continue;
                    case "table:table":
                        TableClear(node, sb);
                        continue;
                    case "text:list":
                        ListClear(node, sb);
                        continue;
                    default: continue;
                }
            }
        }

        /// <summary>
        /// В теге <text:p> иногда попадаются другие теги для мета-сведений:
        /// их мы игнорим, рекурсивно двигаясь до тега <p>
        /// </summary>
        /// <param name="p"></param>
        /// <param name="sb"></param>
        private void ParagraphClear(XmlNode p, StringBuilder sb)
        {
            if (!p.HasChildNodes)
                sb.Append($"<p>{p.InnerText}</p>");
            else
            {
                foreach (XmlNode node in p.ChildNodes)
                {
                    ParagraphClear(node, sb);
                }
            }
        }

        /// <summary>
        /// Обработка таблицы
        /// </summary>
        /// <param name="table"></param>
        /// <param name="sb"></param>
        private void TableClear(XmlNode table, StringBuilder sb)
        {
            sb.Append("<table>");
            foreach (XmlNode row in table.ChildNodes)
            {
                // не обрабатываем лишние теги
                if (row.Name != "table:table-row")
                    continue;
                RowClear(row, sb);
            }
            sb.Append("</table>");
        }

        /// <summary>
        /// Обработка строк. 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="sb"></param>
        private void RowClear(XmlNode row, StringBuilder sb)
        {
            sb.Append("<row>");
            foreach (XmlNode cell in row.ChildNodes)
            {
                // не обрабатываем лишние теги
                if (cell.Name != "table:table-cell")
                    continue;
                CellClear(cell, sb);
            }
            sb.Append("</row>");
        }

        /// <summary>
        /// Обработка ячеек таблицы
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="sb"></param>
        private void CellClear(XmlNode cell, StringBuilder sb)
        {
            sb.Append("<cell>");
            ParagraphClear(cell.FirstChild, sb);
            sb.Append("</cell>");
        }

        /// <summary>
        /// Обработка списка.
        /// </summary>
        /// <param name="list"></param>
        /// <param name="sb"></param>
        private void ListClear(XmlNode list, StringBuilder sb)
        {
            sb.Append("<list>");
            foreach (XmlNode xmlNode in list.ChildNodes)
            {
                // внутри одного списка может быть еще один список, поэтому рекурсия
                switch (xmlNode.Name)
                {
                    case "text:list":
                        ListClear(xmlNode, sb);
                        continue;
                }
                ListItemClear(xmlNode, sb);
            }
            sb.Append("</list>");
        }

        /// <summary>
        /// ОБработка элемента списка
        /// </summary>
        /// <param name="item"></param>
        /// <param name="sb"></param>
        private void ListItemClear(XmlNode item, StringBuilder sb)
        {
            sb.Append("<list-item>");
            foreach (XmlNode node in item.ChildNodes)
            {
                // внутри одного элемента списка может быть другой список или параграф
                switch (node.Name)
                {
                    case "text:p":
                        ParagraphClear(node, sb);
                        break;
                    case "text:list":
                        ListClear(node, sb);
                        break;
                }
            }
            sb.Append("</list-item>");
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
