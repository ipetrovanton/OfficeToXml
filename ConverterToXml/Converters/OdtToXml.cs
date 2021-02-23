using System;
using System.Collections.Generic;
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
                var unzippedEntryStream = archive.GetEntry("content.xml").Open();
                unzippedEntryStream.CopyTo(ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Обработка xml
        /// </summary>
        /// <param name="xmlStream"></param>
        /// <returns>Result as <see cref="string"/>.</returns>
        private string ClearXml(Stream xmlStream)
        {
            // Создаем настройки XmlWriter
            XmlWriterSettings settings = new XmlWriterSettings();
            // Необходимый параметр для формирования вложенности тегов
            settings.ConformanceLevel = ConformanceLevel.Auto;
            // XmlWriter будем вести запись в StringBuilder
            StringBuilder sb = new StringBuilder();

            using (XmlWriter writer = XmlWriter.Create(sb, settings))
            {
                XmlReader reader = XmlReader.Create(xmlStream);
                reader.ReadToFollowing("office:body");
                while (reader.Read())
                {
                    MethodSwitcher(reader, writer);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Метод прокручивает xml до тега body.
        /// </summary>
        /// <param name="reader">Текущий <see cref="XmlReader"/>.</param>
        private void SkipExtra(XmlReader reader)
        {
            reader.ReadToFollowing("office:body");
        }

        /// <summary>
        /// Выбираем подходящий метод разбора текста в зависимости от <see cref="XmlNodeType"/>.
        /// </summary>
        /// <param name="reader"><see cref="XmlReader"/>.</param>
        /// <param name="writer"><see cref="XmlWriter"/>.</param>
        private void MethodSwitcher(XmlReader reader, XmlWriter writer)
        {
            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    if (!reader.IsEmptyElement || reader.Name == "text:s")
                    {
                        TagWriter(reader, writer);
                    }
                    break;
                case XmlNodeType.EndElement:
                    if (tags.Contains(reader.LocalName))
                    {
                        writer.WriteEndElement();
                        writer.Flush();
                    }
                    break;
                case XmlNodeType.Text:
                    writer.WriteString(reader.Value);
                    break;
                default: break;
            }
        }


        /// <summary>
        /// Метод открывает теги в <see cref="XmlWriter"/>.
        /// </summary>
        /// <param name="reader"><see cref="XmlReader"/>.</param>
        /// <param name="writer"><see cref="XmlWriter"/>.</param>
        private void TagWriter(XmlReader reader, XmlWriter writer)
        {
            switch (reader.LocalName)
            {
                case "p":
                    writer.WriteStartElement("p");
                    break;
                case "table":
                    writer.WriteStartElement("table");
                    break;
                case "table-row":
                    writer.WriteStartElement("row");
                    break;
                case "table-cell":
                    writer.WriteStartElement("cell");
                    break;
                case "list":
                    writer.WriteStartElement("list");
                    break;
                case "list-item":
                    writer.WriteStartElement("item");
                    break;
                case "s":
                    writer.WriteString(" ");
                    break;
                default: break;
            }
        }

        /// <summary>
        /// <see cref="XmlReader.LocalName"/> для корректного закрытия тегов.
        /// </summary>
        private readonly List<string> tags = new List<string>()
        {
            "p",
            "table",
            "table-row",
            "table-cell",
            "list",
            "list-item"
        };
    }
}
