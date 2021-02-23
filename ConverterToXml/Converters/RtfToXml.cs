using HtmlAgilityPack;
using RtfPipe;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConverterToXml.Converters
{
    public class RtfToXml : IConvertable
    {
        public string Convert(Stream stream)
        {

            stream.Position = 0;
            string rtf = string.Empty;
            using (StreamReader sr = new StreamReader(stream))
            {
                rtf = sr.ReadToEnd();
            }

            // Эта строчка необходима для работы RtfPipe в Core (о чем сказано в документации)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            // С помощью либы RtfPipe создаем html
            var html = Rtf.ToHtml(rtf);
            // очищаем html от лишних тегов и атрибутов, возвращем готовую xml
            return ClearHtml(html);
        }

        public string ConvertByFile(string path)
        {
            if (!Path.IsPathFullyQualified(path))
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
            }
            using (FileStream fs = File.OpenRead(path))
            {
                return Convert(fs);
            }
        }


        private string ClearHtml(string html)
        {
            // Разбираем html с помощью HtmlAgilityPack
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            // Формируем коллекцию элеметов (тегов), у которых есть атрибут style
            var elementsWithStyleAttribute = doc.DocumentNode.SelectNodes("//@style");
            // формируем коллекцию тегов, которые не нужны 
            var excessNodes = doc.DocumentNode.SelectNodes("//b|//u|//strong|//br");
            // удаляем лишние элементы в html
            foreach (var element in excessNodes)
            {
                element.ParentNode.InnerHtml = element.InnerText;
                element.Remove();
            }
            foreach (var element in elementsWithStyleAttribute)
            {
                element.Attributes["style"].Remove();
            }
            // сохраняем изменения в html
            using (StringWriter writer = new StringWriter())
            {
                doc.Save(writer);
                html = writer.ToString();
            }
            // пишем отфильтрованный html в xml
            StringBuilder xml = new StringBuilder();
            xml.Append("<?xml version=\"1.0\"?><documents><document>");
            xml.Append(html);
            xml.Append("</documents></document>");
            return xml.ToString();
        }
    }
}
