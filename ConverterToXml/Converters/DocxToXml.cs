using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConverterToXml.Converters
{
    public class DocxToXml : IConvertable
    {
        // Id текущего списка
        int CurrentListID = 0;
        // Проверка каретки(в списке или нет): для проставления закрывающегося тега
        bool InList = false;

        /// <summary>
        /// Расстановка простых параграфов
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="p"></param>
        private void SimpleParagraph(StringBuilder sb, Paragraph p)
        {
            sb.Append($"<p>{p.InnerText}</p>");
        }

        /// <summary>
        /// Обработка элементов списка
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="p"></param>
        private void ListParagraph(StringBuilder sb, Paragraph p)
        {
            // уровень списка
            var level = p.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingLevelReference>().Val;
            // id списка
            var id = p.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingId>().Val;
            sb.Append($"<ul id=\"{id}\" level=\"{level}\"><p>{p.InnerText}</p></ul id=\"{id}\" level=\"{level}\">");
        }

        /// <summary>
        /// Обработка таблицы
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="table"></param>
        private void Table(StringBuilder sb, Table table)
        {
            sb.Append("<table>");
            foreach (var row in table.Elements<TableRow>())
            {
                sb.Append("<row>");
                foreach (var cell in row.Elements<TableCell>())
                {
                    sb.Append($"<cell>{cell.InnerText}</cell>");
                }
                sb.Append("</row>");
            }
            sb.Append("</table>");

        }

        /// <summary>
        /// Создание словаря, ключи которого будут соответствовать id списка в xml документе,
        /// а значения - id списка в исходном документе
        /// </summary>
        /// <param name="listEl">Словарь для списков</param>
        /// <param name="docBody">Тело документа</param>
        private void CreateDictList(Dictionary<int, string> listEl, Body docBody)
        {
            foreach(var el in docBody.ChildElements)
            {
                if(el.GetFirstChild<ParagraphProperties>() != null)
                {
                    if (el.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>() == null)
                    {
                        continue;
                    }
                    int key = el.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingId>().Val;
                    listEl[key] = ((DocumentFormat.OpenXml.Wordprocessing.Paragraph)el).ParagraphId.Value;
                }
            }
        }

        public string Convert(Stream memStream)
        {
            Dictionary<int, string> listEl = new Dictionary<int, string>();

            string xml = string.Empty;
            memStream.Position = 0;
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memStream, false))
            {
                StringBuilder sb = new StringBuilder(1000); // врядли в xml будет меньше 1000 символов
                sb.Append("<?xml version=\"1.0\"?><documents><document>");
                Body docBody = doc.MainDocumentPart.Document.Body; // тело документа (размеченный текст без стилей)
                CreateDictList(listEl, docBody);
                foreach (var element in docBody.ChildElements)
                {
                    string type = element.GetType().ToString();
                    try
                    {
                        switch (type)
                        {
                            case "DocumentFormat.OpenXml.Wordprocessing.Paragraph":

                                if (element.GetFirstChild<ParagraphProperties>() != null && element.GetFirstChild<ParagraphProperties>()
                                    .GetFirstChild<NumberingProperties>() != null) // список / не список
                                {
                                    if (element.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingId>().Val != CurrentListID)
                                    {
                                        CurrentListID = element.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingId>().Val;
                                        sb.Append($"<li id=\"{CurrentListID}\">");
                                        InList = true;
                                        ListParagraph(sb, (Paragraph)element);
                                    }
                                    else // текущий список
                                    {
                                        ListParagraph(sb, (Paragraph)element);
                                    }
                                    if (listEl.ContainsValue(((Paragraph)element).ParagraphId.Value))
                                    {
                                        sb.Append($"</li id=\"{element.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingId>().Val}\">");
                                    }
                                    continue;
                                }
                                else // не список
                                {
                                    SimpleParagraph(sb, (Paragraph)element);
                                    continue;
                                }
                            case "DocumentFormat.OpenXml.Wordprocessing.Table":

                                Table(sb, (Table)element);
                                continue;
                        }
                    }
                    catch (Exception e) // В случае наличия в документе тегов отличных от нужных, они будут проигнорированы
                    {
                        continue;
                    }
                }
                sb.Append(@"</document></documents>");
                xml = sb.ToString();
            }
            return xml;
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
    }
}
