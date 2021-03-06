﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConverterToXml.Converters
{
    public class DocxFromDocToXml
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

        public string Convert(Stream memStream, string charset = "UTF-8")
        {
            string xml = string.Empty;
            memStream.Position = 0;
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memStream, false))
            {
                StringBuilder sb = new StringBuilder(1000); // врядли в xml будет меньше 1000 символов
                sb.Append("<?xml version=\"1.0\"?><documents><document>");
                Body docBody = doc.MainDocumentPart.Document.Body; // тело документа (размеченный текст без стилей)
                foreach (var element in docBody.ChildElements)
                {
                    string type = element.GetType().ToString();
                    try
                    {
                        // такая проверка необходима, когда список является последним видимым элементом в документе, 
                        // но после него содержаться еще какие-нибудь служебные теги, 
                        // которые не содержат полезный текст
                        if (InList)
                        {
                            sb.Append($"</li id=\"{CurrentListID}\">");
                            InList = false;
                        }
                        switch (type)
                        {
                            case "DocumentFormat.OpenXml.Wordprocessing.Paragraph":
                            case "w:p":
                                if (element.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>() != null) // список / не список
                                {
                                    if (element.GetFirstChild<ParagraphProperties>().GetFirstChild<NumberingProperties>().GetFirstChild<NumberingId>().Val != CurrentListID) // новый список
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
                                    continue;
                                }
                                else // не список
                                {
                                    if (InList == true)
                                    {
                                        sb.Append($"</li id=\"{CurrentListID}\">");
                                        InList = false;
                                    }
                                    SimpleParagraph(sb, (Paragraph)element);
                                    continue;
                                }
                            case "DocumentFormat.OpenXml.Wordprocessing.Table":
                            case "w:tbl":
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
    }
}
