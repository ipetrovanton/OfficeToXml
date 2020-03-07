using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using CellType = NPOI.SS.UserModel.CellType;

namespace ConverterToXml.Converters
{
    public class XlsToXlsx
    {
        /// <summary>
        /// Создание из потока xls файла xlsx по указанному пути
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="path"></param>
        public void ConvertToXlsxFile(MemoryStream stream, string path)
        {
            var result = Convert(stream);
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                fs.Write(result.ToArray());
            }
        }

        /// <summary>
        /// Создание из файла xls файла xlsx по указанному пути
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="path"></param>
        public void ConvertToXlsxFile(string xlsPath, string destPath)
        {

            MemoryStream result;
            using (FileStream fs = new FileStream(xlsPath, FileMode.Open))
            {
                result = Convert(fs);
            }
            using (FileStream fs = new FileStream(destPath, FileMode.OpenOrCreate))
            {
                fs.Write(result.ToArray());
            }
        }

        /// <summary>
        /// Метод инициализирует процесс конвертации
        /// </summary>
        /// <param name="sourceStream">Поток с xls</param>
        /// <returns>Массив байтов (читать в поток)</returns>
        public MemoryStream Convert(Stream sourceStream)
        {
            // Открытие xls
            var source = new HSSFWorkbook(sourceStream);
            // Создание объекта для будущего xlsx
            var destination = new XSSFWorkbook();
            // Копируем листы из xls и доабвляем в xlsx
            for (int i = 0; i < source.NumberOfSheets; i++)
            {
                var xssfSheet = (XSSFSheet)destination.CreateSheet(source.GetSheetAt(i).SheetName);
                var hssfSheet = (HSSFSheet)source.GetSheetAt(i);
                CopyStyles(hssfSheet, xssfSheet);
                CopySheet(hssfSheet, xssfSheet);
            }
            // Возвращаем сконвертированный результат
            using (var ms = new MemoryStream())
            {
                destination.Write(ms);
                return ms;
            }
        }

        private void CopyStyles(HSSFSheet from, XSSFSheet to)
        {
            for (short i = 0; i <= from.Workbook.NumberOfFonts; i++)
            {
                CopyFont(to.Workbook.CreateFont(), from.Workbook.GetFontAt(i));
            }

            for (short i = 0; i < from.Workbook.NumCellStyles; i++)
            {
                CopyStyle(to.Workbook.CreateCellStyle(), from.Workbook.GetCellStyleAt(i), to.Workbook, from.Workbook);
            }
        }

        private void CopyFont(IFont toFront, IFont fontFrom)
        {
            toFront.Boldweight = fontFrom.Boldweight;
            toFront.Charset = fontFrom.Charset;
            toFront.Color = fontFrom.Color;
            toFront.FontHeightInPoints = fontFrom.FontHeightInPoints;
            toFront.FontName = fontFrom.FontName;
            toFront.IsBold = fontFrom.IsBold;
            toFront.IsItalic = fontFrom.IsItalic;
            toFront.IsStrikeout = fontFrom.IsStrikeout;
        }

        private void CopyStyle(ICellStyle toCellStyle, ICellStyle fromCellStyle, IWorkbook toWorkbook, IWorkbook fromWorkbook)
        {
            toCellStyle.Alignment = fromCellStyle.Alignment;
            toCellStyle.BorderBottom = fromCellStyle.BorderBottom;
            toCellStyle.BorderDiagonal = fromCellStyle.BorderDiagonal;
            toCellStyle.BorderDiagonalColor = fromCellStyle.BorderDiagonalColor;
            toCellStyle.BorderDiagonalLineStyle = fromCellStyle.BorderDiagonalLineStyle;
            toCellStyle.BorderLeft = fromCellStyle.BorderLeft;
            toCellStyle.BorderRight = fromCellStyle.BorderRight;
            toCellStyle.BorderTop = fromCellStyle.BorderTop;
            toCellStyle.BottomBorderColor = fromCellStyle.BottomBorderColor;
            toCellStyle.DataFormat = fromCellStyle.DataFormat;
            toCellStyle.FillBackgroundColor = fromCellStyle.FillBackgroundColor;
            toCellStyle.FillForegroundColor = fromCellStyle.FillForegroundColor;
            toCellStyle.FillPattern = fromCellStyle.FillPattern;
            toCellStyle.Indention = fromCellStyle.Indention;
            toCellStyle.IsHidden = fromCellStyle.IsHidden;
            toCellStyle.IsLocked = fromCellStyle.IsLocked;
            toCellStyle.LeftBorderColor = fromCellStyle.LeftBorderColor;
            toCellStyle.RightBorderColor = fromCellStyle.RightBorderColor;
            toCellStyle.Rotation = fromCellStyle.Rotation;
            toCellStyle.ShrinkToFit = fromCellStyle.ShrinkToFit;
            toCellStyle.TopBorderColor = fromCellStyle.TopBorderColor;
            toCellStyle.VerticalAlignment = fromCellStyle.VerticalAlignment;
            toCellStyle.WrapText = fromCellStyle.WrapText;
            toCellStyle.SetFont(toWorkbook.GetFontAt((short)(fromCellStyle.GetFont(fromWorkbook).Index + 1)));
        }

        /// <summary>
        /// Копипрование содержимого листа
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        private void CopySheet(HSSFSheet source, XSSFSheet destination)
        {
            var maxColumnNum = 0;
            var mergedRegions = new List<CellRangeAddress>();
            for (int i = source.FirstRowNum; i <= source.LastRowNum; i++)
            {
                var srcRow = (HSSFRow)source.GetRow(i);
                var destRow = (XSSFRow)destination.CreateRow(i);
                if (srcRow != null)
                {
                    CopyRow(source, destination, srcRow, destRow, mergedRegions);
                    // поиск максимального номера ячейки в строке для копирования ширины столбцов
                    if (srcRow.LastCellNum > maxColumnNum)
                    {
                        maxColumnNum = srcRow.LastCellNum;
                    }
                }
            }
            // копируем ширину столбцов исходного документа
            for (int i = 0; i <= maxColumnNum; i++)
            {
                destination.SetColumnWidth(i, source.GetColumnWidth(i));
            }
        }

        /// <summary>
        /// Копирование содежимого ячеек
        /// </summary>
        /// <param name="srcSheet"></param>
        /// <param name="destSheet"></param>
        /// <param name="srcRow"></param>
        /// <param name="destRow"></param>
        /// <param name="mergedRegions"></param>
        private void CopyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow, List<CellRangeAddress> mergedRegions)
        {
            // Копирование высоты строки
            destRow.Height = srcRow.Height;

            for (int j = srcRow.FirstCellNum; srcRow.LastCellNum >= 0 && j <= srcRow.LastCellNum; j++)
            {
                var oldCell = (HSSFCell)srcRow.GetCell(j);
                var newCell = (XSSFCell)destRow.GetCell(j);
                if (oldCell != null)
                {
                    // создание новой ячейки в новой таблице
                    if (newCell == null)
                    {
                        newCell = (XSSFCell)destRow.CreateCell(j);
                    }

                    CopyCell(oldCell, newCell);
                    // Ниже идет обработка объединенных ячеек
                    // Проверка на вхождение текущей ячейки в число объединенных
                    var mergedRegion = GetMergedRegion(srcSheet, srcRow.RowNum,
                            (short)oldCell.ColumnIndex);
                    // Если ячейка является объединенное
                    if (mergedRegion != null)
                    {
                        // Проверяем обработывали ли мы уже группу объединенных ячеек или нет
                        var newMergedRegion = new CellRangeAddress(mergedRegion.FirstRow,
                                mergedRegion.LastRow, mergedRegion.FirstColumn, mergedRegion.LastColumn);
                        // Если не обрабатывали, то добавляем в текущий диапазон оъединенных ячеек текущую ячейку
                        if (IsNewMergedRegion(newMergedRegion, mergedRegions))
                        {
                            mergedRegions.Add(newMergedRegion);
                            destSheet.AddMergedRegion(newMergedRegion);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Копирование ячеек
        /// </summary>
        /// <param name="oldCell"></param>
        /// <param name="newCell"></param>
        private void CopyCell(HSSFCell oldCell, XSSFCell newCell)
        {
            CopyCellStyle(oldCell, newCell);
            CopyCellValue(oldCell, newCell);
        }

        /// <summary>
        /// Копирование содержимого ячеек с соранением типа данных
        /// </summary>
        /// <param name="oldCell"></param>
        /// <param name="newCell"></param>
        private void CopyCellValue(HSSFCell oldCell, XSSFCell newCell)
        {
            switch (oldCell.CellType)
            {
                case CellType.String:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;
                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;
                case CellType.Blank:
                    newCell.SetCellType(CellType.Blank);
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;
                default:
                    break;
            }
        }

        private void CopyCellStyle(HSSFCell oldCell, XSSFCell newCell)
        {
            if (oldCell.CellStyle == null)
                return;
            newCell.CellStyle = newCell.Sheet.Workbook.GetCellStyleAt((short)(oldCell.CellStyle.Index + 1));
        }

        /// <summary>
        /// Поиск объединенных ячеек
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowNum"></param>
        /// <param name="cellNum"></param>
        /// <returns>Коллекция адресов объединенных ячеек</returns>
        private CellRangeAddress GetMergedRegion(HSSFSheet sheet, int rowNum, short cellNum)
        {
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var merged = sheet.GetMergedRegion(i);
                if (merged.IsInRange(rowNum, cellNum))
                {
                    return merged;
                }
            }
            return null;
        }
        /// <summary>
        /// Проверка нахождения ячейки в новом объедененном поле, или в уже обработанном
        /// </summary>
        /// <param name="newMergedRegion"></param>
        /// <param name="mergedRegions"></param>
        /// <returns></returns>
        private bool IsNewMergedRegion(CellRangeAddress newMergedRegion,
                List<CellRangeAddress> mergedRegions)
        {
            return !mergedRegions.Any(r =>
            r.FirstColumn == newMergedRegion.FirstColumn &&
            r.LastColumn == newMergedRegion.LastColumn &&
            r.FirstRow == newMergedRegion.FirstRow &&
            r.LastRow == newMergedRegion.LastRow);
        }


    }
}
