using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WordMerge.Constant;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using Italic = DocumentFormat.OpenXml.Wordprocessing.Italic;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;

namespace WordMerge.Helpers
{
    internal static class WordsMergerHelper
    {
        private static readonly Regex PlaceholderRegex = new Regex(RegexPatterns.SearchPlaceholders, RegexOptions.Compiled | RegexOptions.IgnoreCase);

        internal static void CopyStyles(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
        {
            var mainStylePart = mainDoc.MainDocumentPart?.StyleDefinitionsPart ?? mainDoc.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
            var insertStylePart = insertDoc.MainDocumentPart?.StyleDefinitionsPart;

            if (insertStylePart == null)
                return;

            var insertStyles = insertStylePart.Styles;
            var mainStyles = mainStylePart?.Styles ?? new Styles();
            if (mainStylePart?.Styles == null && mainStylePart != null)
                mainStylePart.Styles = mainStyles;

            foreach (var style in insertStyles?.Elements<Style>().ToList() ?? new List<Style>())
            {
                var existing = mainStyles.Elements<Style>().FirstOrDefault(s => s.StyleId == style.StyleId);
                existing?.Remove();
                mainStyles.Append(style.CloneNode(true));
            }

            mainStylePart?.Styles.Save();

            if (insertDoc.MainDocumentPart.ThemePart == null)
                return;

            var mainThemePart = mainDoc.MainDocumentPart?.ThemePart;
            if (mainThemePart != null)
                mainDoc.MainDocumentPart.DeletePart(mainThemePart);

            mainDoc.MainDocumentPart?.AddNewPart<ThemePart>();
            mainDoc.MainDocumentPart?.ThemePart?.FeedData(insertDoc.MainDocumentPart.ThemePart.GetStream());
        }

        internal static Dictionary<int, int> CopyNumbering(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
        {
            var insertPart = insertDoc.MainDocumentPart?.NumberingDefinitionsPart;
            if (insertPart == null) return new Dictionary<int, int>();

            var mainPart = mainDoc.MainDocumentPart?.NumberingDefinitionsPart ?? mainDoc.MainDocumentPart?.AddNewPart<NumberingDefinitionsPart>();

            var mainNumbering = mainPart?.Numbering ?? new Numbering();
            var insertNumbering = insertPart.Numbering;

            var maxAbstractNumId = mainNumbering.Elements<AbstractNum>()
                .ToList()
                .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(0).Max();

            var maxNumId = mainNumbering.Elements<NumberingInstance>()
                .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max();

            var abstractNumMap = new Dictionary<int, int>();
            foreach (var abstractNum in insertNumbering.Elements<AbstractNum>())
            {
                maxAbstractNumId++;
                var oldId = abstractNum.AbstractNumberId?.Value ?? 0;
                var newAbstractNum = (AbstractNum)abstractNum.CloneNode(true);
                newAbstractNum.AbstractNumberId = new Int32Value(maxAbstractNumId);
                mainNumbering.AppendChild(newAbstractNum);
                abstractNumMap[oldId] = maxAbstractNumId;
            }

            var numMap = new Dictionary<int, int>();
            foreach (var num in insertNumbering.Elements<NumberingInstance>())
            {
                maxNumId++;
                var oldId = num.NumberID?.Value ?? 0;
                var newNum = (NumberingInstance)num.CloneNode(true);
                newNum.NumberID = new Int32Value(maxNumId);

                var absNumId = newNum.Descendants<AbstractNumId>().FirstOrDefault();
                if (absNumId != null && abstractNumMap.TryGetValue(absNumId.Val?.Value ?? -1, out var value))
                    absNumId.Val = value;

                mainNumbering.AppendChild(newNum);
                numMap[oldId] = maxNumId;
            }

            mainPart?.Numbering.Save();
            return numMap;
        }

        internal static Dictionary<string, string> CopyImages(WordprocessingDocument mainDoc, WordprocessingDocument insertDoc)
        {
            var imageMap = new Dictionary<string, string>();
            var insertImageParts = insertDoc.MainDocumentPart?.ImageParts.ToList();
            foreach (var imagePart in insertImageParts ?? new List<ImagePart>())
            {
                var oldRelId = insertDoc.MainDocumentPart?.GetIdOfPart(imagePart);
                var imageStream = imagePart.GetStream();
                var newImagePart = mainDoc.MainDocumentPart?.AddImagePart(imagePart.ContentType);
                newImagePart?.FeedData(imageStream);
                if (newImagePart == null)
                    continue;

                var newRelId = mainDoc.MainDocumentPart?.GetIdOfPart(newImagePart);
                if (oldRelId != null) imageMap[oldRelId] = newRelId;
            }
            return imageMap;
        }

        internal static void UpdateImageReferences(OpenXmlElement element, Dictionary<string, string> imageMap)
        {
            foreach (var drawing in element.Descendants<Drawing>().ToList())
            {
                foreach (var blip in drawing.Descendants<Blip>().ToList())
                {
                    var embed = blip.Embed?.Value;
                    if (embed != null && imageMap.TryGetValue(embed, out var value))
                    {
                        blip.Embed.Value = value;
                    }
                }
            }
        }

        internal static void UpdateNumberingReferences(OpenXmlElement element, Dictionary<int, int> numberingMap)
        {
            foreach (var numPr in element.Descendants<NumberingProperties>())
            {
                var numId = numPr.NumberingId;
                if (numId?.Val != null && numberingMap.TryGetValue(numId.Val.Value, out var value))
                {
                    numId.Val = value;
                }
            }
        }

        internal static  Table ConvertExcelToWordTable(byte[] excelBytes)
        {
            using (var excelStream = new MemoryStream(excelBytes))
            using (var spreadsheet = SpreadsheetDocument.Open(excelStream, false))
            {
                var sheet = spreadsheet.WorkbookPart?.Workbook.Sheets?.Elements<Sheet>().First();
                var worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart?.GetPartById(sheet?.Id ?? string.Empty);
                var tableDefPart = worksheetPart?.TableDefinitionParts.FirstOrDefault();
                string excelTableStyle = null;
                if (tableDefPart?.Table.TableStyleInfo != null)
                    excelTableStyle = tableDefPart.Table.TableStyleInfo.Name?.Value;

                var rows = worksheetPart?.Worksheet.Descendants<Row>();
                var table = new Table();
                var tableProperties = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                );

                if (!string.IsNullOrEmpty(excelTableStyle))
                    tableProperties.Append(new TableStyle { Val = excelTableStyle });

                table.AppendChild(tableProperties);

                foreach (var row in rows ?? new List<Row>())
                {
                    var tableRow = new TableRow();
                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellValue = GetCellValue(spreadsheet, cell);

                        // Ottieni formattazione cella
                        var cellFormat = GetCellFormat(spreadsheet, cell);
                        var (bgColor, fontName, fontSize, bold, italic, fontColor) = GetCellStyle(spreadsheet, cellFormat);

                        // Crea run con font e colore
                        var runProps = new RunProperties();
                        if (!string.IsNullOrEmpty(fontName))
                            runProps.Append(new RunFonts() { Ascii = fontName, HighAnsi = fontName });
                        if (fontSize > 0)
                            runProps.Append(new FontSize() { Val = (fontSize * 2).ToString() }); // Word vuole la metà
                        if (bold)
                            runProps.Append(new Bold());
                        if (italic)
                            runProps.Append(new Italic());
                        if (!string.IsNullOrEmpty(fontColor))
                            runProps.Append(new Color() { Val = fontColor });

                        var run = new Run(runProps, new Text(cellValue ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });

                        // Crea cella con shading (colore di sfondo)
                        var cellProps = new TableCellProperties();
                        if (!string.IsNullOrEmpty(bgColor))
                        {
                            cellProps.Append(new Shading()
                            {
                                Val = ShadingPatternValues.Clear,
                                Color = "auto",
                                Fill = bgColor
                            });
                        }

                        var tableCell = new TableCell(cellProps, new Paragraph(run));
                        tableRow.Append(tableCell);
                    }
                    table.Append(tableRow);
                }
                return table;
            }
        }

        // Ottieni il formato della cella (CellFormat) da Stylesheet
        internal static  CellFormat GetCellFormat(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.StyleIndex == null)
                return null;

            var styleIndex = (int)cell.StyleIndex.Value;
            var stylesheet = doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;

            return stylesheet?.CellFormats?.ElementAt(styleIndex) as CellFormat;
        }

        internal static (string bgColor, string fontName, int fontSize, bool bold, bool italic, string fontColor) GetCellStyle(SpreadsheetDocument doc, CellFormat cellFormat)
        {
            string bgColor = null;
            var fontSize = 0;

            var stylesheet = doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
            if (cellFormat == null || stylesheet == null)
                return (null, null, fontSize, false, false, null);

            if (cellFormat.FillId != null && stylesheet.Fills != null)
            {
                var fill = stylesheet.Fills.ElementAt((int)cellFormat.FillId.Value) as Fill;
                var patternFill = fill?.PatternFill;
                if (patternFill != null)
                {
                    if (patternFill.PatternType == null || patternFill.PatternType.Value != PatternValues.None)
                    {
                        if (patternFill.ForegroundColor != null && !string.IsNullOrEmpty(patternFill.ForegroundColor.Rgb?.Value))
                            bgColor = patternFill.ForegroundColor.Rgb.Value;
                        else if (patternFill.BackgroundColor != null && !string.IsNullOrEmpty(patternFill.BackgroundColor.Rgb?.Value))
                            bgColor = patternFill.BackgroundColor.Rgb.Value;
                    }
                }
                if (!string.IsNullOrEmpty(bgColor) && bgColor.Length == 8)
                    bgColor = bgColor.Substring(2);
            }

            if (cellFormat.FontId == null || stylesheet.Fonts == null)
                return (bgColor, null, fontSize, false, false, null);

            var font = stylesheet.Fonts.ElementAt((int) cellFormat.FontId.Value) as Font;
            string fontName = font?.FontName?.Val;
            fontSize = font?.FontSize != null ? (int)font.FontSize?.Val?.Value : 0;
            var bold = font?.Bold != null;
            var italic = font?.Italic != null;
            if (font?.Color == null || string.IsNullOrEmpty(font.Color.Rgb?.Value))
                return (bgColor, fontName, fontSize, bold, italic, null);

            var fontColor = font.Color.Rgb.Value;
            if (fontColor.Length == 8)
                fontColor = fontColor.Substring(2);
            return (bgColor, fontName, fontSize, bold, italic, fontColor);
        }

        internal static  string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell?.CellValue == null)
                return string.Empty;

            var value = cell.CellValue.InnerText;
            if (cell.DataType == null || cell.DataType.Value != CellValues.SharedString)
                return value;

            var stringTable = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            return stringTable?.ElementAt(int.Parse(value)).InnerText;
        }

        internal static Paragraph FindPlaceholderParagraph(Body body, string placeholder)
        {
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                var fullText = paragraph.InnerText;
                if (string.Equals(fullText, placeholder, StringComparison.OrdinalIgnoreCase))
                    return paragraph;

                // Fallback: exact token present as distinct run text
                var runs = paragraph.Descendants<Run>().Select(r => r.InnerText).ToList();
                if (runs.Any(r => string.Equals(r, placeholder, StringComparison.OrdinalIgnoreCase)))
                    return paragraph;

                // If configuration expects tokens like <<CONTENT>> allow exact match within a run (but not substring of other characters)
                if (PlaceholderRegex.IsMatch(placeholder) && runs.Any(r => r.Contains(placeholder)))
                    return paragraph;
            }
            return null;
        }
    }
}
