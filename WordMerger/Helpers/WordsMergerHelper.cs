using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using Italic = DocumentFormat.OpenXml.Wordprocessing.Italic;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace WordMerge.Helpers
{
    internal static class WordsMergerHelper
    {
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
                var sheet = spreadsheet.WorkbookPart.Workbook.Sheets.Elements<Sheet>().First();
                var worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheet.Id);
                var rows = worksheetPart.Worksheet.Descendants<Row>();

                var table = new Table();

                // Applica bordi a tutta la tabella
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
                table.AppendChild(tableProperties);

                foreach (var row in rows)
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
            if (cell.StyleIndex == null) return null;
            var styleIndex = (int)cell.StyleIndex.Value;
            var stylesheet = doc.WorkbookPart.WorkbookStylesPart?.Stylesheet;
            if (stylesheet == null) return null;
            return stylesheet.CellFormats?.ElementAt(styleIndex) as CellFormat;
        }

        // Estrae le proprietà di stile dalla cella Excel
        internal static  (string bgColor, string fontName, int fontSize, bool bold, bool italic, string fontColor) GetCellStyle(SpreadsheetDocument doc, CellFormat cellFormat)
        {
            string bgColor = null, fontName = null, fontColor = null;
            int fontSize = 0;
            bool bold = false, italic = false;

            var stylesheet = doc.WorkbookPart.WorkbookStylesPart?.Stylesheet;
            if (cellFormat != null && stylesheet != null)
            {
                // Background color
                if (cellFormat.FillId != null && stylesheet.Fills != null)
                {
                    var fill = stylesheet.Fills.ElementAt((int)cellFormat.FillId.Value) as Fill;
                    var patternFill = fill?.PatternFill;
                    if (patternFill?.ForegroundColor != null)
                        bgColor = patternFill.ForegroundColor.Rgb?.Value?.Substring(2); // Rimuove "FF" alpha
                    else if (patternFill?.BackgroundColor != null)
                        bgColor = patternFill.BackgroundColor.Rgb?.Value?.Substring(2);
                }

                // Font
                if (cellFormat.FontId != null && stylesheet.Fonts != null)
                {
                    var font = stylesheet.Fonts.ElementAt((int)cellFormat.FontId.Value) as Font;
                    fontName = font?.FontName?.Val;
                    fontSize = font?.FontSize != null ? (int)font.FontSize.Val.Value : 0;
                    bold = font?.Bold != null;
                    italic = font?.Italic != null;
                    fontColor = font?.Color?.Rgb?.Value?.Substring(2); // Rimuove "FF" alpha
                }
            }
            return (bgColor, fontName, fontSize, bold, italic, fontColor);
        }

        internal static  string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell?.CellValue == null)
                return string.Empty;

            var value = cell.CellValue.InnerText;
            if (cell.DataType == null || cell.DataType.Value != CellValues.SharedString)
                return value;

            var stringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
            return stringTable.ElementAt(int.Parse(value)).InnerText;
        }
    }
}
