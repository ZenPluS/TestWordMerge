using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordMerge.Abstract;
using WordMerge.Constant;
using WordMerge.Extensions;
using WordMerge.Helpers;
using WordMerge.Models;
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

namespace WordMerge
{
    /// <summary>
    /// Handler to merge files from a source entity into a main Word file.
    /// </summary>
    public sealed class WordsDocumentMergerHandler
        : BaseAbstractHandler<string>
    {
        private readonly IOrganizationService _service;
        private readonly Entity _sourceEntityDocumentToInject;
        private readonly Entity _annotationMainWordFile;
        private readonly List<Couple<string, string>> _configuration;
        private const string Header = nameof(WordsDocumentMergerHandler);

        public WordsDocumentMergerHandler(
            IOrganizationService service,
            Entity sourceEntityDocumentToInject,
            Entity annotationMainWordFile,
            List<Couple<string, string>> configuration,
            Action<string> logger = null)
            : base(logger)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _sourceEntityDocumentToInject = sourceEntityDocumentToInject ?? throw new ArgumentNullException(nameof(sourceEntityDocumentToInject));
            _annotationMainWordFile = annotationMainWordFile ?? throw new ArgumentNullException(nameof(annotationMainWordFile));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }


        public Entity ExcelDocumentsIntoWordHandle(string excelField, string placeholder)
        {
            try
            {
                Logger($"{Header} - Starting to merge Excel file from field '{excelField}' into main Word file.");

                // Scarica il file Excel
                var excelBytes = DownloadFile(_sourceEntityDocumentToInject.ToEntityReference(), excelField);
                if (excelBytes == null)
                {
                    Logger($"{Header} - Excel file not found or empty.");
                    return null;
                }

                // Scarica il file Word principale
                var wordMainFile = _annotationMainWordFile.GetAttributeValue<string>(Annotation.AnnotationDocumentBody);
                var mainBytes = Convert.FromBase64String(wordMainFile);

                // Estrai la tabella dal file Excel
                var excelTable = ConvertExcelToWordTable(excelBytes);

                // Inserisci la tabella al posto del placeholder
                using (var mainStream = new MemoryStream())
                {
                    mainStream.Write(mainBytes, 0, mainBytes.Length);
                    mainStream.Position = 0;

                    using (var mainDoc = WordprocessingDocument.Open(mainStream, true))
                    {
                        var mainBody = mainDoc.MainDocumentPart?.Document.Body ?? new Body();

                        var placeholderParagraph = mainBody
                            .Descendants<Paragraph>()
                            .FirstOrDefault(p => p.InnerText.Contains(placeholder))
                            ?? throw new InvalidOperationException("Placeholder not found in the principal document");

                        if (placeholderParagraph.Parent is Body parentBody)
                        {
                            var elements = parentBody.Elements().ToList();
                            var index = elements.IndexOf(placeholderParagraph);
                            placeholderParagraph.Remove();

                            parentBody.InsertAt(excelTable, index);
                        }
                        else
                        {
                            throw new InvalidOperationException("Placeholder not present in principal document body");
                        }

                        mainDoc.MainDocumentPart?.Document.Save();
                    }

                    var clonedAnnotation = _annotationMainWordFile.CloneEmpty();
                    clonedAnnotation[Annotation.AnnotationDocumentBody] = Convert.ToBase64String(mainStream.ToArray());
                    return clonedAnnotation;
                }
            }
            catch (Exception e)
            {
                Logger($"{Header} - An error occurred while merging Excel file - Exception {e.Message} - Stack {e.StackTrace}");
                return null;
            }
            finally
            {
                Logger($"{Header} - End Excel merge");
            }
        }

        private Table ConvertExcelToWordTable(byte[] excelBytes)
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
        private CellFormat GetCellFormat(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.StyleIndex == null) return null;
            var styleIndex = (int)cell.StyleIndex.Value;
            var stylesheet = doc.WorkbookPart.WorkbookStylesPart?.Stylesheet;
            if (stylesheet == null) return null;
            return stylesheet.CellFormats?.ElementAt(styleIndex) as CellFormat;
        }

        // Estrae le proprietà di stile dalla cella Excel
        private (string bgColor, string fontName, int fontSize, bool bold, bool italic, string fontColor) GetCellStyle(SpreadsheetDocument doc, CellFormat cellFormat)
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


        // Metodo di supporto: estrae il valore di una cella Excel
        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            var value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                return stringTable.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }



        /// <summary>
        /// Handles the merging of Word files from the source entity into the main Word file.
        /// </summary>
        /// <returns>Updated Annotation</returns>
        public Entity WordDocumentsIntoWordHandle()
        {
            try
            {
                Logger($"{Header} - Starting to merge files from entity '{_sourceEntityDocumentToInject.LogicalName}' with ID '{_sourceEntityDocumentToInject.Id}' into main Word file.");

                var allFileFields = _configuration
                    .ConvertAll(i => i.Left);

                var allFiles = allFileFields
                    .ConvertAll(
                        f => (Field: f, File: DownloadFile(_sourceEntityDocumentToInject.ToEntityReference(), f)))
                    .Where(
                        i => i.File != null)
                    .ToDictionary(
                        i => i.Field, i => i.File);

                if (allFileFields.Count > allFiles.Count)
                {
                    Logger("Retrieved files are less than required files inside configuration - exit immediately");
                    return null;
                }

                var wordMainFile = _annotationMainWordFile.GetAttributeValue<string>(Annotation.AnnotationDocumentBody);
                var mainBytes = Convert.FromBase64String(wordMainFile);
                var check = Array.TrueForAll(_configuration.ToArray(),
                    c =>
                    {
                        try
                        {
                            var currentFile = allFiles[c.Left];
                            mainBytes = MergeDocumentsBase64(
                                mainBytes,
                                currentFile,
                                c
                            );

                            return true;
                        }
                        catch (Exception e)
                        {
                            Logger($"{Header} - An error occurred while merging file for field {c.Right} - Exception {e.Message} - Stack {e.StackTrace}");
                            return false;
                        }
                    });

                if (!check)
                    return null;

                var clonedAnnotation = _annotationMainWordFile.CloneEmpty();
                clonedAnnotation[Annotation.AnnotationDocumentBody] = Convert.ToBase64String(mainBytes);

                return clonedAnnotation;
            }
            catch (Exception e)
            {
                Logger($"{Header} - An error occurred while merging files - Exception {e.Message} - Stack {e.StackTrace}");
                return null;
            }
            finally
            {
                Logger($"{Header} - End");
            }
        }

        /// <summary>
        /// Merges two Word documents by replacing a placeholder in the main document with the content of the insert document.
        /// </summary>
        /// <param name="mainBytes"></param>
        /// <param name="insertBytes"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        private byte[] MergeDocumentsBase64(
            byte[] mainBytes,
            byte[] insertBytes,
            Couple<string, string> configuration
            )
        {
            try
            {
                using (var mainStream = new MemoryStream())
                using (var insertStream = new MemoryStream(insertBytes))
                {
                    mainStream.Write(mainBytes, 0, mainBytes.Length);
                    mainStream.Position = 0;

                    using (var mainDoc = WordprocessingDocument.Open(mainStream, true))
                    using (var insertDoc = WordprocessingDocument.Open(insertStream, false))
                    {
                        WordsMergerHelper.CopyStyles(mainDoc, insertDoc);
                        var numberingMap = WordsMergerHelper.CopyNumbering(mainDoc, insertDoc);
                        var imageMap = WordsMergerHelper.CopyImages(mainDoc, insertDoc);

                        var mainBody = mainDoc.MainDocumentPart?.Document.Body ?? new Body();
                        var insertBody = insertDoc.MainDocumentPart?.Document.Body ?? new Body();

                        var placeholderParagraph = mainBody
                            .Descendants<Paragraph>()
                            .FirstOrDefault(p => p.InnerText.Contains(configuration.Right)) ?? throw new InvalidOperationException("Placeholder not found int he principal document");

                        if (placeholderParagraph.Parent is Body parentBody)
                        {
                            var elements = parentBody.Elements().ToList();
                            var index = elements.IndexOf(placeholderParagraph);
                            placeholderParagraph.Remove();

                            foreach (var element in insertBody.Elements())
                            {
                                var imported = element.CloneNode(true);
                                WordsMergerHelper.UpdateImageReferences(imported, imageMap);
                                WordsMergerHelper.UpdateNumberingReferences(imported, numberingMap);
                                parentBody.InsertAt(imported, index++);
                            }
                        }
                        else
                        {
                            throw new InvalidOperationException("Placeholder not present in principal document body");
                        }

                        mainDoc.MainDocumentPart?.Document.Save();
                    }

                    return mainStream.ToArray();
                }
            }
            catch (Exception e)
            {
                Logger($"An Error Occured while merging file for field {configuration.Left} Exception {e.Message} - Stack {e.StackTrace}");
                return null;
            }

        }

        /// <summary>
        /// Downloads a file field attribute from the specified entity reference
        /// </summary>
        /// <param name="entityReference">Entity reference</param>
        /// <param name="attributeName">File field attribute</param>
        /// <returns> byte array as retrieved file</returns>
        private byte[] DownloadFile(
            EntityReference entityReference,
            string attributeName)
        {
            try
            {
                var initializeFileBlocksDownloadRequest = new InitializeFileBlocksDownloadRequest
                {
                    Target = entityReference,
                    FileAttributeName = attributeName
                };

                var initializeFileBlocksDownloadResponse = _service.Execute<InitializeFileBlocksDownloadResponse, InitializeFileBlocksDownloadRequest>(
                    initializeFileBlocksDownloadRequest
                );
                var fileContinuationToken = initializeFileBlocksDownloadResponse.FileContinuationToken;
                var fileSizeInBytes = initializeFileBlocksDownloadResponse.FileSizeInBytes;
                var fileBytes = new List<byte>((int)fileSizeInBytes);

                long offset = 0;
                var blockSizeDownload = !initializeFileBlocksDownloadResponse.IsChunkingSupported ? fileSizeInBytes : 4 * 1024 * 1024;
                if (fileSizeInBytes < blockSizeDownload)
                    blockSizeDownload = fileSizeInBytes;

                while (fileSizeInBytes > 0)
                {
                    var downLoadBlockRequest = new DownloadBlockRequest()
                    {
                        BlockLength = blockSizeDownload,
                        FileContinuationToken = fileContinuationToken,
                        Offset = offset
                    };

                    var downloadBlockResponse = _service.Execute<DownloadBlockResponse, DownloadBlockRequest>(
                        downLoadBlockRequest
                    );
                    fileBytes.AddRange(downloadBlockResponse.Data);
                    fileSizeInBytes -= (int)blockSizeDownload;
                    offset += blockSizeDownload;
                }

                return fileBytes.ToArray();
            }
            catch (Exception e)
            {
                Logger($"An Error Occured while downloading file for field {attributeName} Exception {e.Message} - Stack {e.StackTrace}");
                return null;
            }
        }
    }
}
