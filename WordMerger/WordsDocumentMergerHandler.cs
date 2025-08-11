using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

namespace WordMerge
{
    /// <summary>
    /// Handler to merge files from a source entity into a main Word file.
    /// </summary>
    public sealed class WordsDocumentMergerHandler
        : BaseAbstractHandler<string>
    {
        private readonly Entity _sourceEntityDocumentToInject;
        private readonly Entity _annotationMainWordFile;
        private readonly List<Couple<string, string>> _configuration;
        private const string Header = nameof(WordsDocumentMergerHandler);
        private readonly IFileDownloader _fileDownloader;
        public WordsDocumentMergerHandler(
            IOrganizationService service,
            Entity sourceEntityDocumentToInject,
            Entity annotationMainWordFile,
            List<Couple<string, string>> configuration,
            IFileDownloader fileDownloader,
            Action<string> logger = null)
            : base(logger)
        {
            if (service == null)
                throw new ArgumentNullException(nameof(service));
            _sourceEntityDocumentToInject = sourceEntityDocumentToInject ?? throw new ArgumentNullException(nameof(sourceEntityDocumentToInject));
            _annotationMainWordFile = annotationMainWordFile ?? throw new ArgumentNullException(nameof(annotationMainWordFile));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _fileDownloader = fileDownloader ?? new FileDownloader(service);
        }

        private Entity MergeDocumentsIntoWordHandle(
            Func<Couple<string, string>, Dictionary<string, byte[]>, byte[], (bool, byte[])> mergeIfWordAction,
            Func<Couple<string, string>, Dictionary<string, byte[]>, byte[], (bool, byte[])> mergeIfExcelAction,
            string logStart,
            string logEnd,
            string errorPrefix)
        {
            try
            {
                Logger($"{Header} - {logStart}");

                var allFileFields = _configuration.ConvertAll(i => i.Left);
                var isExcelDictionary = new Dictionary<string, bool>();

                var allFiles = allFileFields
                    .ConvertAll(f =>
                    {
                        var file = _fileDownloader.DownloadFile(Logger, _sourceEntityDocumentToInject.ToEntityReference(), f, out var isExcel);
                        isExcelDictionary[f] = isExcel;
                        return (Field: f, File: file);
                    })
                    .Where(i => i.File != null)
                    .ToDictionary(i => i.Field, i => i.File);

                if (allFileFields.Count > allFiles.Count)
                {
                    Logger("Retrieved files are less than required files inside configuration - exit immediately");
                    return null;
                }

                var wordMainFile = _annotationMainWordFile.GetAttributeValue<string>(Annotation.AnnotationDocumentBody);
                var mainBytes = Convert.FromBase64String(wordMainFile);

                var check = Array.TrueForAll(_configuration.ToArray(), c =>
                {
                    try
                    {
                        var isExcel = isExcelDictionary[c.Left];
                        var result = isExcel ? mergeIfExcelAction(c, allFiles, mainBytes) : mergeIfWordAction(c, allFiles, mainBytes);
                        mainBytes = result.Item2;
                        return result.Item1;
                    }
                    catch (Exception e)
                    {
                        Logger($"{Header} - {errorPrefix} {c.Right} - Exception {e.Message} - Stack {e.StackTrace}");
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
                Logger($"{Header} - {logEnd}");
            }
        }

        /// <summary>
        /// Merges multiple files from the configured source entity into the main Word document.
        /// Handles both Word and Excel files, replacing placeholders in the main document with the content of the corresponding files.
        /// </summary>
        /// <returns>
        /// An <see cref="Entity"/> representing the main annotation with Word document and merged content, or <c>null</c> if merging fails.
        /// </returns>
        public Entity FileDocumentsIntoWordHandle()
        {
            return MergeDocumentsIntoWordHandle(
                WordDocumentsHandle(),
                ExcelDocumentsHandle(),
                Logs.Start,
                Logs.End,
                Logs.Error
            );
        }

        /// <summary>
        /// Func to handle merging Excel documents.
        /// </summary>
        private Func<Couple<string, string>, Dictionary<string, byte[]>, byte[], (bool, byte[])> ExcelDocumentsHandle()
            => (c, allFiles, mainBytes) =>
            {
                var currentFile = allFiles[c.Left];
                var excelTable = WordsMergerHelper.ConvertExcelToWordTable(currentFile);
                var result = MergeExcelDocumentsBase64(mainBytes, excelTable, c);
                if (result == null)
                    return (false, null);

                mainBytes = result;
                return (true, mainBytes);
            };

        /// <summary>
        /// Func to handle merging Word documents.
        /// </summary>
        private Func<Couple<string, string>, Dictionary<string, byte[]>, byte[], (bool, byte[])> WordDocumentsHandle()
            => (c, allFiles, mainBytes) =>
            {
                var currentFile = allFiles[c.Left];
                var result = MergeWordDocumentsBase64(mainBytes, currentFile, c);
                return ResultManaging(result, ref mainBytes);
            };


        private static (bool, byte[]) ResultManaging(byte[] result, ref byte[] mainBytes)
        {
            if (result == null)
                return (false, null);

            mainBytes = result;
            return (true, mainBytes);
        }

        /// <summary>
        /// Merges an Excel document into a Word document by replacing a placeholder in the main document with the content of the Excel table.
        /// </summary>
        /// <param name="mainBytes">byte array of main file</param>
        /// <param name="excelTable">The table object representing the Excel data to be inserted into the Word document.</param>
        /// <param name="configuration"> field-placeholder configuration</param>
        /// <returns>byte array of main filed modified with inserted file</returns>
        private byte[] MergeExcelDocumentsBase64(
            byte[] mainBytes,
            Table excelTable,
            Couple<string, string> configuration
        )
        {
            try
            {
                using (var mainStream = new MemoryStream())
                {
                    mainStream.Write(mainBytes, 0, mainBytes.Length);
                    mainStream.Position = 0;

                    using (var mainDoc = WordprocessingDocument.Open(mainStream, true))
                    {
                        var mainBody = mainDoc.MainDocumentPart?.Document.Body ?? new Body();

                        var placeholderParagraph = mainBody
                                                       .Descendants<Paragraph>()
                                                       .FirstOrDefault(p => p.InnerText.Contains(configuration.Right))
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
                        return mainStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                Logger($"An Error Occured while merging file for field {configuration.Left} Exception {e.Message} - Stack {e.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Merges two Word documents by replacing a placeholder in the main document with the content of the insert document.
        /// </summary>
        /// <param name="mainBytes">byte array of main file</param>
        /// <param name="insertBytes"> byte array of file to be inserted into main file</param>
        /// <param name="configuration"> field-placeholder configuration</param>
        /// <returns> byte array of main filed modified with inserted file</returns>
        private byte[] MergeWordDocumentsBase64(
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
    }
}
