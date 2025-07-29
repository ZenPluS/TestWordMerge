using DocumentFormat.OpenXml.Packaging;
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
using WordMerge.Models;

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

        /// <summary>
        /// Handles the merging of files from the source entity into the main Word file.
        /// </summary>
        /// <returns></returns>
        public Entity Handle()
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
                            var updatedMain = MergeDocumentsBase64(
                                mainBytes,
                                currentFile,
                                c
                            );
                            mainBytes = updatedMain;
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
                            .FirstOrDefault(p => p.InnerText.Contains(configuration.Right));

                        if (placeholderParagraph == null)
                            throw new InvalidOperationException("Placeholder non trovato nel documento principale.");

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
                            throw new InvalidOperationException("Il placeholder non si trova nel body principale.");
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
