using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordMerge.Abstract;
using WordMerge.Constant;
using WordMerge.Core;
using WordMerge.Extensions;
using WordMerge.Helpers;
using WordMerge.Models;
using WordMerge.Results;
using WordMerger.Factories;

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
        private readonly IFileDownloader _fileDownloader;
        private readonly IMergeLogger _mergeLogger;
        private const string Header = nameof(WordsDocumentMergerHandler);

        public WordsDocumentMergerHandler(
            IOrganizationService service,
            Entity sourceEntityDocumentToInject,
            Entity annotationMainWordFile,
            List<Couple<string, string>> configuration,
            IFileDownloader fileDownloader,
            Action<string> logger = null)
            : base(logger)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            _sourceEntityDocumentToInject = sourceEntityDocumentToInject ?? throw new ArgumentNullException(nameof(sourceEntityDocumentToInject));
            _annotationMainWordFile = annotationMainWordFile ?? throw new ArgumentNullException(nameof(annotationMainWordFile));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _fileDownloader = fileDownloader ?? new FileDownloader(service);
            _mergeLogger = MergeLoggerFactory.Create(logger);
        }

        #region Public API

        /// <summary>
        /// Structured merge returning status and errors.
        /// </summary>
        public MergeResult MergeConfiguredFiles()
        {
            var errors = new List<string>();
            _mergeLogger.Log(MergeLogSeverity.Info, $"{Header} - {Logs.Start}");

            if (_configuration.Count == 0)
            {
                _mergeLogger.LogError(Errors.ConfigListEmpty, errors);
                return MergeResult.Fail(errors);
            }

            if (!_annotationMainWordFile.Contains(Annotation.AnnotationDocumentBody))
            {
                _mergeLogger.LogError(Errors.MainAnnotationMissingDocBodyAttribute, errors);
                return MergeResult.Fail(errors);
            }

            var mainBase64 = _annotationMainWordFile.GetAttributeValue<string>(Annotation.AnnotationDocumentBody);
            if (string.IsNullOrWhiteSpace(mainBase64))
            {
                _mergeLogger.LogError(Errors.MainAnnotationBodyEmpty, errors);
                return MergeResult.Fail(errors);
            }

            byte[] mainBytes;
            try
            {
                mainBytes = Convert.FromBase64String(mainBase64);
            }
            catch (FormatException fe)
            {
                _mergeLogger.LogError(string.Format(Errors.MainAnnotationBodyNotValidBase64, fe.Message), errors);
                return MergeResult.Fail(errors);
            }

            var fileDescriptors = new List<(Couple<string, string> Config, byte[] Bytes, bool IsExcel)>();
            foreach (var cfg in _configuration)
            {
                if (!_sourceEntityDocumentToInject.Contains(cfg.Left) || _sourceEntityDocumentToInject[cfg.Left] == null)
                {
                    _mergeLogger.LogError(string.Format(Errors.FileFieldMissing, cfg.Left), errors);
                    continue;
                }

                var bytes = _fileDownloader.DownloadFile(m => _mergeLogger.Log(MergeLogSeverity.Warning, m),
                    _sourceEntityDocumentToInject.ToEntityReference(),
                    cfg.Left,
                    out var isExcel);

                if (bytes == null)
                {
                    _mergeLogger.LogError(string.Format(Errors.FileFieldEmptyOrErrorDownloading, cfg.Left), errors);
                    continue;
                }

                fileDescriptors.Add((cfg, bytes, isExcel));
            }

            if (errors.Count > 0)
            {
                return MergeResult.Fail(errors);
            }

            try
            {
                mainBytes = ApplyAllMergesInSingleSession(mainBytes, fileDescriptors, errors);
            }
            catch (Exception ex)
            {
                _mergeLogger.LogError(string.Format(Errors.UnexpectedErrorDuringMerge, ex.Message), errors);
            }

            if (errors.Count > 0)
                return MergeResult.Fail(errors);

            var cloned = _annotationMainWordFile.CloneEmpty();
            cloned[Annotation.AnnotationDocumentBody] = Convert.ToBase64String(mainBytes);

            _mergeLogger.Log(MergeLogSeverity.Info, $"{Header} - {Logs.End}");
            return MergeResult.Ok(cloned);
        }

        /// <summary>
        /// Legacy method kept for backward compatibility (returns Entity or null).
        /// Internally uses structured merge.
        /// </summary>
        public Entity FileDocumentsIntoWordHandle()
        {
            var result = MergeConfiguredFiles();
            return result.Success ? result.OutputAnnotation : null;
        }

        #endregion

        #region Core Merge Logic
        private byte[] ApplyAllMergesInSingleSession(
            byte[] mainBytes,
            List<(Couple<string, string> Config, byte[] Bytes, bool IsExcel)> files,
            List<string> errors)
        {
            using (var mainStream = new MemoryStream())
            {
                mainStream.Write(mainBytes, 0, mainBytes.Length);
                mainStream.Position = 0;

                using (var mainDoc = WordprocessingDocument.Open(mainStream, true))
                {
                    var body = mainDoc.MainDocumentPart?.Document.Body ?? new Body();

                    foreach (var ((left, placeholderToken), bytes, isExcel) in files)
                    {
                        if (string.IsNullOrWhiteSpace(placeholderToken))
                        {
                            _mergeLogger.LogError(string.Format(Errors.EmptyPlaceholder, left), errors);
                            continue;
                        }

                        var paragraph = WordsMergerHelper.FindPlaceholderParagraph(body, placeholderToken);
                        if (paragraph == null)
                        {
                            _mergeLogger.LogError(string.Format(Errors.PlaceholderNotFoundInMainDocument, placeholderToken), errors);
                            continue;
                        }

                        if (!(paragraph.Parent is Body parentBody))
                        {
                            _mergeLogger.LogError(string.Format(Errors.PlaceholderNotFoundInMainDocumentBody, placeholderToken), errors);
                            continue;
                        }

                        var index = parentBody.Elements().ToList().IndexOf(paragraph);
                        paragraph.Remove();

                        if (isExcel)
                        {
                            var table = WordsMergerHelper.ConvertExcelToWordTable(bytes);
                            parentBody.InsertAt(table, index);
                        }
                        else
                        {
                            using (var insertStream = new MemoryStream(bytes))
                            using (var insertDoc = WordprocessingDocument.Open(insertStream, false))
                            {
                                WordsMergerHelper.CopyStyles(mainDoc, insertDoc);
                                var numberingMap = WordsMergerHelper.CopyNumbering(mainDoc, insertDoc);
                                var imageMap = WordsMergerHelper.CopyImages(mainDoc, insertDoc);

                                var insertBody = insertDoc.MainDocumentPart?.Document.Body ?? new Body();
                                foreach (var element in insertBody.Elements())
                                {
                                    var imported = element.CloneNode(true);
                                    WordsMergerHelper.UpdateImageReferences(imported, imageMap);
                                    WordsMergerHelper.UpdateNumberingReferences(imported, numberingMap);
                                    parentBody.InsertAt(imported, index++);
                                }
                            }
                        }
                    }

                    mainDoc.MainDocumentPart?.Document.Save();
                }

                return mainStream.ToArray();
            }
        }

        #endregion
    }
}