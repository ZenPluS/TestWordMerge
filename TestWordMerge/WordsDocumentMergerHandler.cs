using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using TestWordMerge.Abstract;
using TestWordMerge.Extensions;
using TestWordMerge.Models;

namespace TestWordMerge
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

                //ToDo: Implement the actual merging logic here.

                return null;
            }
            catch (Exception e)
            {
                Logger($"{Header} - An error occurred while merging files: {e.Message}");
                return null;
            }
            finally
            {
                Logger($"{Header} - End");
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
