using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using TestWordMerge.Extensions;
using TestWordMerge.Models;

namespace TestWordMerge
{
    public sealed class WordsDocumentMergerHandler
    {
        private readonly IOrganizationService _service;
        private readonly Entity _sourceEntityDocumentToInject;
        private readonly Entity _annotationMainWordFile;
        private readonly List<Couple<string, string>> _configuration;
        private readonly Action<string> _logger;

        public WordsDocumentMergerHandler(
            IOrganizationService service,
            Entity sourceEntityDocumentToInject,
            Entity annotationMainWordFile,
            List<Couple<string, string>> configuration,
            Action<string> logger = null)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _sourceEntityDocumentToInject = sourceEntityDocumentToInject ?? throw new ArgumentNullException(nameof(sourceEntityDocumentToInject));
            _annotationMainWordFile = annotationMainWordFile ?? throw new ArgumentNullException(nameof(annotationMainWordFile));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _logger = logger ?? (message => Debug.WriteLine(message));
        }

        public Entity Handle()
        {
            var allFileFields = _configuration
                .Select(i => i.Left)
                .ToList();

            var allFiles = allFileFields
                .ConvertAll(
                    f => (Field: f, File: DownloadFile(_sourceEntityDocumentToInject.ToEntityReference(), f)))
                .Where(
                    i => i.File != null)
                .ToDictionary(
                    i => i.Field,
                    i => i.File);

            return null;
        }

        /// <summary>  
        /// Download a file from Dynamics 365 using the IOrganizationService.
        /// </summary>  
        /// <param name="entityReference"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>  
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
                _logger(e.Message);
                return null;
            }
        }
    }
}
