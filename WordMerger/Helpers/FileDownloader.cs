using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordMerge.Constant;
using WordMerge.Core;
using WordMerge.Extensions;

namespace WordMerge.Helpers
{
    public class FileDownloader : IFileDownloader
    {
        private readonly IOrganizationService _service;

        public FileDownloader(IOrganizationService service)
            => _service = service;

        public byte[] DownloadFile(Action<string> logger, EntityReference entityReference, string attributeName, out bool isExcel)
        {
            isExcel = false;
            try
            {
                var initializeFileBlocksDownloadRequest = new InitializeFileBlocksDownloadRequest
                {
                    Target = entityReference,
                    FileAttributeName = attributeName
                };

                var initializeFileBlocksDownloadResponse =
                    _service.Execute<InitializeFileBlocksDownloadResponse, InitializeFileBlocksDownloadRequest>(
                        initializeFileBlocksDownloadRequest
                    );
                var fileContinuationToken = initializeFileBlocksDownloadResponse.FileContinuationToken;
                var fileSizeInBytes = initializeFileBlocksDownloadResponse.FileSizeInBytes;
                var fileBytes = new List<byte>((int)fileSizeInBytes);

                isExcel = Regex.IsMatch(initializeFileBlocksDownloadResponse.FileName, RegexPatterns.ExcelPattern, RegexOptions.IgnoreCase);

                long offset = 0;
                var blockSizeDownload = !initializeFileBlocksDownloadResponse.IsChunkingSupported
                    ? fileSizeInBytes
                    : 4 * 1024 * 1024;
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
                logger($"An Error Occured while downloading file for field {attributeName} Exception {e.Message} - Stack {e.StackTrace}");
                return null;
            }
        }
    }
}