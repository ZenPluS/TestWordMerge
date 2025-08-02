using System;
using FakeXrmEasy.Abstractions;
using FakeXrmEasy.Abstractions.FakeMessageExecutors;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using UnitTestWordMerge.Helpers;

namespace UnitTestWordMerge.Executors
{
    public class InitializeFileBlocksDownloadExecutor : IFakeMessageExecutor
    {
        public bool CanExecute(OrganizationRequest request) => request is InitializeFileBlocksDownloadRequest;

        public OrganizationResponse Execute(OrganizationRequest request, IXrmFakedContext ctx)
        {
            var fileWordId = Guid.Parse("00000000-0000-0000-0000-000000000001");
            var fileContent = InMemoryFileStorage.GetFile(fileWordId) ?? throw new InvalidPluginExecutionException($"File not found for ID {fileWordId}");

            return new InitializeFileBlocksDownloadResponse
            {
                Results =
                {
                    { "FileContinuationToken", fileWordId.ToString() },
                    { "FileName", "Insert.docx" },
                    { "FileSizeInBytes", (long)fileContent.Length },
                },
            };
        }

        public Type GetResponsibleRequestType() => typeof(InitializeFileBlocksDownloadRequest);
    }
}
