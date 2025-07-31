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
            var fileContent = InMemoryFileStorage.GetFile(Guid.Empty) ?? throw new InvalidPluginExecutionException($"File not found for ID {Guid.Empty}");

            return new InitializeFileBlocksDownloadResponse
            {
                Results =
                {
                    { "FileContinuationToken", Guid.Empty.ToString() },
                    { "FileName", "Insert.docx" },
                    { "FileSizeInBytes", (long)fileContent.Length },
                },
            };
        }

        public Type GetResponsibleRequestType() => typeof(InitializeFileBlocksDownloadRequest);
    }
}
