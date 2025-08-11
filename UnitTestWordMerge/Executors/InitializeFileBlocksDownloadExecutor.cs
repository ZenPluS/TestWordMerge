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
        private readonly Func<Guid> _idResolver;
        public InitializeFileBlocksDownloadExecutor(Func<Guid> idResolver)
        {
            _idResolver = idResolver;
        }
        public bool CanExecute(OrganizationRequest request) => request is InitializeFileBlocksDownloadRequest;

        public OrganizationResponse Execute(OrganizationRequest request, IXrmFakedContext ctx)
        {
            var fileId = _idResolver();
            var fileContent = InMemoryFileStorage.GetFileWithContext(fileId, out var fileName) ?? throw new InvalidPluginExecutionException($"File not found for ID {fileId}");

            return new InitializeFileBlocksDownloadResponse
            {
                Results =
                {
                    { "FileContinuationToken", fileId.ToString() },
                    { "FileName", fileName },
                    { "FileSizeInBytes", (long)fileContent.Length },
                },
            };
        }

        public Type GetResponsibleRequestType() => typeof(InitializeFileBlocksDownloadRequest);
    }
}
