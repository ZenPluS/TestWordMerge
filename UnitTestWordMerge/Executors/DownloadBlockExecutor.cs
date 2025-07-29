using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FakeXrmEasy.Abstractions;
using FakeXrmEasy.Abstractions.FakeMessageExecutors;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace UnitTestWordMerge.Executors
{
    public class DownloadBlockExecutor : IFakeMessageExecutor
    {
        public bool CanExecute(OrganizationRequest request)
            => request is DownloadBlockRequest;

        public OrganizationResponse Execute(OrganizationRequest request, IXrmFakedContext ctx)
        {
            var req = (DownloadBlockRequest)request;

            var data = Enumerable.Repeat((byte)0x42, (int)req.BlockLength).ToArray();

            return new DownloadBlockResponse
            {
                Results =
                {
                    { "Data", data }
                }
            };
        }

        public Type GetResponsibleRequestType() => typeof(DownloadBlockRequest);
    }
}
