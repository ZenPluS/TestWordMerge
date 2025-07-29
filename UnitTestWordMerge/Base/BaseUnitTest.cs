using FakeXrmEasy.Abstractions;
using FakeXrmEasy.Abstractions.Enums;
using FakeXrmEasy.Middleware;
using FakeXrmEasy.Middleware.Crud;
using FakeXrmEasy.Middleware.Messages;
using FakeXrmEasy.Middleware.Pipeline;
using Microsoft.Xrm.Sdk;

namespace UnitTestWordMerge.Base
{
    public class BaseUnitTest
    {
        protected readonly IXrmFakedContext Context;
        protected readonly IOrganizationService Service;
        protected readonly ITracingService TracingService;

        public BaseUnitTest()
        {
            var context =  MiddlewareBuilder
                .New()
                .AddCrud()
                .AddPipelineSimulation()
                .UsePipelineSimulation()
                .UseCrud()
                .UseMessages()
                .SetLicense(FakeXrmEasyLicense.NonCommercial)
                .Build();

            Context = context;
            Service = context.GetOrganizationService();
            TracingService = context.GetTracingService();
        }
    }
}
