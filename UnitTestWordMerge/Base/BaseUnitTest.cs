using FakeXrmEasy.Abstractions;
using FakeXrmEasy.Abstractions.Enums;
using FakeXrmEasy.Middleware;
using FakeXrmEasy.Middleware.Crud;
using FakeXrmEasy.Middleware.Messages;
using FakeXrmEasy.Middleware.Pipeline;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using FakeXrmEasy.Extensions;
using UnitTestWordMerge.Executors;
using UnitTestWordMerge.Helpers;
using WordMerge.Models;

namespace UnitTestWordMerge.Base
{
    public class BaseUnitTest
    {
        protected readonly IXrmFakedContext Context;
        protected readonly IOrganizationService Service;
        protected readonly ITracingService TracingService;
        protected readonly List<Couple<string, string>> Conf;
        protected readonly Guid FileId;
        protected readonly Guid MainFileId;

        public BaseUnitTest()
        {
            var insertFile = TestHelper.ReadFully(TestHelper.GetEmbeddedResourceStream("UnitTestWordMerge.Resources.Insert.docx"));
            var mainFile = TestHelper.ReadFully(TestHelper.GetEmbeddedResourceStream("UnitTestWordMerge.Resources.Main.docx"));
            Conf = new List<Couple<string, string>>
            {
                new Couple<string, string>("dev_fileid", "<<CONTENT>>")
            };

            FileId = Guid.NewGuid();
            var fileEntity = new Entity
            {
                ["dev_fileid"] = Guid.Empty,
                Id = FileId,
                LogicalName = "incident"
            };


            InMemoryFileStorage.AddFile(FileId, insertFile);

            MainFileId = Guid.NewGuid();
            var annotationEntity = new Entity()
            {
                ["filename"] = "Main.docx",
                ["documentbody"] = Convert.ToBase64String(mainFile),
                ["mimetype"] = "officedocument.wordprocessingml.document",
                Id = MainFileId,
                LogicalName = "annotation"
            };

            var context = MiddlewareBuilder
                .New()
                .AddCrud()
                .AddPipelineSimulation()
                .UsePipelineSimulation()
                .AddFakeMessageExecutor<InitializeFileBlocksDownloadRequest>(new InitializeFileBlocksDownloadExecutor())
                .AddFakeMessageExecutor<DownloadBlockRequest>(new DownloadBlockExecutor())
                .UseCrud()
                .UseMessages()
                .SetLicense(FakeXrmEasyLicense.NonCommercial)
                .Build();

            context.Initialize(new [] { fileEntity, annotationEntity });

            var entityMetadataAnnotation = new EntityMetadata();
            entityMetadataAnnotation.LogicalName = "dio";
            var attributeAnnotation = new StringAttributeMetadata() { LogicalName = "annotation" };
            entityMetadataAnnotation.SetAttribute(attributeAnnotation);

            var entityMetadataIncident = new EntityMetadata();
            entityMetadataIncident.LogicalName = "cristo";
            var attributeIncident = new StringAttributeMetadata() { LogicalName = "incident" };
            entityMetadataIncident.SetAttribute(attributeIncident);

            context.InitializeMetadata(new List<EntityMetadata>()
            {
                entityMetadataAnnotation,
                entityMetadataIncident
            });

            Context = context;
            Service = context.GetOrganizationService();
            TracingService = context.GetTracingService();

        }
    }
}
