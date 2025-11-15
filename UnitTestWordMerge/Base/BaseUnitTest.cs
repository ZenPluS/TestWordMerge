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
        protected readonly List<Couple<string, string>> ConfWord;
        protected readonly List<Couple<string, string>> ConfExcel;
        protected readonly Guid FileIdWord;
        protected readonly Guid FileIdExcel;
        protected readonly Guid MainFileId;
        protected bool IsExcelExecution;

        public BaseUnitTest()
        {
            var insertWordFile = TestHelper.ReadFully(TestHelper.GetEmbeddedResourceStream("UnitTestWordMerge.Resources.Insert.docx"));
            var insertExcelFile = TestHelper.ReadFully(TestHelper.GetEmbeddedResourceStream("UnitTestWordMerge.Resources.Insert.xlsx"));
            var mainFile = TestHelper.ReadFully(TestHelper.GetEmbeddedResourceStream("UnitTestWordMerge.Resources.Main.docx"));

            var fakeWordFileId = Guid.Parse("00000000-0000-0000-0000-000000000001");
            var fakeExcelFileId = Guid.Parse("00000000-0000-0000-0000-000000000002");

            ConfWord = new List<Couple<string, string>>
            {
                new Couple<string, string>("dev_fileid", "<<CONTENT>>")
            };

            ConfExcel = new List<Couple<string, string>>
            {
                new Couple<string, string>("dev_fileid", "<<CONTENT>>")
            };

            FileIdWord = Guid.NewGuid();
            var fileWordEntity = new Entity
            {
                ["dev_fileid"] = fakeWordFileId,
                Id = FileIdWord,
                LogicalName = "incident"
            };

            FileIdExcel = Guid.NewGuid();
            var fileExcelEntity = new Entity
            {
                ["dev_fileid"] = fakeExcelFileId,
                Id = FileIdExcel,
                LogicalName = "task"
            };


            InMemoryFileStorage.AddFileWithContext(fakeWordFileId, insertWordFile, "Insert.docx");
            InMemoryFileStorage.AddFileWithContext(fakeExcelFileId, insertExcelFile, "Insert.xlsx");

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
                .AddFakeMessageExecutor<InitializeFileBlocksDownloadRequest>(new InitializeFileBlocksDownloadExecutor(() => IsExcelExecution ? fakeExcelFileId : fakeWordFileId))
                .AddFakeMessageExecutor<DownloadBlockRequest>(new DownloadBlockExecutor(() => IsExcelExecution ? fakeExcelFileId : fakeWordFileId))
                .UseCrud()
                .UseMessages()
                .SetLicense(FakeXrmEasyLicense.NonCommercial)
                .Build();

            context.Initialize(new[] { fileWordEntity, fileExcelEntity, annotationEntity });

            var entityMetadataAnnotation = new EntityMetadata
            {
                LogicalName = "a"
            };
            var attributeAnnotation = new StringAttributeMetadata() { LogicalName = "annotation" };
            entityMetadataAnnotation.SetAttribute(attributeAnnotation);

            var entityMetadataIncident = new EntityMetadata
            {
                LogicalName = "b"
            };
            var attributeIncident = new StringAttributeMetadata() { LogicalName = "incident" };
            entityMetadataIncident.SetAttribute(attributeIncident);

            var entityMetadataTask = new EntityMetadata
            {
                LogicalName = "c"
            };
            var attributeTask = new StringAttributeMetadata() { LogicalName = "task" };
            entityMetadataTask.SetAttribute(attributeTask);

            context.InitializeMetadata(new List<EntityMetadata>()
            {
                entityMetadataAnnotation,
                entityMetadataIncident,
                entityMetadataTask
            });

            Context = context;
            Service = context.GetOrganizationService();
            TracingService = context.GetTracingService();

        }
    }
}
