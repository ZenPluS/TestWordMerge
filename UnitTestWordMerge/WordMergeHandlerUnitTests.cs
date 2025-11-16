using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Xrm.Sdk.Query;
using UnitTestWordMerge.Base;
using UnitTestWordMerge.Helpers;
using WordMerge;
using WordMerge.Models;
using Xunit;

namespace UnitTestWordMerge
{
    public class WordMergeHandlerUnitTests : BaseUnitTest
    {
        [Fact]
        public void MergeConfiguredFiles_ReturnsSuccess_ForValidWordInsertion()
        {
            IsExcelExecution = false;
            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                ConfWord,
                null,
                _ => { }
            );

            var result = handler.MergeConfiguredFiles();

            Assert.True(result.Success);
            Assert.NotNull(result.OutputAnnotation);
            Assert.Empty(result.Errors);

            var base64Body = result.OutputAnnotation.GetAttributeValue<string>("documentbody");
            Assert.False(string.IsNullOrWhiteSpace(base64Body));

            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MergedWordWithWord.docx");
            TestHelper.CreateFile(base64Body, filePath);
        }

        [Fact]
        public void MergeConfiguredFiles_ReturnsSuccess_ForValidExcelInsertion()
        {
            IsExcelExecution = true;
            var source = Context.GetEntityById("task", FileIdExcel);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                ConfExcel,
                null,
                _ => { }
            );

            var result = handler.MergeConfiguredFiles();

            Assert.True(result.Success);
            Assert.NotNull(result.OutputAnnotation);
            Assert.Empty(result.Errors);

            var base64Body = result.OutputAnnotation.GetAttributeValue<string>("documentbody");
            Assert.False(string.IsNullOrWhiteSpace(base64Body));

            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MergedWordWithExcel.docx");
            TestHelper.CreateFile(base64Body, filePath);
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenServiceIsNull()
        {
            var entity = new Entity("incident");
            var annotation = new Entity("annotation");
            var config = new List<Couple<string, string>>();
            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(null, entity, annotation, config, null));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenSourceEntityIsNull()
        {
            var service = new FakeOrganizationService();
            var annotation = new Entity("annotation");
            var config = new List<Couple<string, string>>();
            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, null, annotation, config, null));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenAnnotationIsNull()
        {
            var service = new FakeOrganizationService();
            var source = new Entity("incident");
            var config = new List<Couple<string, string>>();
            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, source, null, config, null));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenConfigurationIsNull()
        {
            var service = new FakeOrganizationService();
            var source = new Entity("incident");
            var annotation = new Entity("annotation");
            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, source, annotation, null, null));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenConfiguredFileIsMissing()
        {
            IsExcelExecution = false;
            var badConfig = new List<Couple<string, string>>
            {
                new Couple<string, string>("missing_field", "<<CONTENT>>")
            };

            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                badConfig,
                null,
                _ => { }
            );

            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Null(result.OutputAnnotation);
            Assert.Contains(result.Errors, e => e.Contains("missing_field"));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenPlaceholderNotFound()
        {
            IsExcelExecution = false;
            var wrongPlaceholder = new List<Couple<string, string>>
            {
                new Couple<string, string>("dev_fileid", "<<WRONG>>")
            };

            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                wrongPlaceholder,
                null,
                _ => { }
            );

            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Null(result.OutputAnnotation);
            Assert.Contains(result.Errors, e => e.Contains("not found"));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WithEmptyConfiguration()
        {
            IsExcelExecution = false;
            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var emptyConfig = new List<Couple<string, string>>();

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                emptyConfig,
                null,
                _ => { }
            );

            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Contains(result.Errors, e => e.Contains("empty"));
        }

        [Fact]
        public void FileDocumentsIntoWordHandle_ReturnsNull_OnFailure()
        {
            IsExcelExecution = false;
            var wrongPlaceholder = new List<Couple<string, string>>
            {
                new Couple<string, string>("dev_fileid", "<<WRONG>>")
            };

            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                wrongPlaceholder,
                null,
                _ => { }
            );

            var legacyResult = handler.FileDocumentsIntoWordHandle();
            Assert.Null(legacyResult);
        }
    }

    public class FakeOrganizationService : IOrganizationService
    {
        public Guid Create(Entity entity) { return Guid.NewGuid(); }
        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) { return null; }
        public void Update(Entity entity) { }
        public void Delete(string entityName, Guid id) { }
        public OrganizationResponse Execute(OrganizationRequest request) { return null; }
        public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { }
        public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { }
        public EntityCollection RetrieveMultiple(QueryBase query) { return null; }
    }
}
