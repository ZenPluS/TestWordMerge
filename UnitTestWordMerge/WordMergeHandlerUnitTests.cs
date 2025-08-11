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
    public class WordMergeHandlerUnitTests
        : BaseUnitTest
    {
        /// <summary>
        /// Test the functionality of merging a Word file in a field into another file as an annotation.
        /// </summary>
        [Fact]
        public void MergeWordFileInFieldIntoAnotherFileAsAnnotation()
        {
            IsExcelExecution = false;
            var file = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);
            var wordMergeHandler = new WordsDocumentMergerHandler(
                Service,
                file,
                annotation,
                ConfWord,
                null,
                message => TracingService.Trace(message)
                );

            var resultAnnotation = wordMergeHandler.FileDocumentsIntoWordHandle();

            var bas64Body = resultAnnotation.GetAttributeValue<string>("documentbody");
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MergedWordWithWord.docx");
            TestHelper.CreateFile(bas64Body, filePath);
        }

        [Fact]
        public void MergeExcelFileInFieldIntoWordFileAsAnnotation()
        {
            IsExcelExecution = true;
            var file = Context.GetEntityById("task", FileIdExcel);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var wordMergeHandler = new WordsDocumentMergerHandler(
                Service,
                file,
                annotation,
                ConfExcel,
                null,
                message => TracingService.Trace(message)
            );

            var resultAnnotation = wordMergeHandler.FileDocumentsIntoWordHandle();

            var bas64Body = resultAnnotation.GetAttributeValue<string>("documentbody");
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MergedWordWithExcel.docx");
            TestHelper.CreateFile(bas64Body, filePath);
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenServiceIsNull()
        {
            var entity = new Entity("incident");
            var annotation = new Entity("annotation");
            var config = new List<Couple<string, string>>();

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(null, entity, annotation, config, null));
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenSourceEntityIsNull()
        {
            var service = new FakeOrganizationService();
            var annotation = new Entity("annotation");
            var config = new List<Couple<string, string>>();

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, null, annotation, config, null));
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenAnnotationIsNull()
        {
            var service = new FakeOrganizationService();
            var entity = new Entity("incident");
            var config = new List<Couple<string, string>>();

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, entity, null, config, null));
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenConfigurationIsNull()
        {
            var service = new FakeOrganizationService();
            var entity = new Entity("incident");
            var annotation = new Entity("annotation");

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, entity, annotation, null, null));
        }
    }

    public class FakeOrganizationService : IOrganizationService
    {
        public Guid Create(Entity entity) => Guid.NewGuid();
        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) => null;
        public void Update(Entity entity) { }
        public void Delete(string entityName, Guid id) { }
        public OrganizationResponse Execute(OrganizationRequest request) => null;
        public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { }
        public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { }
        public EntityCollection RetrieveMultiple(QueryBase query) => null;
    }
}
