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
            var file = Context.GetEntityById("incident", FileId);
            var annotation = Context.GetEntityById("annotation", MainFileId);
            var wordMergeHandler = new WordsDocumentMergerHandler(
                Service,
                file,
                annotation,
                Conf,
                message => TracingService.Trace(message)
                );

            var resultAnnotation = wordMergeHandler.WordDocumentsIntoWordHandle();

            var bas64Body = resultAnnotation.GetAttributeValue<string>("documentbody");
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Merged.docx");
            TestHelper.CreateFile(bas64Body, filePath);
        }

        //Da testare
        [Fact]
        public void MergeExcelFileInFieldIntoWordFileAsAnnotation()
        {
            // Arrange: recupera l'entità sorgente (con file Excel) e l'annotazione Word principale
            var file = Context.GetEntityById("incident", FileId); // L'entità che contiene il file Excel
            var annotation = Context.GetEntityById("annotation", MainFileId); // L'annotazione Word principale

            // Scegli il campo che contiene il file Excel e il placeholder da sostituire
            string excelField = "new_excel_file"; // Sostituisci con il nome effettivo del campo file Excel
            string placeholder = "<<EXCEL_PLACEHOLDER>>"; // Sostituisci con il placeholder effettivo nel Word

            var wordMergeHandler = new WordsDocumentMergerHandler(
                Service,
                file,
                annotation,
                Conf,
                message => TracingService.Trace(message)
            );

            // Act: esegui la fusione
            var resultAnnotation = wordMergeHandler.ExcelDocumentsIntoWordHandle(excelField, placeholder);

            // Assert: verifica che il risultato non sia nullo e che il corpo del documento sia presente
            Assert.NotNull(resultAnnotation);
            var base64Body = resultAnnotation.GetAttributeValue<string>("documentbody");
            Assert.False(string.IsNullOrEmpty(base64Body));

            // Salva il file risultante per ispezione manuale
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MergedExcel.docx");
            TestHelper.CreateFile(base64Body, filePath);

            // (Opzionale) Potresti aggiungere controlli sul contenuto del file Word risultante,
            // ad esempio verificando che il placeholder non sia più presente e che la tabella sia stata inserita.
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenServiceIsNull()
        {
            var entity = new Entity("incident");
            var annotation = new Entity("annotation");
            var config = new List<Couple<string, string>>();

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(null, entity, annotation, config));
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenSourceEntityIsNull()
        {
            var service = new FakeOrganizationService();
            var annotation = new Entity("annotation");
            var config = new List<Couple<string, string>>();

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, null, annotation, config));
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenAnnotationIsNull()
        {
            var service = new FakeOrganizationService();
            var entity = new Entity("incident");
            var config = new List<Couple<string, string>>();

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, entity, null, config));
        }

        [Fact]
        public void Constructor_ThrowsArgumentNullException_WhenConfigurationIsNull()
        {
            var service = new FakeOrganizationService();
            var entity = new Entity("incident");
            var annotation = new Entity("annotation");

            Assert.Throws<ArgumentNullException>(() =>
                new WordsDocumentMergerHandler(service, entity, annotation, null));
        }
    }

    // Dummy implementation for testing purposes
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
