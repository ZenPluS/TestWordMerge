using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Xrm.Sdk;
using UnitTestWordMerge.Base;
using WordMerge;
using WordMerge.Helpers;
using WordMerge.Models;
using Xunit;

namespace UnitTestWordMerge
{
    public class WordsDocumentMergerHandlerAdditionalTests : BaseUnitTest
    {
        private const string MainAttributeName = "documentbody";

        private static string CreateMainDocBase64(params Paragraph[] paragraphs)
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(paragraphs));
                    mainPart.Document.Save();
                }
                return Convert.ToBase64String(ms.ToArray());
            }
        }

        private static Paragraph Para(string text)
        {
            var run = new Run(new Text(text));
            return new Paragraph(run);
        }

        private static Entity CreateAnnotationWithBody(string base64)
        {
            var ann = new Entity("annotation")
            {
                [MainAttributeName] = base64
            };
            return ann;
        }

        private class TestFileDownloader : IFileDownloader
        {
            private readonly byte[] _content;
            private readonly bool _isExcel;
            private readonly bool _returnNull;

            public TestFileDownloader(byte[] content, bool isExcel = false, bool returnNull = false)
            {
                _content = content;
                _isExcel = isExcel;
                _returnNull = returnNull;
            }

            public byte[] DownloadFile(Action<string> logger, EntityReference source, string attributeName, out bool isExcel)
            {
                isExcel = _isExcel;
                return _returnNull ? null : _content;
            }
        }

        private static byte[] CreateInsertedWordContent(params string[] paragraphTexts)
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    var body = new Body(paragraphTexts.Select(t => new Paragraph(new Run(new Text(t)))));
                    mainPart.Document = new Document(body);
                    mainPart.Document.Save();
                }
                return ms.ToArray();
            }
        }

        [Fact]
        public void MergeConfiguredFiles_Succeeds_WithCaseInsensitivePlaceholder()
        {
            var mainDoc = CreateMainDocBase64(Para("<<CONTENT>>"));
            var annotation = CreateAnnotationWithBody(mainDoc);

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid() // presence
            };

            var insertedDoc = CreateInsertedWordContent("Inserted Paragraph");
            var downloader = new TestFileDownloader(insertedDoc);

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<content>>") // different case
            };

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.True(result.Success);
            Assert.Empty(result.Errors);
            Assert.NotNull(result.OutputAnnotation);
        }

        [Fact]
        public void MergeConfiguredFiles_Succeeds_WhenPlaceholderIsRunTextNotFullParagraph()
        {
            // Paragraph with extra run before placeholder
            var p = new Paragraph(
                new Run(new Text("Intro")),
                new Run(new Text("<<CONTENT>>")),
                new Run(new Text("Outro"))
            );
            var mainDoc = CreateMainDocBase64(p);
            var annotation = CreateAnnotationWithBody(mainDoc);

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var insertedDoc = CreateInsertedWordContent("Inserted Body");
            var downloader = new TestFileDownloader(insertedDoc);

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<CONTENT>>")
            };

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.True(result.Success);
            Assert.Empty(result.Errors);
        }

        [Fact]
        public void MergeConfiguredFiles_Succeeds_WhenPlaceholderSubstringInRun()
        {
            // Placeholder appears inside longer run text
            var p = Para("Prefix <<CONTENT>> Suffix");
            var mainDoc = CreateMainDocBase64(p);
            var annotation = CreateAnnotationWithBody(mainDoc);

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var insertedDoc = CreateInsertedWordContent("Inserted Table");
            var downloader = new TestFileDownloader(insertedDoc);

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<CONTENT>>")
            };

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.True(result.Success);
            Assert.Empty(result.Errors);
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenInvalidBase64()
        {
            var annotation = new Entity("annotation")
            {
                [MainAttributeName] = "NOT_BASE64!!"
            };

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<CONTENT>>")
            };

            var downloader = new TestFileDownloader(CreateInsertedWordContent("X"));

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Contains(result.Errors, e => e?.Contains("not valid Base64") == true);
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenMainAnnotationMissingDocumentBodyAttribute()
        {
            var annotation = new Entity("annotation"); // no documentbody
            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<CONTENT>>")
            };

            var downloader = new TestFileDownloader(CreateInsertedWordContent("X"));

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Contains(result.Errors, e => e.Contains("missing documentbody"));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenEmptyPlaceholderToken()
        {
            var mainDoc = CreateMainDocBase64(Para("<<CONTENT>>"));
            var annotation = CreateAnnotationWithBody(mainDoc);

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var downloader = new TestFileDownloader(CreateInsertedWordContent("Inserted"));

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "") // empty placeholder
            };

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Contains(result.Errors, e => e.Contains("Empty placeholder"));
        }

        [Fact]
        public void MergeConfiguredFiles_Fails_WhenDownloadReturnsNull()
        {
            var mainDoc = CreateMainDocBase64(Para("<<CONTENT>>"));
            var annotation = CreateAnnotationWithBody(mainDoc);

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var downloader = new TestFileDownloader(null, returnNull: true);

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<CONTENT>>")
            };

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.Contains(result.Errors, e => e.Contains("download failed"));
        }

        [Fact]
        public void FileDocumentsIntoWordHandle_Succeeds_OnValidMerge()
        {
            var mainDoc = CreateMainDocBase64(Para("<<CONTENT>>"));
            var annotation = CreateAnnotationWithBody(mainDoc);

            var source = new Entity("incident")
            {
                ["file1"] = Guid.NewGuid()
            };

            var downloader = new TestFileDownloader(CreateInsertedWordContent("Inserted"));

            var config = new List<Couple<string, string>>
            {
                new Couple<string, string>("file1", "<<CONTENT>>")
            };

            var handler = new WordsDocumentMergerHandler(new FakeOrganizationService(), source, annotation, config, downloader, _ => { });
            var legacyResult = handler.FileDocumentsIntoWordHandle();

            Assert.NotNull(legacyResult);
            Assert.True(legacyResult.Attributes.ContainsKey(MainAttributeName));
        }
    }
}