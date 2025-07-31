using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using WordMerge.Helpers;
using Xunit;

namespace UnitTestWordMerge
{
    public class WordsMergerHelperTests
    {
        [Fact]
        public void CopyStyles_DoesNotThrow_WhenInsertStylePartIsNull()
        {
            using (var mainStream = new MemoryStream())
            using (var insertStream = new MemoryStream())
            using (var mainDoc = WordprocessingDocument.Create(mainStream, WordprocessingDocumentType.Document, true))
            using (var insertDoc =
                   WordprocessingDocument.Create(insertStream, WordprocessingDocumentType.Document, true))
            {
                // No StyleDefinitionsPart in insertDoc
                mainDoc.AddMainDocumentPart();
                insertDoc.AddMainDocumentPart();

                // Should not throw
                WordsMergerHelper.CopyStyles(mainDoc, insertDoc);
            }
        }

        [Fact]
        public void CopyNumbering_ReturnsEmptyDictionary_WhenInsertPartIsNull()
        {
            using (var mainStream = new MemoryStream())
            using (var insertStream = new MemoryStream())
            using (var mainDoc = WordprocessingDocument.Create(mainStream, WordprocessingDocumentType.Document, true))
            using (var insertDoc =
                   WordprocessingDocument.Create(insertStream, WordprocessingDocumentType.Document, true))
            {
                mainDoc.AddMainDocumentPart();
                insertDoc.AddMainDocumentPart();
                // No NumberingDefinitionsPart in insertDoc

                var result = WordsMergerHelper.CopyNumbering(mainDoc, insertDoc);
                Assert.NotNull(result);
                Assert.Empty(result);
            }
        }

        [Fact]
        public void CopyImages_ReturnsEmptyDictionary_WhenNoImages()
        {
            using (var mainStream = new MemoryStream())
            using (var insertStream = new MemoryStream())
            using (var mainDoc = WordprocessingDocument.Create(mainStream, WordprocessingDocumentType.Document, true))
            using (var insertDoc =
                   WordprocessingDocument.Create(insertStream, WordprocessingDocumentType.Document, true))
            {
                mainDoc.AddMainDocumentPart();
                insertDoc.AddMainDocumentPart();
                // No images in insertDoc

                var result = WordsMergerHelper.CopyImages(mainDoc, insertDoc);
                Assert.NotNull(result);
                Assert.Empty(result);
            }
        }

        [Fact]
        public void UpdateImageReferences_DoesNotThrow_WhenNoDrawings()
        {
            var element = new Paragraph();
            var imageMap = new Dictionary<string, string>();
            // Should not throw
            WordsMergerHelper.UpdateImageReferences(element, imageMap);
        }

        [Fact]
        public void UpdateNumberingReferences_DoesNotThrow_WhenNoNumberingProperties()
        {
            var element = new Paragraph();
            var numberingMap = new Dictionary<int, int>();
            // Should not throw
            WordsMergerHelper.UpdateNumberingReferences(element, numberingMap);
        }
    }
}