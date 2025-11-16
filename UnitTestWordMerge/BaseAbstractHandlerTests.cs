using WordMerge;
using WordMerge.Models;
using Xunit;
using UnitTestWordMerge.Base;
using System.Collections.Generic;

namespace UnitTestWordMerge
{
    public class BaseAbstractHandlerTests : BaseUnitTest
    {
        [Fact]
        public void Constructor_DefaultLogger_AndStructuredMerge_Succeeds()
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
                null // default logger
            );

            var result = handler.MergeConfiguredFiles();

            Assert.True(result.Success);
            Assert.NotNull(result.OutputAnnotation);
            Assert.Empty(result.Errors);
        }

        [Fact]
        public void StructuredMerge_Fails_ForMissingFile()
        {
            IsExcelExecution = false;
            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);

            var cfg = new List<Couple<string, string>>
            {
                new Couple<string, string>("missing_field", "<<CONTENT>>")
            };

            var handler = new WordsDocumentMergerHandler(
                Service,
                source,
                annotation,
                cfg,
                null,
                null
            );

            var result = handler.MergeConfiguredFiles();

            Assert.False(result.Success);
            Assert.NotEmpty(result.Errors);
        }
    }
}