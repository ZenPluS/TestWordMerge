using WordMerge;
using UnitTestWordMerge.Base;
using Xunit;

namespace UnitTestWordMerge
{
    public class WordMergeHandlerUnitTests
        : BaseUnitTest
    {
        [Fact]
        public void MyFirstTest()
        {
            var wordMergeHandler = new WordsDocumentMergerHandler(
                Service,
                null,
                null,
                null,
                message=> TracingService.Trace(message)
                );

            wordMergeHandler.Handle();
        }
    }
}
