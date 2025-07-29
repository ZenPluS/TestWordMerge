using System;
using System.IO;
using System.Reflection;
using Microsoft.Xrm.Sdk;
using UnitTestWordMerge.Base;
using UnitTestWordMerge.Helpers;
using WordMerge;
using Xunit;

namespace UnitTestWordMerge
{
    public class WordMergeHandlerUnitTests
        : BaseUnitTest
    {
        [Fact]
        public void MyFirstTest()
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

            var resultAnnotation = wordMergeHandler.Handle();

            var bas64Body = resultAnnotation.GetAttributeValue<string>("documentbody");
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Merged.docx");
            TestHelper.CreateFile(bas64Body, filePath);
        }
    }
}
