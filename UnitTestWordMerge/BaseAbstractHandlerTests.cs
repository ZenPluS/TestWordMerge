using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using UnitTestWordMerge.Base;
using WordMerge;
using WordMerge.Models;
using Xunit;

namespace UnitTestWordMerge
{
    public class BaseAbstractHandlerTests : BaseUnitTest
    {
        [Fact]
        public void Constructor_DefaultLogger_DoesNotThrow()
        {
            var source = Context.GetEntityById("incident", FileIdWord);
            var annotation = Context.GetEntityById("annotation", MainFileId);
            var handler = new WordsDocumentMergerHandler(Service, source, annotation, ConfWord, null);
            var result = handler.FileDocumentsIntoWordHandle();
            Assert.NotNull(result); // Happy path merge
        }
    }
}