using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using TestWordMerge.Models;

namespace TestWordMerge
{
    public class WordsDocumentMergerHandler
    {
        private readonly List<Couple<string, string>> _placeholderFieldConfiguration;

        public WordsDocumentMergerHandler(
            List<Couple<string, string>> configuration
            )
        {
            _placeholderFieldConfiguration = configuration;

        }

        public Entity MergeDocuments(
            Entity mainDocument
            )
        {

            return null;
        }
    }
}
