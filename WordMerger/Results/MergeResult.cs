using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;

namespace WordMerge.Results
{
    public sealed class MergeResult
    {
        public bool Success { get; }
        public Entity OutputAnnotation { get; }
        public IReadOnlyList<string> Errors { get; }

        private MergeResult(bool success, Entity outputAnnotation, IReadOnlyList<string> errors)
        {
            Success = success;
            OutputAnnotation = outputAnnotation;
            Errors = errors ?? Array.Empty<string>();
        }

        public static MergeResult Ok(Entity annotation) => new MergeResult(true, annotation, Array.Empty<string>());
        public static MergeResult Fail(IEnumerable<string> errors) => new MergeResult(false, null, new List<string>(errors ?? Array.Empty<string>()));
    }
}