using WordMerge.Core;
using WordMerge.Constant;

namespace WordMerger.Logging
{
    internal class NullMergeLogger
        : IMergeLogger
    {
        public void Log(MergeLogSeverity severity, string message) { }
    }
}
