using WordMerge.Core;
using WordMerge.Globals;

namespace WordMerger.Logging
{
    internal class NullMergeLogger
        : IMergeLogger
    {
        public void Log(MergeLogSeverity severity, string message) { }
    }
}
