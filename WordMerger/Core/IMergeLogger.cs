using System.Collections.Generic;
using WordMerge.Constant;

namespace WordMerge.Core
{
    public interface IMergeLogger
    {
        void Log(MergeLogSeverity severity, string message);
        void LogError(string message, List<string> errors);
    }
}