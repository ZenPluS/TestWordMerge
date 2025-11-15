using WordMerge.Globals;

namespace WordMerge.Core
{
    public interface IMergeLogger
    {
        void Log(MergeLogSeverity severity, string message);
    }
}