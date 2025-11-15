using System;
using WordMerge.Core;
using WordMerge.Globals;

namespace WordMerge.Logging
{
    internal sealed class ActionMergeLogger
        : IMergeLogger
    {
        private readonly Action<string> _action;

        public ActionMergeLogger(Action<string> action) => _action = action ?? (_ => { });

        public void Log(MergeLogSeverity severity, string message)
        {
            _action?.Invoke($"[{severity}] {message}");
        }
    }
}