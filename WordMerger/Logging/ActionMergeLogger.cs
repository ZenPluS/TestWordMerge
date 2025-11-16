using System;
using System.Collections.Generic;
using WordMerge.Core;
using WordMerge.Constant;

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

        public void LogError(string message, List<string> errors)
        {
            errors.Add(message);
            _action?.Invoke($"[{MergeLogSeverity.Error}] {message}");
        }
    }
}