using System;
using WordMerge.Core;
using WordMerge.Logging;
using WordMerger.Logging;

namespace WordMerger.Factories
{
    internal static class MergeLoggerFactory
    {
        internal static IMergeLogger Create(Action<string> logger)
        {
            switch (true)
            {
                case true when logger == null:
                    {
                        return new NullMergeLogger();
                    }
                default:
                    {
                        return new ActionMergeLogger(logger);
                    }
            }
        }
    }
}
