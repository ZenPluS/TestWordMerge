using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using WordMerge.Constant;

namespace WordMerge.Core
{
    public interface IMergeLogger
    {
        void Log(MergeLogSeverity severity, string message);
        void LogError(string message, List<string> errors);
    }
}