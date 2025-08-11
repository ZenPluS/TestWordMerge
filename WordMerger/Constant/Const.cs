namespace WordMerge.Constant
{
    internal static class RegexPatterns
    {
        internal static string ExcelPattern = @"^[^<>:""/\\|?*\n]+\.(xls|xlsx|xlsm|xlt|xltx|xltm|xlsb)$";
    }

    internal static class Logs
    {
        internal const string Start = "Starting to merging files";
        internal const string End = "End merge";
        internal const string Error = "An error occurred while merging file for field";
    }
}