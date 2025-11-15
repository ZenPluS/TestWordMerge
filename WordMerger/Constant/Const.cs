namespace WordMerge.Constant
{
    internal static class RegexPatterns
    {
        internal static string ExcelPattern = @"^[^<>:""/\\|?*\n]+\.(xls|xlsx|xlsm|xlt|xltx|xltm|xlsb)$";
        internal static string SearchPlaceholders = "<<[A-Z0-9_]+>>";
    }

    internal static class Logs
    {
        internal const string Start = "Starting to merging files";
        internal const string End = "End merge";
    }
}