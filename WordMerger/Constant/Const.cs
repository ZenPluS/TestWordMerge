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

    internal static class Annotation
    {
        internal const string AnnotationDocumentBody = "documentbody";
    }

    internal static class Errors
    {
        internal const string ConfigListEmpty = "Configuration list is empty.";
        internal const string MainAnnotationMissingDocBodyAttribute = "Main annotation entity missing documentbody attribute.";
        internal const string MainAnnotationBodyEmpty = "Main annotation documentbody is empty.";
        internal const string MainAnnotationBodyNotValidBase64 = "Main annotation documentbody is not valid Base64: {0}";
        internal const string FileFieldMissing = "File attribute '{0}' missing on source entity.";
        internal const string FileFieldEmptyOrErrorDownloading = "File missing or download failed for field '{0}'.";
        internal const string UnexpectedErrorDuringMerge = "Unexpected error during merge: {0}";
        internal const string EmptyPlaceholder = "Empty placeholder for field '{0}'.";
        internal const string PlaceholderNotFoundInMainDocument = "Placeholder '{0}' not found in main document.";
        internal const string PlaceholderNotFoundInMainDocumentBody = "Placeholder '{0}' not inside document body.";
        internal const string CoupleNotCompatibleError = "Object is not a compatible ICouple";
    }

    public enum MergeLogSeverity
    {
        Info,
        Warning,
        Error
    }
}