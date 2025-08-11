using Microsoft.Xrm.Sdk;
using System;

namespace WordMerge.Helpers
{
    public interface IFileDownloader
    {
        byte[] DownloadFile(Action<string> logger, EntityReference entityReference, string attributeName, out bool isExcel);
    }
}