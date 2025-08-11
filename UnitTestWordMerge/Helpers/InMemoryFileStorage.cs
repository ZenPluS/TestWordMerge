using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace UnitTestWordMerge.Helpers
{
    internal static class InMemoryFileStorage
    {
        private static readonly Dictionary<Guid, byte[]> FileStore = new  Dictionary<Guid, byte[]>();
        private static readonly Dictionary<Guid, string> FileStoreContext = new  Dictionary<Guid, string>();

        public static void AddFile(Guid id, byte[] content)
        {
            FileStore[id] = content;
        }

        public static void AddFileWithContext(Guid id, byte[] content, string fileName)
        {
            FileStore[id] = content;
            FileStoreContext[id] = fileName;
        }

        public static byte[] GetFile(Guid id)
        {
            return FileStore.TryGetValue(id, out var content) ? content : null;
        }

        public static byte[] GetFileWithContext(Guid id, out string fileName)
        {
            fileName = FileStoreContext.TryGetValue(id, out var name) ? name : null;
            return FileStore.TryGetValue(id, out var content) ? content : null;
        }
    }
}
