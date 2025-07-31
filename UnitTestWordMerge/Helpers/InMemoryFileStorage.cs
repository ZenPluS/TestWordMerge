using System;
using System.Collections.Generic;

namespace UnitTestWordMerge.Helpers
{
    internal static class InMemoryFileStorage
    {
        private static readonly Dictionary<Guid, byte[]> FileStore = new  Dictionary<Guid, byte[]>();

        public static void AddFile(Guid id, byte[] content)
        {
            FileStore[id] = content;
        }

        public static byte[] GetFile(Guid id)
        {
            return FileStore.TryGetValue(id, out var content) ? content : null;
        }
    }
}
