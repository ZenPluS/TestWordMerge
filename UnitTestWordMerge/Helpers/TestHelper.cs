using System;
using System.IO;
using System.Reflection;

namespace UnitTestWordMerge.Helpers
{
    internal static class TestHelper
    {
        internal static Stream GetEmbeddedResourceStream(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            return assembly.GetManifestResourceStream(resourceName);
        }

        internal static byte[] ReadFully(Stream input)
        {
            var buffer = new byte[16 * 1024];
            using (var ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

        internal static void CreateFile(string base64, string path)
            => File.WriteAllBytes(path, Convert.FromBase64String(base64));
    }
}
