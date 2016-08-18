using System;
using System.IO;

namespace UnitTests
{
    public static class TestHelper
    {
        static TestHelper()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            DirectoryWithFiles = Path.Combine(baseDirectory, "documents");
        }

        public static string DirectoryWithFiles { get; }
    }
}