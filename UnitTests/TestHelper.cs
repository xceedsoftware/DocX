namespace UnitTests
{
    public class TestHelper
    {
        public string DirectoryWithFiles { get; }

        public TestHelper()
        {
            var relativeDirectory = new RelativeDirectory(); // prepares the files for testing
            relativeDirectory.Up(3);
            DirectoryWithFiles = relativeDirectory.Path + @"\UnitTests\documents\";
        }
    }
}