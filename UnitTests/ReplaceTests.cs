using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Novacode;

namespace UnitTests
{
    [TestClass]
    public class RegExTest
    {
        private readonly Dictionary<string, string> _testPatterns = new Dictionary<string, string>
        {
            { "COURT NAME", "Fred Frump" },
            { "Case Number", "cr-md-2011-1234567" }
        };
        private readonly TestHelper _testHelper;

        public RegExTest()
        {
            _testHelper = new TestHelper();
        }

        [TestMethod]
        public void ReplaceText_Can_ReplaceViaFunctionHandler()
        {
            using (var replaceDoc = DocX.Load(_testHelper.DirectoryWithFiles + "ReplaceTests.docx"))
            {
                foreach (var t in replaceDoc.Tables)
                {
                    // each table has 1 row and 3 columns
                    Assert.IsTrue(t.Rows[0].Cells.Count == 3);
                    Assert.IsTrue(t.ColumnCount == 3);
                    Assert.IsTrue(t.Rows.Count == 1);
                    Assert.IsTrue(t.RowCount == 1);
                }

                // Make sure the origional strings are in the document.
                Assert.IsTrue(replaceDoc.FindAll("<COURT NAME>").Count == 2);
                Assert.IsTrue(replaceDoc.FindAll("<Case Number>").Count == 2);

                // There are only two patterns, even though each pattern is used more than once
                Assert.IsTrue(replaceDoc.FindUniqueByPattern(@"<[\w \=]{4,}>", RegexOptions.IgnoreCase).Count == 2);

                // Make sure the new strings are not in the document.
                Assert.IsTrue(replaceDoc.FindAll("Fred Frump").Count == 0);
                Assert.IsTrue(replaceDoc.FindAll("cr-md-2011-1234567").Count == 0);

                // Do the replacing
                replaceDoc.ReplaceText("<(.*?)>", ReplaceTextHandler, false, RegexOptions.IgnoreCase);

                // Make sure the origional string are no longer in the document.
                Assert.IsTrue(replaceDoc.FindAll("<COURT NAME>").Count == 0);
                Assert.IsTrue(replaceDoc.FindAll("<Case Number>").Count == 0);

                // Make sure the new strings are now in the document.
                Assert.IsTrue(replaceDoc.FindAll("FRED FRUMP").Count == 2);
                Assert.IsTrue(replaceDoc.FindAll("cr-md-2011-1234567").Count == 2);

                // Make sure the replacement worked.
                Assert.IsTrue(replaceDoc.Text
                              == "\t\t\t\t\t\t\t\t\t\t\t\t\t\tThese two tables should look identical:\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tSTATE OF IOWA,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tPlaintiff,\t\t\t\t\t\t\t\t\t\t\t\t\t\tvs.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tFRED FRUMP,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tDefendant.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tCase No.: cr-md-2011-1234567\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tORDER SETTING ASIDE DEFAULT JUDGMENT\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tSTATE OF IOWA,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tPlaintiff,\t\t\t\t\t\t\t\t\t\t\t\t\t\tvs.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tFRED FRUMP,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tDefendant.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tCase No.: cr-md-2011-1234567\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tORDER SETTING ASIDE DEFAULT JUDGMENT\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
            }
        }

        private string ReplaceTextHandler(string findStr)
        {
            if (_testPatterns.ContainsKey(findStr))
            {
                return _testPatterns[findStr];
            }
            return findStr;
        }
    }
}